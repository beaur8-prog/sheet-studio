import {address,parseAddress,createEvaluator,transformFormulaReferences} from './engine.mjs';

export const chartTypes={bar:'세로 막대형','bar-horizontal':'가로 막대형',line:'꺾은선형',area:'영역형',pie:'원형',doughnut:'도넛형',scatter:'분산형',bubble:'거품형',radar:'방사형',combo:'혼합 (막대·꺾은선)'};
export const chartColors=['#4472c4','#ed7d31','#70ad47','#ffc000','#5b9bd5','#a5a5a5','#8064a2','#264478'];
export const chartRef=(sheetId,range)=>({sheetId,range});
export function parseChartFormula(formula,workbook){
  const m=/^(?:'((?:[^']|'')+)'|([^!\[\]]+))!\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i.exec(formula||'');
  if(!m){const named=workbook.chartNames?.find(n=>n.name.toLowerCase()===String(formula).toLowerCase());return named?parseChartFormula(named.formula,{...workbook,chartNames:[]}):null;}
  const sheet=workbook.sheets.find(s=>s.name.toLowerCase()===(m[1]?.replace(/''/g,"'")||m[2]).toLowerCase());
  if(!sheet)return null;
  const start=m[3].toUpperCase()+m[4],end=(m[5]||m[3]).toUpperCase()+(m[6]||m[4]);
  if(!parseAddress(start)||!parseAddress(end))return null;
  return {sheetId:sheet.id,range:start+':'+end};
}
export function refFormula(ref,workbook,names,offsets){
  const sheet=workbook.sheets.find(s=>s.id===ref?.sheetId);if(!sheet||!ref.range||ref.range.includes('#REF!'))return ref?.sheetId?'#REF!':ref?.formula||'#REF!';
  const name=names?.get(sheet.id)||sheet.name,offset=offsets?.get(sheet.id)||0;
  const range=ref.range.split(':').map(a=>{const p=parseAddress(a);return p?address(p.r+offset,p.c).replace(/^([A-Z]+)(\d+)$/,'$$$1$$$2'):'#REF!';}).join(':');
  return "'"+name.replace(/'/g,"''")+"'!"+range;
}
export function refValues(ref,workbook,{limit=10000}={}){
  if(!ref)return [];
  const sheet=workbook.sheets.find(s=>s.id===ref.sheetId),parts=ref.range?.split(':').map(parseAddress);
  if(!sheet||!parts?.[0]||!parts.at(-1))return (ref.cache||[]).slice(0,limit);
  const [a,b=a]=parts,evaluate=createEvaluator(workbook,sheet.id),out=[];
  for(let r=Math.min(a.r,b.r);r<=Math.max(a.r,b.r);r++)for(let c=Math.min(a.c,b.c);c<=Math.max(a.c,b.c);c++){
    if(out.length>=limit)return out;
    const cell=sheet.cells[address(r,c)];let v=evaluate(address(r,c));
    if(typeof v==='string'&&v.startsWith('#')&&cell?.cachedValue!==undefined)v=cell.cachedValue;
    out.push(cell?.raw===''||!cell?null:v);
  }
  return out;
}
export function referenceSize(ref){const [a,b]=ref?.range?.split(':').map(parseAddress)||[];return a&&b?(Math.abs(a.r-b.r)+1)*(Math.abs(a.c-b.c)+1):ref?.cache?.length||0;}
// A horizontal vector is one series, just like a vertical vector. Persist the
// chosen direction on new charts; imported explicit series keep their own refs.
function sourceBounds(chart){
  const parts=String(chart.range).split(':'),[a,b]=parts.map(parseAddress);
  return parts.length===2&&a&&b&&b.r>=a.r&&b.c>=a.c?{a,b}:null;
}
export function chartSeriesBy(chart,sheet,workbook){
  if(['rows','columns'].includes(chart.seriesBy))return chart.seriesBy;
  const bounds=sourceBounds(chart);if(!bounds)return 'columns';
  const {a,b}=bounds;if(a.r===b.r)return 'rows';
  // A header row followed by one observation describes named categories,
  // rather than unrelated one-point series with a redundant legend.
  if(sheet&&workbook&&b.r===a.r+1&&b.c>a.c&&b.c-a.c<=10000&&!['scatter','bubble'].includes(chart.type)){
    const evaluate=createEvaluator(workbook,sheet.id),number=v=>typeof v==='number'&&Number.isFinite(v),start=a.c+Number(!number(evaluate(address(b.r,a.c))));
    let named=b.c-start>=1;for(let c=start;c<=b.c&&named;c++){const heading=evaluate(address(a.r,c));named=typeof heading==='string'&&heading.trim()!==''&&!heading.startsWith('#')&&number(evaluate(address(b.r,c)));}
    if(named)return 'rows';
  }
  return 'columns';
}
function sourceAxes(chart,sheet,workbook){
  const bounds=sourceBounds(chart);if(!bounds)return null;
  const {a,b}=bounds,rows=chartSeriesBy(chart,sheet,workbook)==='rows';
  const lane1=rows?a.r:a.c,lane2=rows?b.r:b.c,point1=rows?a.c:a.r,point2=rows?b.c:b.r;
  if(lane2-lane1>32||point2-point1>10000)return null;
  return {rows,lane1,lane2,point1,point2,xy:['scatter','bubble'].includes(chart.type),implicitX:chart.type==='scatter'&&lane1===lane2};
}
export function chartLayout(chart,sheet,workbook){
  const axes=sourceAxes(chart,sheet,workbook);if(!axes)return {firstRowAsHeader:false,firstColumnAsLabels:false};
  const {rows,lane1,lane2,point1,point2,xy,implicitX}=axes,evaluate=createEvaluator(workbook,sheet.id),numeric=v=>typeof v==='number'&&Number.isFinite(v);
  const value=(lane,point)=>{const key=rows?address(lane,point):address(point,lane),cell=sheet.cells[key];if(!cell||cell.raw==='')return null;const v=evaluate(key);return typeof v==='string'&&v.startsWith('#')&&cell.cachedValue!==undefined?cell.cachedValue:v;};
  let numericFirstLane=false;for(let point=point1;point<=point2;point++)if(numeric(value(lane1,point))){numericFirstLane=true;break;}
  const categories=(rows?chart.firstRowAsHeader:chart.firstColumnAsLabels)??(lane1<lane2&&!numericFirstLane);
  const firstPoint=[];for(let lane=lane1+Number(xy?!implicitX:categories);lane<=lane2;lane++)firstPoint.push(value(lane,point1));
  const names=(rows?chart.firstColumnAsLabels:chart.firstRowAsHeader)??(firstPoint.some(v=>typeof v==='string'&&v.trim()!==''&&!v.startsWith('#'))&&!firstPoint.some(numeric));
  return rows?{firstRowAsHeader:categories,firstColumnAsLabels:names}:{firstRowAsHeader:names,firstColumnAsLabels:categories};
}
function seriesShape(chart,sheet,workbook){
  const axes=sourceAxes(chart,sheet,workbook);if(!axes)return null;
  const layout=chartLayout(chart,sheet,workbook),{rows,xy,implicitX,lane1,point1}=axes;
  const names=rows?layout.firstColumnAsLabels:layout.firstRowAsHeader,categories=rows?layout.firstRowAsHeader:layout.firstColumnAsLabels;
  return {...axes,names,categories,startLane:lane1+Number(xy?!implicitX:categories),startPoint:point1+Number(names)};
}
export function chartRangeError(chart,sheet,workbook){
  if(!sourceBounds(chart))return '올바른 셀 범위를 입력하세요. 예: A1:C5';
  const limit='차트에는 계열당 최대 10,000개 값과 32개 값 행·열을 사용하세요.',shape=seriesShape(chart,sheet,workbook);
  if(!shape)return limit;
  const {startLane,lane2,startPoint,point2,rows}=shape,points=point2-startPoint+1,lanes=lane2-startLane+1;
  if(chart.type==='bubble'&&(lanes<2||lanes%2!==0))return `거품형에는 X값·Y값·크기 ${rows?'행':'열'}이 필요합니다. X값 다음에 Y값·크기 쌍을 선택하세요.`;
  if(points<1||lanes<1)return '값으로 사용할 행과 열이 필요합니다. 첫 행·첫 열 사용 설정을 확인하세요.';
  if(points>10000||lanes>32)return limit;
  return null;
}
export function chartSeries(chart,sheet,workbook){
  if(chart.series?.length)return chart.series;
  if(chartRangeError(chart,sheet,workbook))return [];
  const {rows,lane1,lane2,point1,point2,startLane,startPoint,names,categories,xy,implicitX}=seriesShape(chart,sheet,workbook);
  const key=(lane,point)=>rows?address(lane,point):address(point,lane),ref=(lane,from=startPoint,to=point2)=>chartRef(sheet.id,key(lane,from)+':'+key(lane,to));
  const evaluate=createEvaluator(workbook,sheet.id);let populated=false;for(let lane=startLane;lane<=lane2&&!populated;lane++)for(let point=startPoint;point<=point2;point++){const cell=sheet.cells[key(lane,point)];if(!cell||cell.raw==='')continue;let value=evaluate(key(lane,point));if(typeof value==='string'&&value.startsWith('#')&&cell.cachedValue!==undefined)value=cell.cachedValue;if(typeof value!=='boolean'&&value!==null&&value!==''&&Number.isFinite(Number(value))){populated=true;break;}}
  if(!populated)return [];
  const indexed={cache:Array.from({length:point2-startPoint+1},(_,i)=>xy?i+1:String(i+1))},categoryRef=xy&&!implicitX||!xy&&categories?ref(lane1):indexed,result=[];
  for(let lane=startLane;lane<=lane2;lane+=chart.type==='bubble'?2:1){
    const i=result.length,isLine=chart.type==='combo'&&startLane<lane2&&lane===lane2;
    result.push({...(names?{nameRef:ref(lane,point1,point1)}:{name:`계열 ${i+1}`}),catRef:categoryRef,valRef:ref(lane),...(xy?{xRef:categoryRef}:{}),...(chart.type==='bubble'?{sizeRef:ref(lane+1)}:{}),color:chartColors[i%chartColors.length],type:chart.type==='combo'?(isLine?'line':'bar'):chart.type,secondary:isLine&&chart.secondaryAxis!==false});
  }
  return result;
}
export function chartData(chart,sheet,workbook){
  const preview=ref=>refValues(ref,workbook,{limit:1000});
  return chartSeries(chart,sheet,workbook).map((s,i)=>({
    ...s,name:String(refValues(s.nameRef,workbook)[0]??s.name??`계열 ${i+1}`),
    labels:preview(s.catRef||s.xRef).map(v=>v===null?'':String(v)),
    values:preview(s.valRef).map(v=>v===null||v===''||typeof v==='boolean'||!Number.isFinite(Number(v))?null:Number(v)),
    xs:preview(s.xRef).map(v=>v===null?null:Number(v)),
    sizes:preview(s.sizeRef).map(v=>Math.max(0,Number(v)||0)),
    color:/^#[0-9a-f]{6}$/i.test(s.color||'')?s.color:chartColors[i%chartColors.length],
    type:s.type||chart.type,
  }));
}
export function updateChartReferences(workbook,sheetId,axis,index,delta){
  const target=workbook.sheets.find(s=>s.id===sheetId);if(!target)return;
  for(const named of workbook.chartNames||[])named.formula=transformFormulaReferences(named.formula,{axis,index,delta,targetSheet:target.name,ownerSheet:target.name});
  for(const owner of workbook.sheets)for(const chart of owner.charts||[]){
    if(owner.id===sheetId&&!chart.source)chart.range=transformFormulaReferences(chart.range,{axis,index,delta,targetSheet:target.name,ownerSheet:target.name});
    for(const series of chart.series||[])for(const key of ['nameRef','catRef','valRef','xRef','sizeRef']){const ref=series[key];if(ref?.sheetId===sheetId)ref.range=transformFormulaReferences(ref.range,{axis,index,delta,targetSheet:target.name,ownerSheet:target.name});}
    for(const ref of chart.source?.formulas||[])if(ref?.sheetId===sheetId)ref.range=transformFormulaReferences(ref.range,{axis,index,delta,targetSheet:target.name,ownerSheet:target.name});
  }
}
