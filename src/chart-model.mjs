import {worksheetCellStyle} from './cell-style.mjs';
import {filteredRowsForSheet} from './table-schema.mjs';
import {address,parseAddress,createEvaluator,transformFormulaReferences,formatValue,numberFormatKind,formulaReference} from './engine.mjs';

export const chartTypes={bar:'세로 막대형','bar-horizontal':'가로 막대형',line:'꺾은선형',area:'영역형',pie:'원형',doughnut:'도넛형',scatter:'분산형',bubble:'거품형',radar:'방사형',combo:'혼합 (막대·꺾은선)'};
export const chartColors=['#4472c4','#ed7d31','#70ad47','#ffc000','#5b9bd5','#a5a5a5','#8064a2','#264478'];
export const chartRef=(sheetId,range)=>({sheetId,range});
export function parseChartFormula(formula,workbook,ownerSheetId=workbook.activeSheetId){
  const m=/^(?:'((?:[^']|'')+)'|([^!\[\]]+))!\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i.exec(formula||'');
  if(!m){const ref=formulaReference(workbook,ownerSheetId,formula||'');return ref?{sheetId:ref.sheetId,range:address(ref.range.r1,ref.range.c1)+':'+address(ref.range.r2,ref.range.c2)}:null;}
  const sheet=workbook.sheets.find(s=>s.name.toLowerCase()===(m[1]?.replace(/''/g,"'")||m[2]).toLowerCase());
  if(!sheet)return null;
  const start=m[3].toUpperCase()+m[4],end=(m[5]||m[3]).toUpperCase()+(m[6]||m[4]);
  if(!parseAddress(start)||!parseAddress(end))return null;
  return {sheetId:sheet.id,range:start+':'+end};
}
const dynamicReference=formula=>formula&&!/!\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/i.test(formula);
function resolveChartRef(ref,workbook){if(!ref||!dynamicReference(ref.formula))return ref;const resolved=parseChartFormula(ref.formula,workbook,ref.ownerSheetId||ref.sheetId||workbook.activeSheetId);return resolved?{...ref,...resolved}:{...ref,sheetId:undefined,range:undefined,...(workbook.referenceCachesValid===false?{cache:[]}: {})};}
export function refFormula(ref,workbook,names,offsets){
  // Resolving a defined name for the preview must not replace the original
  // named reference when saving: changes to that name should still reach Excel.
  if(dynamicReference(ref?.formula))return ref.formula;
  const sheet=workbook.sheets.find(s=>s.id===ref?.sheetId);if(!sheet||!ref.range||ref.range.includes('#REF!'))return ref?.sheetId?'#REF!':ref?.formula||'#REF!';
  const name=names?.get(sheet.id)||sheet.name,offset=offsets?.get(sheet.id)||0;
  const range=ref.range.split(':').map(a=>{const p=parseAddress(a);return p?address(p.r+offset,p.c).replace(/^([A-Z]+)(\d+)$/,'$$$1$$$2'):'#REF!';}).join(':');
  return "'"+name.replace(/'/g,"''")+"'!"+range;
}
function chartCellValue(cell,evaluate,key,onCached,workbook){
  const result=evaluate.result?.(key),value=result?result.value:evaluate(key);
  // A calculation error is current data, not a reason to resurrect an old
  // successful result. Only an unchanged, unsupported formula may use a cache.
  if((result?result.error==='#NAME?':value==='#NAME?')&&workbook?.referenceCachesValid!==false&&cell?.cachedValue!==undefined&&(!cell.cachedFormula||cell.cachedFormula===cell.raw)){onCached?.();return cell.cachedValue;}
  return value;
}
export function refValues(ref,workbook,{limit=10000,onCached}={}){
  if(!ref)return [];ref=resolveChartRef(ref,workbook);
  const sheet=workbook.sheets.find(s=>s.id===ref.sheetId),parts=ref.range?.split(':').map(parseAddress);
  if(!sheet||!parts?.[0]||!parts.at(-1))return (ref.cache||[]).slice(0,limit);
  const [a,b=a]=parts,evaluate=createEvaluator(workbook,sheet.id),out=[];
  for(let r=Math.min(a.r,b.r);r<=Math.max(a.r,b.r);r++)for(let c=Math.min(a.c,b.c);c<=Math.max(a.c,b.c);c++){
    if(out.length>=limit)return out;
    const key=address(r,c),cell=sheet.cells[key],v=chartCellValue(cell,evaluate,key,onCached,workbook);
    out.push((cell?.raw===''||!cell)&&v===''?null:v);
  }
  return out;
}
function refPositions(ref,workbook,limit=1000){
  ref=resolveChartRef(ref,workbook);const sheet=workbook.sheets.find(s=>s.id===ref?.sheetId),parts=ref?.range?.split(':').map(parseAddress);
  if(!sheet||!parts?.[0]||!parts.at(-1))return [];
  const [a,b=a]=parts,out=[];
  for(let r=Math.min(a.r,b.r);r<=Math.max(a.r,b.r)&&out.length<limit;r++)for(let c=Math.min(a.c,b.c);c<=Math.max(a.c,b.c)&&out.length<limit;c++)out.push({sheet,r,c});
  return out;
}
function refFormat(ref,workbook){
  for(const {sheet,r,c} of refPositions(ref,workbook,16)){
    const style=worksheetCellStyle(workbook,sheet,r,c);
    if(style?.numberFormat||style?.format&&style.format!=='general')return style;
  }
  return ref?.formatCode?{numberFormat:ref.formatCode}:{};
}
export function chartNumberFormat(ref,workbook){
  const style=refFormat(ref,workbook);
  if(style.numberFormat)return style.numberFormat;
  const decimals=Number.isInteger(style.decimals)?Math.max(0,Math.min(12,style.decimals)):2;
  const fraction=decimals?'.'+'0'.repeat(decimals):'';
  return ({number:'#,##0'+fraction,currency:'₩#,##0'+fraction,accounting:'₩#,##0'+fraction,percent:'0'+fraction+'%',date:'yyyy-mm-dd'})[style.format]||'General';
}
function visiblePoints(ref,workbook){
  const positions=refPositions(ref,workbook),hidden=new Map();
  return positions.map(({sheet,r,c})=>{
    if(!hidden.has(sheet))hidden.set(sheet,filteredRowsForSheet(sheet));
    return sheet.rowHeights?.[r]!==0&&sheet.columnWidths?.[c]!==0&&!hidden.get(sheet).has(r);
  });
}
function filterChartData(data,chart,workbook){
  if(chart.plotVisibleOnly===false)return data;
  const masks=data.map(s=>{
    const refs=[s.valRef,...(['scatter','bubble'].includes(chart.type)?[s.xRef,s.sizeRef]:[])].filter(Boolean).map(ref=>visiblePoints(ref,workbook));
    return s.values.map((_,i)=>refs.every(mask=>mask[i]!==false));
  });
  const included=data.map((s,i)=>({s,mask:masks[i]})).filter(({mask})=>mask.some(Boolean));
  const indices=Array.from({length:Math.max(0,...data.map(s=>s.values.length))},(_,i)=>i).filter(i=>included.some(({mask})=>mask[i]));
  return included.map(({s,mask})=>{
    const keep=['scatter','bubble'].includes(chart.type)?indices.filter(i=>mask[i]):indices;
    const pointColors=Object.fromEntries(keep.flatMap((old,i)=>s.pointColors?.[old]?[[i,s.pointColors[old]]]:[]));
    return {...s,values:keep.map(i=>mask[i]?s.values[i]??null:null),labels:keep.map(i=>s.labels[i]??String(i+1)),categories:keep.map(i=>s.categories[i]),xs:keep.map(i=>s.xs[i]),sizes:keep.map(i=>s.sizes[i]),pointColors};
  });
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
  const value=(lane,point)=>{const key=rows?address(lane,point):address(point,lane),cell=sheet.cells[key];return !cell||cell.raw===''?null:chartCellValue(cell,evaluate,key,undefined,workbook);};
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
  if(chart.series?.length)return chart.series.map(series=>{const copy={...series};for(const key of ['nameRef','catRef','valRef','xRef','sizeRef'])if(copy[key])copy[key]=resolveChartRef(copy[key],workbook);return copy;});
  if(chartRangeError(chart,sheet,workbook))return [];
  const {rows,lane1,lane2,point1,point2,startLane,startPoint,names,categories,xy,implicitX}=seriesShape(chart,sheet,workbook);
  const key=(lane,point)=>rows?address(lane,point):address(point,lane),ref=(lane,from=startPoint,to=point2)=>chartRef(sheet.id,key(lane,from)+':'+key(lane,to));
  const evaluate=createEvaluator(workbook,sheet.id);let populated=false;for(let lane=startLane;lane<=lane2&&!populated;lane++)for(let point=startPoint;point<=point2;point++){const cell=sheet.cells[key(lane,point)];if((!cell||cell.raw==='')&&!evaluate.isComputed?.(key(lane,point)))continue;const value=chartCellValue(cell,evaluate,key(lane,point),undefined,workbook);if(typeof value!=='boolean'&&value!==null&&value!==''&&Number.isFinite(Number(value))){populated=true;break;}}
  if(!populated)return [];
  const indexed={cache:Array.from({length:point2-startPoint+1},(_,i)=>xy?i+1:String(i+1))},categoryRef=xy&&!implicitX||!xy&&categories?ref(lane1):indexed,result=[];
  for(let lane=startLane;lane<=lane2;lane+=chart.type==='bubble'?2:1){
    const i=result.length,isLine=chart.type==='combo'&&startLane<lane2&&lane===lane2;
    result.push({...(names?{nameRef:ref(lane,point1,point1)}:{name:`계열 ${i+1}`}),catRef:categoryRef,valRef:ref(lane),...(xy?{xRef:categoryRef}:{}),...(chart.type==='bubble'?{sizeRef:ref(lane+1)}:{}),color:chartColors[i%chartColors.length],type:chart.type==='combo'?(isLine?'line':'bar'):chart.type,secondary:isLine&&chart.secondaryAxis!==false});
  }
  return result;
}
function arrangeDateCategories(data,chart){
  if(!['bar','bar-horizontal','line','area','combo'].includes(chart.type))return data;
  const source=data.find(s=>s.dateCategories);
  if(!source||source.categories.length<2||!source.categories.every(v=>typeof v==='number'&&Number.isFinite(v)))return data;
  const order=source.categories.map((_,i)=>i).sort((a,b)=>source.categories[a]-source.categories[b]);
  return data.map(s=>({...s,...Object.fromEntries(['values','labels','categories','xs','sizes'].map(key=>[key,order.map(i=>s[key][i])])),pointColors:Object.fromEntries(order.flatMap((old,i)=>s.pointColors?.[old]?[[i,s.pointColors[old]]]:[]))}));
}
export function chartData(chart,sheet,workbook){
  const data=chartSeries(chart,sheet,workbook).map((s,i)=>{
    let usesCachedValues=false;
    const preview=ref=>refValues(ref,workbook,{limit:1000,onCached:()=>{usesCachedValues=true;}});
    const categoryFormat=refFormat(s.catRef||s.xRef,workbook);
    return {
    ...s,name:String(refValues(s.nameRef,workbook)[0]??s.name??`계열 ${i+1}`),
    categories:preview(s.catRef||s.xRef),dateCategories:categoryFormat.format==='date'||numberFormatKind(categoryFormat.numberFormat||'General')==='date',
    labels:preview(s.catRef||s.xRef).map(v=>v===null?'':formatValue(v,categoryFormat,workbook)),
    values:preview(s.valRef).map(v=>v===null||v===''||typeof v==='boolean'||!Number.isFinite(Number(v))?null:Number(v)),
    xs:preview(s.xRef).map(v=>v===null?null:Number(v)),
    sizes:preview(s.sizeRef).map(v=>Math.max(0,Number(v)||0)),
    color:/^#[0-9a-f]{6}$/i.test(s.color||'')?s.color:chartColors[i%chartColors.length],
    type:s.type||chart.type,
    valueFormat:refFormat(s.valRef,workbook),xFormat:refFormat(s.xRef,workbook),
    truncated:referenceSize(s.valRef)>1000,
    usesCachedValues,
  };});
  return arrangeDateCategories(filterChartData(data,chart,workbook),chart);
}
export function updateChartReferences(workbook,sheetId,axis,index,delta){
  const target=workbook.sheets.find(s=>s.id===sheetId);if(!target)return;
  for(const named of workbook.chartNames||[])named.formula=transformFormulaReferences(named.formula,{axis,index,delta,targetSheet:target.name,ownerSheet:workbook.sheets.find(s=>s.id===named.sheetId)?.name||target.name});
  for(const owner of workbook.sheets)for(const chart of owner.charts||[]){
    if(owner.id===sheetId&&!chart.source)chart.range=transformFormulaReferences(chart.range,{axis,index,delta,targetSheet:target.name,ownerSheet:target.name});
    for(const series of chart.series||[])for(const key of ['nameRef','catRef','valRef','xRef','sizeRef']){const ref=series[key];if(ref?.sheetId===sheetId)ref.range=transformFormulaReferences(ref.range,{axis,index,delta,targetSheet:target.name,ownerSheet:target.name});if(ref?.formula)ref.formula=transformFormulaReferences(ref.formula,{axis,index,delta,targetSheet:target.name,ownerSheet:workbook.sheets.find(s=>s.id===ref.ownerSheetId)?.name||owner.name});}
    for(const ref of chart.source?.formulas||[]){if(ref?.sheetId===sheetId)ref.range=transformFormulaReferences(ref.range,{axis,index,delta,targetSheet:target.name,ownerSheet:target.name});if(ref?.formula)ref.formula=transformFormulaReferences(ref.formula,{axis,index,delta,targetSheet:target.name,ownerSheet:workbook.sheets.find(s=>s.id===ref.ownerSheetId)?.name||owner.name});}
  }
}
