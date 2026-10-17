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
export function chartSeries(chart,sheet,workbook){
  if(chart.series?.length)return chart.series;
  const [a,b]=String(chart.range).split(':').map(parseAddress);if(!a||!b)return [];
  const ref=(c,r1=a.r+1,r2=b.r)=>chartRef(sheet.id,address(r1,c)+':'+address(r2,c));
  const result=[];
  for(let c=a.c+1;c<=b.c&&result.length<32;c+=chart.type==='bubble'?2:1){
    const i=result.length;result.push({nameRef:ref(c,a.r,a.r),catRef:ref(a.c),valRef:ref(c),...(chart.type==='scatter'||chart.type==='bubble'?{xRef:ref(a.c)}:{}),...(chart.type==='bubble'?{sizeRef:c+1<=b.c?ref(c+1):{cache:[]}}:{}),color:chartColors[i%chartColors.length],type:chart.type==='combo'?(c===b.c?'line':'bar'):chart.type,secondary:chart.type==='combo'&&c===b.c&&chart.secondaryAxis!==false});
  }
  return result;
}
export function chartData(chart,sheet,workbook){
  const preview=ref=>refValues(ref,workbook,{limit:1000});
  return chartSeries(chart,sheet,workbook).map((s,i)=>({
    ...s,name:String(refValues(s.nameRef,workbook)[0]??s.name??`계열 ${i+1}`),
    labels:preview(s.catRef).map(v=>v===null?'':String(v)),
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
