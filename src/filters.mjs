import {cloneData} from './clone.mjs';
import {MAX_DOCUMENT_CELLS} from './limits.mjs';
import {address,parseAddress,createEvaluator,formatValue,shiftFormula} from './engine.mjs';
import {validateAutoFilter} from './filter-schema.mjs';
import {createConditionalEvaluator} from './conditional.mjs';
import {normalizeColor} from './spreadsheet-colors.mjs';

export const filterActive=s=>!!(Object.keys(s.autoFilter?.columns||{}).length||s.autoFilter?.unsupported?.length);
export const valueKey=value=>String(value??'').toLocaleLowerCase();
export function filterValue(sheet,evaluate,r,c){
  const cell=sheet.cells[address(r,c)];
  return !cell?.raw||cell.raw==="'"?'':evaluate(address(r,c));
}
export function inferFilterRange(sheet,selection){
  const points=Object.entries(sheet.cells).filter(([,v])=>v.raw!=='').map(([key])=>parseAddress(key));
  let b={...selection};
  if(b.r1===b.r2&&b.c1===b.c2){
    if(!points.some(p=>p.r===b.r1&&p.c===b.c1))throw Error('필터를 적용할 데이터 안의 셀을 선택하세요.');
    let changed=true;
    while(changed){changed=false;for(const p of points){if(p.r>=b.r1-1&&p.r<=b.r2+1&&p.c>=b.c1-1&&p.c<=b.c2+1){
      const next={r1:Math.min(b.r1,p.r),r2:Math.max(b.r2,p.r),c1:Math.min(b.c1,p.c),c2:Math.max(b.c2,p.c)};
      if(JSON.stringify(next)!==JSON.stringify(b)){b=next;changed=true;}
    }}}
  }else if(b.r1===b.r2){const occupied=new Set(points.filter(p=>p.c>=b.c1&&p.c<=b.c2).map(p=>p.r));while(occupied.has(b.r2+1))b.r2++;}else if(b.r2===1048575){b.r2=points.filter(p=>p.c>=b.c1&&p.c<=b.c2).reduce((max,p)=>Math.max(max,p.r),b.r1);}
  if(!points.some(p=>p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2))throw Error('선택한 범위에 데이터가 없습니다.');
  validateAutoFilter({range:b,columns:{}});return b;
}
function pattern(value){
  let result='';for(let i=0;i<value.length;i++){
    let c=value[i];if(c==='~'&&i+1<value.length)c=value[++i];
    else if(c==='*'){result+='.*';continue;}else if(c==='?'){result+='.';continue;}
    result+=c.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  }return new RegExp('^'+result+'$','iu');
}
export const searchFilterValue=(label,query)=>pattern('*'+query+'*').test(label);
export function filterDateParts(serial){
  if(Math.floor(serial)===60)return {year:1900,month:2,day:29};
  const d=new Date(Date.UTC(1899,11,31)+(Math.floor(serial)-(serial>=60?1:0))*86400000);
  return {year:d.getUTCFullYear(),month:d.getUTCMonth()+1,day:d.getUTCDate()};
}
export function dateNumber(text){
  if(/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)){
    const [y,m,d]=text.split('-').map(Number),t=Date.UTC(y,m-1,d),date=new Date(t);
    if(date.getUTCFullYear()!==y||date.getUTCMonth()!==m-1||date.getUTCDate()!==d)return NaN;
    return (t-Date.UTC(1899,11,31))/86400000+(t>=Date.UTC(1900,2,1)?1:0);
  }return text.trim()===''?NaN:Number(text);
}
function matchesCondition(value,condition,kind){
  const {operator:op,value:operand}=condition;
  if(['contains','notContains','beginsWith','endsWith'].includes(op)){
    if(typeof value!=='string')return op==='notContains';
    const p=op==='beginsWith'?operand+'*':op==='endsWith'?'*'+operand:'*'+operand+'*';
    const match=pattern(p).test(value);return op==='notContains'?!match:match;
  }
  if(kind==='text'){
    const a=String(value??''),b=operand;
    if(op==='equal'||op==='notEqual'){const match=pattern(b).test(a);return op==='equal'?match:!match;}
    const cmp=a.localeCompare(b,undefined,{sensitivity:'base'});
    return {greaterThan:cmp>0,greaterThanOrEqual:cmp>=0,lessThan:cmp<0,lessThanOrEqual:cmp<=0}[op];
  }
  const n=kind==='date'?dateNumber(operand):Number(operand);
  if(typeof value!=='number'||!Number.isFinite(n)||operand.trim()==='')return op==='notEqual'&&value!=='';
  return {equal:value===n,notEqual:value!==n,greaterThan:value>n,greaterThanOrEqual:value>=n,lessThan:value<n,lessThanOrEqual:value<=n}[op];
}
function predicate(rule,values){
  if(rule.type==='values'){const set=new Set(rule.values.map(valueKey));return value=>set.has(valueKey(rule.kind==='date'&&typeof value==='number'?Math.floor(value):value));}
  if(rule.type==='custom')return value=>rule.conditions[rule.join==='and'?'every':'some'](c=>matchesCondition(value,c,rule.kind));
  const numbers=values.filter(v=>typeof v==='number'&&Number.isFinite(v));
  if(!numbers.length)return ()=>false;
  if(rule.type==='average'){const mean=numbers.reduce((a,b)=>a+b,0)/numbers.length;return v=>typeof v==='number'&&(rule.above?v>mean:v<mean);}
  numbers.sort((a,b)=>rule.top?b-a:a-b);
  const count=rule.percent?Math.max(1,Math.floor(numbers.length*rule.count/100)):Math.trunc(rule.count),limit=numbers[Math.min(count,numbers.length)-1];
  return v=>typeof v==='number'&&(rule.top?v>=limit:v<=limit);
}
export function filterPredicates(book,sheet,evaluate,omitColumn=-1){
  const f=sheet.autoFilter;if(!f)return [];
  const conditional=Object.values(f.columns).some(r=>r.type==='color')?createConditionalEvaluator(book,sheet,evaluate):null;
  return Object.entries(f.columns).filter(([c])=>+c!==omitColumn).map(([c,rule])=>{
    if(rule.type==='color')return {column:+c,test:(_value,r)=>normalizeColor({...sheet.cells[address(r,+c)]?.style,...conditional(r,+c,evaluate(address(r,+c)))}[rule.target])===normalizeColor(rule.color)};
    const values=[];if(rule.type==='top'||rule.type==='average')for(let r=f.range.r1+1;r<=f.range.r2;r++)values.push(filterValue(sheet,evaluate,r,+c));
    return {column:+c,test:predicate(rule,values)};
  });
}
export function applyFilters(book,sheet){
  const f=sheet.autoFilter;if(!f)return;
  validateAutoFilter(f);if(f.unsupported?.length)throw Error('이 문서의 미지원 아이콘·색 또는 동적 날짜 필터는 Excel에서 편집하거나 지우기를 먼저 사용하세요.');const evaluate=createEvaluator(book,sheet.id),rules=filterPredicates(book,sheet,evaluate),hidden=[];
  for(let r=f.range.r1+1;r<=f.range.r2;r++)if(!rules.every(({column,test})=>test(filterValue(sheet,evaluate,r,column),r)))hidden.push(r);
  f.hiddenRows=hidden;
}
export function filterOptions(book,sheet,column){
  const evaluate=createEvaluator(book,sheet.id),rules=filterPredicates(book,sheet,evaluate,column),values=new Map(),conditional=createConditionalEvaluator(book,sheet,evaluate),colors={background:new Map(),color:new Map()};let dates=0,numbers=0,texts=0;
  const b=sheet.autoFilter.range;
  for(let r=b.r1+1;r<=b.r2;r++){
    if(!rules.every(({column,test})=>test(filterValue(sheet,evaluate,r,column),r)))continue;
    const value=filterValue(sheet,evaluate,r,column),key=valueKey(value),style=sheet.cells[address(r,column)]?.style;
    const displayed={...style,...conditional(r,column,evaluate(address(r,column)))};for(const target of ['background','color']){const color=normalizeColor(displayed[target]);colors[target].set(color,(colors[target].get(color)||0)+1);}
    if(!values.has(key))values.set(key,{key,value,label:value===''?'(빈 셀)':formatValue(value,style)});
    if(value!==''){if(typeof value==='number'){if(style?.format==='date')dates++;else numbers++;}else texts++;}
  }
  const kind=dates>numbers&&dates>texts?'date':numbers>texts?'number':'text';
  const colorOptions=Object.fromEntries(Object.entries(colors).map(([target,list])=>[target,[...list].map(([color,count])=>({color,count})).sort((a,b)=>(a.color||'').localeCompare(b.color||''))]));
  if(kind==='date'){
    const grouped=new Map();for(const v of values.values()){
      if(typeof v.value==='number'){v.value=Math.floor(v.value);v.key=valueKey(v.value);const d=filterDateParts(v.value);v.label=`${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;}
      grouped.set(v.key,v);
    }return {kind,colors:colorOptions,values:[...grouped.values()].sort((a,b)=>compareFilterValues(a.value,b.value,false))};
  }
  return {kind,colors:colorOptions,values:[...values.values()].sort((a,b)=>compareFilterValues(a.value,b.value,false))};
}
export function compareFilterValues(a,b,desc){
  if(a==='')return b===''?0:1;if(b==='')return -1;
  const rank=v=>typeof v==='number'?0:typeof v==='string'?1:2;
  const cmp=rank(a)!==rank(b)?rank(a)-rank(b):typeof a==='number'?a-b:String(a).localeCompare(String(b),'ko',{sensitivity:'base'});
  return desc?-cmp:cmp;
}
export function sortFilterRange(book,sheet,column,desc){
  const f=sheet.autoFilter,b=f.range,evaluate=createEvaluator(book,sheet.id),rows=[],destinations=[],hidden=new Set(f.hiddenRows||[]);
  if(f.unsupported?.length)throw Error('지원하지 않는 필터를 지운 뒤 정렬하세요.');
  if((b.r2-b.r1)*(b.c2-b.c1+1)>MAX_DOCUMENT_CELLS)throw Error('한 번에 1,000,000개 이하의 셀을 정렬하세요.');
  if(sheet.merges?.some(m=>{const [a,z]=m.split(':').map(parseAddress);return a.r<=b.r2&&z.r>b.r1&&a.c<=b.c2&&z.c>=b.c1;}))throw Error('병합된 셀이 있는 범위는 정렬할 수 없습니다.');
  for(let r=b.r1+1;r<=b.r2;r++){
    if(hidden.has(r))continue;
    destinations.push(r);
    const cells=[];for(let c=b.c1;c<=b.c2;c++)cells.push(cloneData(sheet.cells[address(r,c)]));
    rows.push({r,cells,value:filterValue(sheet,evaluate,r,column)});
  }
  rows.sort((a,b)=>compareFilterValues(a.value,b.value,desc));
  rows.forEach((row,i)=>row.cells.forEach((cell,c)=>{
    const r=destinations[i],key=address(r,b.c1+c);
    if(!cell){delete sheet.cells[key];return;}
    if(cell.raw.startsWith('='))cell.raw=shiftFormula(cell.raw,r-row.r,0);
    sheet.cells[key]=cell;
  }));f.sort={column,desc};
}

export function sortRange(book,sheet,b,column,desc){
  if(!b||![b.r1,b.r2,b.c1,b.c2].every(Number.isInteger)||b.r1<0||b.c1<0||b.r2>=1048576||b.c2>=16384||b.r2<b.r1||b.c2<b.c1||(b.r2-b.r1+1)*(b.c2-b.c1+1)>MAX_DOCUMENT_CELLS)throw Error('정렬 범위는 최대 1,000,000셀입니다.');
  if(sheet.merges?.some(m=>{const [a,z]=m.split(':').map(parseAddress);return a.r<=b.r2&&z.r>=b.r1&&a.c<=b.c2&&z.c>=b.c1;}))throw Error('병합 셀이 포함된 범위는 정렬할 수 없습니다.');
  const evaluate=createEvaluator(book,sheet.id),rows=[];
  for(let r=b.r1;r<=b.r2;r++){
    const cells=[];for(let c=b.c1;c<=b.c2;c++)cells.push(sheet.cells[address(r,c)]);
    rows.push({r,cells,value:filterValue(sheet,evaluate,r,column)});
  }
  rows.sort((a,b)=>compareFilterValues(a.value,b.value,desc)||a.r-b.r);
  rows.forEach((row,i)=>row.cells.forEach((source,c)=>{
    const ref=address(b.r1+i,b.c1+c);if(!source){delete sheet.cells[ref];return;}
    sheet.cells[ref]=source.raw?.startsWith('=')?{...source,raw:shiftFormula(source.raw,b.r1+i-row.r,0)}:source;
  }));
}
export function shiftAutoFilter(sheet,axis,index,delta){
  const f=sheet.autoFilter;if(!f)return;const b=f.range,lo=axis==='r'?'r1':'c1',hi=axis==='r'?'r2':'c2';
  if(delta<0&&index===b.r1&&axis==='r'){delete sheet.autoFilter;return;}
  if(index<b[lo]||delta>0&&index===b[lo]){b[lo]+=delta;b[hi]+=delta;}else if(index<=b[hi])b[hi]+=delta;
  b[hi]=Math.min(b[hi],axis==='r'?1048575:16383);
  if(b[hi]<b[lo]||b.r1===b.r2){delete sheet.autoFilter;return;}
  if(axis==='c'){
    f.columns=Object.fromEntries(Object.entries(f.columns).filter(([c])=>!(delta<0&&+c===index)).map(([c,v])=>[+c>=index?+c+delta:+c,v]));
    f.columns=Object.fromEntries(Object.entries(f.columns).filter(([c])=>+c<=b.c2));
    if(f.unsupported)f.unsupported=f.unsupported.filter(c=>!(delta<0&&c===index)).map(c=>c>=index?c+delta:c).filter(c=>c<=b.c2);
    if(f.sort){if(delta<0&&f.sort.column===index)delete f.sort;else if(f.sort.column>=index)f.sort.column+=delta;}
    if(f.sort?.column>b.c2)delete f.sort;
  }else f.hiddenRows=(f.hiddenRows||[]).filter(r=>!(delta<0&&r===index)).map(r=>r>=index?r+delta:r).filter(r=>r<=b.r2);
}
