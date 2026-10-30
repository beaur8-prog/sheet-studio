import {cloneData} from './clone.mjs';
import {MAX_DOCUMENT_CELLS} from './limits.mjs';
import {address,parseAddress,createEvaluator,transformFormulaReferences} from './engine.mjs';
import {aggregations,scalarKey,pivotLabel,validatePivot} from './pivot-schema.mjs';
export {aggregations,scalarKey,pivotLabel};
const LIMIT = 100000;
const fail = message => {throw new Error(message);};

export function sourceData(book, sourceSheetId, sourceRange, {evaluate}={}) {
  const source = book.sheets.find(s=>s.id===sourceSheetId);
  if (!source) fail('원본 시트가 없습니다. 원본을 다시 선택하세요.');
  if (source.pivot) fail('원본에는 일반 데이터 시트를 선택하세요.');
  const parts=String(sourceRange).toUpperCase().replace(/\$/g,'').split(':');
  const a=parseAddress(parts[0]),b=parseAddress(parts[1]);
  if(parts.length!==2||!a||!b||a.r>=b.r||a.c>b.c) fail('머리글과 데이터 행을 포함하는 범위를 입력하세요. 예: A1:D100');
  const width=b.c-a.c+1,height=b.r-a.r+1;
  if(width*height>LIMIT) fail('피벗 원본은 최대 100,000셀까지 지원합니다.');
  evaluate||=createEvaluator(book,source.id);
  const read=(r,c)=>{const ref=address(r,c),cell=source.cells[ref];if(!cell?.raw&&!evaluate.isComputed?.(ref))return '';const result=evaluate.result?.(ref);if(result?.error)fail(`${ref}에 계산 오류가 있습니다: ${result.error}`);return result?result.value:evaluate(ref);};
  const headers=Array.from({length:width},(_,i)=>String(read(a.r,a.c+i)).trim());
  if(headers.some(h=>!h||h.length>255)||new Set(headers.map(h=>h.toLowerCase())).size!==headers.length) fail('첫 행에는 비어 있지 않은 고유한 머리글이 필요합니다 (최대 255자).');
  const records=[];
  for(let r=a.r+1;r<=b.r;r++){
    const row=Array.from({length:width},(_,c)=>read(r,a.c+c));
    if(row.every(v=>v===''||v===null))continue;
    records.push(row);
  }
  return {headers,records,source,range:`${address(a.r,a.c)}:${address(b.r,b.c)}`};
}

function state(){return {count:0,n:0,sum:0,min:Infinity,max:-Infinity};}
function add(s,value){if(value!==''&&value!==null&&value!==undefined)s.count++;if(typeof value==='number'&&Number.isFinite(value)){s.n++;s.sum+=value;s.min=Math.min(s.min,value);s.max=Math.max(s.max,value);}}
function result(s,kind){if(!s)s=state();switch(kind){case 'sum':return s.sum;case 'count':return s.count;case 'countNumbers':return s.n;case 'average':return s.n?s.sum/s.n:'';case 'min':return s.n?s.min:'';case 'max':return s.n?s.max:'';}}
const compare=(a,b)=>{for(let i=0;i<a.length;i++){const x=a[i],y=b[i];const d=typeof x==='number'&&typeof y==='number'?x-y:pivotLabel(x).localeCompare(pivotLabel(y),'ko',{numeric:true});if(d)return d;}return 0;};

export function buildPivot(book,spec,options){
  validatePivot(spec);
  const data=sourceData(book,spec.sourceSheetId,spec.sourceRange,options),{headers,records}=data;
  const index=name=>{const i=headers.indexOf(name);if(i<0)fail(`원본의 “${name}” 필드를 찾을 수 없습니다. 필드 설정을 확인하세요.`);return i;};
  const ri=spec.rows.map(index),ci=spec.columns.map(index),vi=spec.values.map(v=>index(v.field));
  const filters=spec.filters.map(f=>({index:index(f.field),values:new Set(f.values.map(scalarKey))}));
  const rowGroups=new Map(),columnGroups=new Map(),buckets=new Map(),rowTotals=new Map(),columnTotals=new Map(),grand=vi.map(state);
  const update=(map,key,record)=>{let bucket=map.get(key);if(!bucket)map.set(key,bucket=vi.map(state));vi.forEach((col,i)=>add(bucket[i],record[col]));};
  let matched=0;
  for(const record of records){
    if(filters.some(f=>!f.values.has(scalarKey(record[f.index]))))continue;
    matched++;
    const rv=ri.map(i=>record[i]),cv=ci.map(i=>record[i]),rk=JSON.stringify(rv),ck=JSON.stringify(cv);
    rowGroups.set(rk,rv);columnGroups.set(ck,cv);
    if((rowGroups.size+2)*(columnGroups.size+(spec.totals&&ci.length?1:0))*vi.length>LIMIT)fail('피벗 결과가 너무 큽니다. 필드나 원본 범위를 줄이세요.');
    update(buckets,JSON.stringify([rk,ck]),record);update(rowTotals,rk,record);update(columnTotals,ck,record);vi.forEach((col,i)=>add(grand[i],record[col]));
  }
  if(!ri.length)rowGroups.set('[]',[]);
  if(!ci.length)columnGroups.set('[]',[]);
  const order=groups=>{const values=[...groups];return spec.sort==='source'?values:values.sort((a,b)=>(spec.sort==='desc'?-1:1)*compare(a[1],b[1]));};
  const rows=order(rowGroups),columns=order(columnGroups),labels=spec.values.map(v=>`${aggregations[v.aggregate]}: ${v.field}`);
  const matrix=[spec.rows.length?[...spec.rows]:['행 레이블']];
  for(const [,cv] of columns)matrix[0].push(...labels.map(label=>ci.length?`${cv.map(pivotLabel).join(' / ')} · ${label}`:label));
  if(spec.totals&&ci.length)matrix[0].push(...labels.map(label=>`총합계 · ${label}`));
  for(const [rk,rv] of rows){
    const row=rv.length?rv.map(pivotLabel):['전체'];
    for(const [ck] of columns){const bucket=buckets.get(JSON.stringify([rk,ck]));row.push(...spec.values.map((v,i)=>result(bucket?.[i],v.aggregate)));}
    if(spec.totals&&ci.length)row.push(...spec.values.map((v,i)=>result(rowTotals.get(rk)?.[i],v.aggregate)));
    matrix.push(row);
  }
  if(spec.totals&&ri.length){
    const row=['총합계',...Array(Math.max(0,ri.length-1)).fill('')];
    for(const [ck] of columns)row.push(...spec.values.map((v,i)=>result(columnTotals.get(ck)?.[i],v.aggregate)));
    if(ci.length)row.push(...spec.values.map((v,i)=>result(grand[i],v.aggregate)));
    matrix.push(row);
  }
  if(matrix.length*matrix[0].length>LIMIT||matrix[0].length>16384)fail('피벗 결과가 너무 큽니다. 필드나 원본 범위를 줄이세요.');
  if(matrix.some(row=>row.some(v=>typeof v==='number'&&!Number.isFinite(v))))fail('집계 결과가 숫자 범위를 초과했습니다.');
  return {matrix,matched,total:records.length,data,rows,columns};
}

export function pivotCells(output,spec){
  const cells={},labelWidth=Math.max(1,spec.rows.length);
  output.matrix.forEach((row,r)=>row.forEach((value,c)=>{
    const heading=r===0,total=spec.totals&&spec.rows.length>0&&r===output.matrix.length-1;
    cells[address(r,c)]={raw:typeof value==='string'?(value?"'"+value:''):String(value),style:{background:heading?'#107c41':total?'#deeee5':r%2?'#f2f7f4':'#ffffff',color:heading?'#ffffff':'#24382c',bold:heading||total,format:typeof value==='number'?'number':'general',decimals:2,wrap:heading,align:c<labelWidth?'left':'right'}};
  }));
  return cells;
}

// Calculate before mutation: a bad source must not destroy the last good report.
export function applyPivot(book,target,spec,output=buildPivot(book,spec)){
  const cells=pivotCells(output,spec);
  const other=book.sheets.filter(s=>s.id!==target.id).reduce((n,s)=>n+Object.keys(s.cells).length,0);
  if(other+Object.keys(cells).length>MAX_DOCUMENT_CELLS)fail('통합 문서의 1,000,000셀 제한을 초과합니다.');
  target.cells=cells;target.merges=[];target.rowHeights={0:44};target.columnWidths=Object.fromEntries(output.matrix[0].map((_,i)=>[i,i<Math.max(1,spec.rows.length)?145:190]));
  target.pivot=cloneData(spec);delete target.pivotError;
  return output;
}

export function shiftPivotSources(book,sourceId,axis,index,delta){
  const source=book.sheets.find(s=>s.id===sourceId);
  for(const sheet of book.sheets)if(sheet.pivot?.sourceSheetId===sourceId){
    const next=transformFormulaReferences(sheet.pivot.sourceRange,{axis,index,delta,targetSheet:source.name,ownerSheet:source.name});
    sheet.pivot.sourceRange=next.includes('#REF!')?'#REF!':next;
  }
}
