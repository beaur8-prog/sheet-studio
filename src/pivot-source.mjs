import {cloneData} from './clone.mjs';
import {tableBounds} from './table-schema.mjs';
import {address} from './engine.mjs';

const fields=spec=>[...new Set([...spec.rows,...spec.columns,...spec.values.map(v=>v.field),...spec.filters.map(f=>f.field)])];
export function pivotTableRange(table){
  if(!table?.headerRow)throw Error('피벗 원본 표에는 머리글 행이 필요합니다.');
  const b=tableBounds(table.range);if(!b||b.r2-(table.totalsRow?1:0)<=b.r1)throw Error('피벗 원본 표에 데이터 행이 없습니다.');
  return address(b.r1,b.c1)+':'+address(b.r2-(table.totalsRow?1:0),b.c2);
}
function tableFor(book,spec,id){
  const source=book.sheets.find(s=>s.id===spec.sourceSheetId),table=source?.tables?.find(t=>t.id===id);
  if(!table)throw Error('연결된 원본 표가 없습니다. 이전 결과를 유지했습니다. 피벗 설정에서 원본을 다시 선택하세요.');
  const ids=table.columns.map(c=>c.id);if(ids.some(id=>!Number.isInteger(id)||id<1)||new Set(ids).size!==ids.length)throw Error('원본 표의 열 식별자가 중복되거나 없습니다. Excel 또는 표 관리에서 구성을 확인하세요.');
  return table;
}
// Bind only fields used by the report. This stays small even for very wide
// tables and prevents a deleted/re-created column with the same name taking over.
export function bindPivotTable(book,spec,id){
  const table=tableFor(book,spec,id),next=cloneData(spec);
  next.sourceTableId=table.id;next.sourceTableName=table.name;next.sourceRange=pivotTableRange(table);
  next.sourceFields=fields(next).map(name=>{const column=table.columns.find(c=>c.name===name);if(!column)throw Error(`원본 표에서 “${name}” 열을 찾을 수 없습니다.`);return {id:column.id,name:column.name};});
  return next;
}
export function unlinkPivotTable(spec){const next=cloneData(spec);for(const key of ['sourceTableId','sourceTableName','sourceFields'])delete next[key];return next;}
export function resolvePivotSpec(book,spec){
  if(!spec.sourceTableId)return cloneData(spec);
  const table=tableFor(book,spec,spec.sourceTableId),mapping=new Map();
  for(const field of fields(spec)){
    const saved=spec.sourceFields?.find(c=>c.name===field),column=saved&&table.columns.find(c=>c.id===saved.id);
    if(!column)throw Error(`피벗에서 사용하는 “${field}” 열이 삭제되었거나 바뀌었습니다. 이전 결과를 유지했습니다. 원본 필드를 다시 선택하세요.`);
    mapping.set(field,column.name);
  }
  const next=cloneData(spec);next.rows=next.rows.map(f=>mapping.get(f));next.columns=next.columns.map(f=>mapping.get(f));next.values=next.values.map(v=>({...v,field:mapping.get(v.field)}));next.filters=next.filters.map(f=>({...f,field:mapping.get(f.field)}));
  return bindPivotTable(book,next,table.id);
}

// A verified app-owned XLSX has new in-memory table IDs after import. Rebind by
// its saved native table name only at that boundary; runtime refresh never falls
// back to a same-named replacement table.
export function restorePivotTableBinding(book,spec){
  if(!spec.sourceTableId)return spec;
  const source=book.sheets.find(s=>s.id===spec.sourceSheetId),table=source?.tables?.find(t=>t.name.toLocaleLowerCase()===spec.sourceTableName.toLocaleLowerCase());
  if(!table)throw Error('파일의 피벗 원본 표를 찾을 수 없습니다. 원본을 다시 선택하세요.');
  const rebound={...cloneData(spec),sourceTableId:table.id};return resolvePivotSpec(book,rebound);
}
