import {MAX_DOCUMENT_CELLS} from './limits.mjs';
const object=v=>v&&typeof v==='object'&&!Array.isArray(v);
export const filterOperators=['equal','notEqual','greaterThan','greaterThanOrEqual','lessThan','lessThanOrEqual','contains','notContains','beginsWith','endsWith'];
export function validateAutoFilter(f){
  const bad=()=>{throw Error('자동 필터 데이터가 올바르지 않습니다.');};
  if(!object(f)||!object(f.range)||!object(f.columns))bad();
  const {r1,r2,c1,c2}=f.range;
  if(![r1,r2,c1,c2].every(Number.isInteger)||r1<0||r2<r1||r2>=1048576||c1<0||c2<c1||c2>=16384||r2-r1>MAX_DOCUMENT_CELLS)bad();
  for(const [column,rule] of Object.entries(f.columns)){
    if(!/^\d+$/.test(column)||+column<c1||+column>c2||!object(rule))bad();
    if(rule.type==='values'){
      if(rule.kind!==undefined&&rule.kind!=='date')bad();
      if(!Array.isArray(rule.values)||rule.values.length>MAX_DOCUMENT_CELLS||rule.values.some(v=>typeof v!=='string'||v.length>32768))bad();
    }else if(rule.type==='custom'){
      if(!['text','number','date'].includes(rule.kind)||!['and','or'].includes(rule.join)||!Array.isArray(rule.conditions)||rule.conditions.length<1||rule.conditions.length>2)bad();
      if(rule.conditions.some(c=>!object(c)||!filterOperators.includes(c.operator)||typeof c.value!=='string'||c.value.length>32768))bad();
    }else if(rule.type==='top'){
      if(typeof rule.top!=='boolean'||typeof rule.percent!=='boolean'||!Number.isFinite(rule.count)||rule.count<1||rule.count>500||rule.percent&&rule.count>100)bad();
    }else if(rule.type==='average'){
      if(typeof rule.above!=='boolean')bad();
    }else if(rule.type==='color'){
      if(!['background','color'].includes(rule.target)||rule.color!==null&&(typeof rule.color!=='string'||!/^#[\da-f]{6}$/i.test(rule.color)))bad();
    }else bad();
  }
  if(f.hiddenRows!==undefined&&(!Array.isArray(f.hiddenRows)||f.hiddenRows.length>MAX_DOCUMENT_CELLS||f.hiddenRows.some(r=>!Number.isInteger(r)||r<=r1||r>r2)))bad();
  if(f.unsupported!==undefined&&(!Array.isArray(f.unsupported)||f.unsupported.length>16384||f.unsupported.some(c=>!Number.isInteger(c)||c<c1||c>c2)))bad();
  if(f.sort&&(!Number.isInteger(f.sort.column)||f.sort.column<c1||f.sort.column>c2||typeof f.sort.desc!=='boolean'))bad();
  return f;
}
