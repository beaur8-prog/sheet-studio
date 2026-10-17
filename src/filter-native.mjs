import {address} from './engine.mjs';
import {validateAutoFilter} from './filter-schema.mjs';
import {dateNumber} from './filters.mjs';
export function nativeFilterPlan(f){
  if(!f)return null;validateAutoFilter(f);
  if(f.unsupported?.length)throw Error('지원하지 않는 필터를 변경하려면 먼저 지우기를 사용하세요.');
  const filters=Object.entries(f.columns).map(([c,rule])=>{
    const filter={field:+c-f.range.c1+1};
    if(rule.type==='values'){
      if(!rule.values.length)throw Error('필터 값이 비어 있습니다.');
      if(rule.kind==='date'){
        const dates=rule.values.filter(v=>v!==''&&Number.isFinite(Number(v))),other=rule.values.filter(v=>!dates.includes(v));
        if(dates.length)return {...filter,operator:7,criteria1:other.map(v=>v===''?'=':v),dateValues:dates.map(Number)};
      }
      return {...filter,operator:7,criteria1:rule.values.map(v=>v===''?'=':v)};
    }
    if(rule.type==='top')return {...filter,operator:rule.top?(rule.percent?5:3):(rule.percent?6:4),criteria1:String(rule.count)};
    if(rule.type==='average')return {...filter,operator:11,criteria1:rule.above?33:34};
    const conditions=rule.conditions.map(c=>{
      let v=rule.kind==='date'?String(dateNumber(c.value)):c.value,op=c.operator;
      if(['contains','notContains','beginsWith','endsWith'].includes(op)){
        v=op==='beginsWith'?v+'*':op==='endsWith'?'*'+v:'*'+v+'*';op=op==='notContains'?'notEqual':'equal';
      }
      return ({equal:'=',notEqual:'<>',greaterThan:'>',greaterThanOrEqual:'>=',lessThan:'<',lessThanOrEqual:'<='}[op])+v;
    });return {...filter,criteria1:conditions[0],...(conditions.length>1?{criteria2:conditions[1],operator:rule.join==='and'?1:2}:{})};
  });
  return {range:`${address(f.range.r1,f.range.c1)}:${address(f.range.r2,f.range.c2)}`,filters,hiddenRows:f.hiddenRows||[]};
}
