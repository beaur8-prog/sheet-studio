import {address,parseAddress} from './engine.mjs';
import {autoSumPlan} from './formula-tools.mjs';

export const mobileAggregateNames={SUM:'합계',AVERAGE:'평균',COUNT:'숫자 개수',MIN:'최솟값',MAX:'최댓값'};
export function mobileAggregateFormula(name,range){
  if(!Object.hasOwn(mobileAggregateNames,name))throw Error('지원하는 집계 함수를 선택하세요.');
  const parts=String(range).trim().toUpperCase().replaceAll('$','').split(':'),[a,b=a]=parts.map(parseAddress);
  if(parts.length>2||!a||!b)throw Error('계산 범위를 셀 주소로 입력하세요. 예: A1:A10');
  const first=address(Math.min(a.r,b.r),Math.min(a.c,b.c)),last=address(Math.max(a.r,b.r),Math.max(a.c,b.c));
  return `=${name}(${first===last?first:first+':'+last})`;
}
export function mobileAggregateEntry(name,range,target){
  const raw=mobileAggregateFormula(name,range),ref=String(target).trim().toUpperCase().replaceAll('$',''),point=parseAddress(ref);
  if(!point)throw Error('결과 셀 주소를 입력하세요. 예: A11');
  const [a,b=a]=raw.slice(raw.indexOf('(')+1,-1).split(':').map(parseAddress);
  if(point.r>=a.r&&point.r<=b.r&&point.c>=a.c&&point.c<=b.c)throw Error('순환 참조를 피하려면 결과 셀을 계산 범위 밖으로 지정하세요.');
  return {raw,ref,point};
}

// Suggest, but never write, a result. A selected rectangle means all selected
// values; the dialog makes the range and destination explicit before insertion.
export function mobileAggregateSuggestion(sheet,selection,evaluate){
  const r1=Math.min(selection.r,selection.r2),r2=Math.max(selection.r,selection.r2),c1=Math.min(selection.c,selection.c2),c2=Math.max(selection.c,selection.c2);
  const occupied=(r,c)=>!!sheet.cells[address(r,c)]?.raw||!!evaluate.isComputed?.(address(r,c));
  if(r1!==r2||c1!==c2){
    const range=address(r1,c1)+':'+address(r2,c2),target=address(r1===r2?r1:r2+1,r1===r2?c2+1:c1);
    return {range,target};
  }
  let target={r:r1,c:c1};
  if(occupied(r1,c1)){
    const numeric=(r,c)=>r>=0&&c>=0&&typeof evaluate(address(r,c))==='number';
    target=numeric(r1,c1-1)&&!numeric(r1-1,c1)?{r:r1,c:c1+1}:{r:r1+1,c:c1};
  }
  if(!parseAddress(address(target.r,target.c)))return {range:address(r1,c1),target:''};
  if(occupied(target.r,target.c))return {range:address(r1,c1),target:address(target.r,target.c)};
  const first=autoSumPlan(sheet,{...target,r2:target.r,c2:target.c},evaluate)[0];
  return {range:/^=SUM\((.+)\)$/.exec(first.raw)?.[1]||address(r1,c1),target:address(first.r,first.c)};
}
