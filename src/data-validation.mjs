import {address,parseAddress,shiftFormula} from './engine.mjs';
import {xml,stringify,children,child,attr} from './chart-xml.mjs';
import {sheetIsProtected,validationRules} from './cell-constraints.mjs';
import {preservedPivotTouches} from './native-pivots.mjs';

const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const VALIDATION_TYPES={any:'모든 값',whole:'정수',decimal:'소수',list:'목록',date:'날짜',time:'시간',textLength:'텍스트 길이',custom:'사용자 수식'};
export const VALIDATION_OPERATORS={between:'사이',notBetween:'사이가 아님',equal:'같음',notEqual:'같지 않음',greaterThan:'보다 큼',lessThan:'보다 작음',greaterThanOrEqual:'크거나 같음',lessThanOrEqual:'작거나 같음'};
const intersects=(a,b)=>a.r1<=b.r2&&a.r2>=b.r1&&a.c1<=b.c2&&a.c2>=b.c1;
export function validationBounds(value){
  const text=String(value||'').trim().toUpperCase().replace(/\$/g,''),parts=text.split(':');
  const a=parseAddress(parts[0]),b=parseAddress(parts[1]||parts[0]);
  if(parts.length>2||!a||!b||a.r>b.r||a.c>b.c)throw Error('적용 범위를 A1 또는 A1:B10 형식으로 입력하세요.');
  return {r1:a.r,c1:a.c,r2:b.r,c2:b.c};
}
export const validationRange=b=>address(b.r1,b.c1)+(b.r1===b.r2&&b.c1===b.c2?'':':'+address(b.r2,b.c2));
const rebase=(formula,from,to)=>shiftFormula('='+formula.replace(/^=/,''),to.r-from.r,to.c-from.c).slice(1);
function subtract(a,b){
  if(!intersects(a,b))return [a];
  const r1=Math.max(a.r1,b.r1),r2=Math.min(a.r2,b.r2),c1=Math.max(a.c1,b.c1),c2=Math.min(a.c2,b.c2),out=[];
  if(a.r1<r1)out.push({...a,r2:r1-1});
  if(c1>a.c1)out.push({r1,r2,c1:a.c1,c2:c1-1});
  if(c2<a.c2)out.push({r1,r2,c1:c2+1,c2:a.c2});
  if(r2<a.r2)out.push({...a,r1:r2+1});
  return out;
}
export function assertValidationEditable(book,sheet,range){
  if(sheet.kind==='chart')throw Error('차트 전용 시트에는 입력 규칙을 만들 수 없습니다.');
  if(sheetIsProtected(sheet))throw Error('보호된 시트의 입력 규칙은 변경할 수 없습니다. Excel에서 시트 보호를 해제하세요.');
  if(book.nativeSource)throw Error('XLSM의 입력 규칙은 Excel에서 변경하거나 XLSX 사본을 사용하세요.');
  if(sheet.pivot)throw Error('피벗 결과에는 입력 규칙을 만들 수 없습니다. 일반 범위로 변환하거나 원본에 규칙을 설정하세요.');
  if(preservedPivotTouches(book,sheet.id,range))throw Error('원본 피벗의 결과에는 입력 규칙을 만들 수 없습니다. 일반 범위로 변환하거나 Excel에서 변경하세요.');
}
// Prefill relative formulas for the new selection, rather than moving their
// meaning when the dialog is opened inside an existing validation range.
export function validationAtSelection(sheet,range){
  const b=validationBounds(range),rules=validationRules(sheet),touching=rules.filter(rule=>rule.ranges.some(r=>intersects(r,b)));
  const rule=rules.find(rule=>rule.ranges.some(r=>b.r1>=r.r1&&b.r1<=r.r2&&b.c1>=r.c1&&b.c1<=r.c2));
  if(!rule)return {rule:null,mixed:touching.length>0};
  const whole=rule.ranges.some(r=>r.r1<=b.r1&&r.r2>=b.r2&&r.c1<=b.c1&&r.c2>=b.c2);
  return {rule:{...rule,formulae:rule.formulae.map(f=>rebase(f,rule.anchor||{r:0,c:0},{r:b.r1,c:b.c1}))},mixed:!whole||touching.length>1};
}
function operand(value,type){
  let text=String(value??'').trim();
  if(!text)throw Error('조건에 필요한 값 또는 수식을 입력하세요.');
  if(type==='date'&&/^\d{4}-\d\d-\d\d$/.test(text)){
    const [y,m,d]=text.split('-').map(Number),date=new Date(Date.UTC(y,m-1,d));
    if(y<1900||y>9999||date.getUTCFullYear()!==y||date.getUTCMonth()!==m-1||date.getUTCDate()!==d)throw Error('올바른 날짜를 YYYY-MM-DD 형식으로 입력하세요.');
    text=`DATE(${y},${m},${d})`;
  }else if(type==='time'&&/^\d{1,2}:\d\d(?::\d\d)?$/.test(text)){
    const [h,m,s=0]=text.split(':').map(Number);if(h>23||m>59||s>59)throw Error('시간은 00:00:00부터 23:59:59 사이로 입력하세요.');text=`TIME(${h},${m},${s})`;
  }
  if(text.length>255||/[\u0000-\u001f]/.test(text))throw Error('값 또는 수식은 줄바꿈 없이 255자 이내로 입력하세요.');
  return text.replace(/^=/,'');
}
export function createValidationRule(input){
  const type=input.type||'any';if(!Object.hasOwn(VALIDATION_TYPES,type))throw Error('지원하는 입력 규칙을 선택하세요.');
  const rule={type,operator:input.operator||'between',allowBlank:input.allowBlank!==false,showList:input.showList!==false,showError:input.showError!==false,severity:input.severity||'stop',showInput:!!input.showInput,promptTitle:String(input.promptTitle||''),prompt:String(input.prompt||''),title:String(input.title||''),message:String(input.message||''),formulae:[]};
  if(!['stop','warning','information'].includes(rule.severity))throw Error('오류 알림의 종류를 선택하세요.');
  for(const key of ['promptTitle','title'])if(rule[key].length>32)throw Error('안내·오류 제목은 32자 이내로 입력하세요.');
  for(const key of ['prompt','message'])if(rule[key].length>255)throw Error('안내·오류 내용은 255자 이내로 입력하세요.');
  if(type==='list'){
    const source=String(input.source??input.formulae?.[0]??'').trim();
    if(source.startsWith('='))rule.formulae=[operand(source,type)];
    else{
      const literal=/^"(?:[^"]|"")*"$/.test(source)?source.slice(1,-1).replace(/""/g,'"'):source;
      const list=literal.split(/[,\n\r]+/).map(v=>v.trim()).filter(Boolean).join(',');
      if(!list)throw Error('목록 항목을 쉼표로 구분하거나 =A1:A5 형식의 참조를 입력하세요.');
      if(list.length>255)throw Error('직접 입력한 목록은 255자까지입니다. 더 긴 목록은 셀 범위를 참조하세요.');
      rule.formulae=['"'+list.replace(/"/g,'""')+'"'];
    }
  }else if(type==='custom')rule.formulae=[operand(input.formulae?.[0],type)];
  else if(type!=='any'){
    if(!Object.hasOwn(VALIDATION_OPERATORS,rule.operator))throw Error('비교 조건을 선택하세요.');
    const count=rule.operator.includes('etween')?2:1;rule.formulae=Array.from({length:count},(_,i)=>operand(input.formulae?.[i],type));
    if(rule.formulae.every(f=>/^[-+]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(f))){
      const [a,b]=rule.formulae.map(Number);
      if(!Number.isFinite(a)||b!==undefined&&!Number.isFinite(b))throw Error('유한한 숫자를 입력하세요.');
      if(count===2&&a>b)throw Error('최솟값은 최댓값보다 클 수 없습니다.');
      if(type==='textLength'&&rule.formulae.some(f=>Number(f)<0||!Number.isInteger(Number(f))))throw Error('텍스트 길이는 0 이상의 정수로 입력하세요.');
    }
  }
  return rule;
}
function ruleNode(doc,b,rule){
  const node=doc.createElementNS(NS,'dataValidation');
  for(const [key,value] of Object.entries({type:rule.type,operator:!['list','custom','any'].includes(rule.type)?rule.operator:undefined,allowBlank:+rule.allowBlank,showDropDown:rule.type==='list'?+!rule.showList:undefined,showInputMessage:+rule.showInput,promptTitle:rule.promptTitle||undefined,prompt:rule.prompt||undefined,showErrorMessage:+rule.showError,errorStyle:rule.severity,errorTitle:rule.title||undefined,error:rule.message||undefined,sqref:validationRange(b)}))if(value!==undefined)node.setAttribute(key,String(value));
  rule.formulae.forEach((formula,i)=>{const child=doc.createElementNS(NS,'formula'+(i+1));child.appendChild(doc.createTextNode(formula));node.appendChild(child);});return node;
}
// Work only on a private XML tree. Existing cells and unrelated/opaque rules are
// untouched, and even an entire-column rule never allocates a cell per address.
export function planValidationChange(book,sheet,range,input){
  const b=validationBounds(range);assertValidationEditable(book,sheet,b);
  const rule=input===null?null:createValidationRule(input),source=sheet.xlsxFeatures?.dataValidations;
  const doc=xml(source||`<dataValidations xmlns="${NS}" count="0"/>`),root=doc.documentElement;
  for(const node of children(root,'dataValidation')){
    const refs=attr(node,'sqref').trim().split(/\s+/),ranges=refs.map(validationBounds),kept=ranges.flatMap(r=>subtract(r,b));
    if(kept.length===ranges.length&&kept.every((r,i)=>validationRange(r)===validationRange(ranges[i])))continue;
    if(!kept.length){root.removeChild(node);continue;}
    const before=parseAddress(refs[0].replace(/\$/g,'').split(':')[0]),after={r:kept[0].r1,c:kept[0].c1};
    for(const name of ['formula1','formula2']){const f=child(node,name);if(f)f.textContent=rebase(f.textContent,before,after);}
    node.setAttribute('sqref',kept.map(validationRange).join(' '));
  }
  if(rule&&(rule.type!=='any'||rule.showInput&&(rule.promptTitle||rule.prompt)))root.appendChild(ruleNode(doc,b,rule));
  const count=children(root,'dataValidation').length;root.setAttribute('count',String(count));
  if(count>10000||children(root,'dataValidation').some(node=>attr(node,'sqref').length>32767))throw Error('규칙이 너무 많은 영역으로 나뉩니다. 더 큰 범위를 한 번에 정리하세요.');
  const result=count?stringify(doc):undefined;if(result?.length>12_000_000)throw Error('입력 규칙의 크기 한도를 넘었습니다.');
  return {xml:result,range:validationRange(b),changed:result!==source};
}
export function applyValidationChange(sheet,plan){
  if(plan.xml){sheet.xlsxFeatures||={};sheet.xlsxFeatures.dataValidations=plan.xml;}
  else if(sheet.xlsxFeatures)delete sheet.xlsxFeatures.dataValidations;
}
