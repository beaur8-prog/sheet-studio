import {unwrap} from './clone.mjs';
import {address,parseAddress,evaluateCell,evaluateFormula,evaluateFormulaArray,shiftFormula} from './engine.mjs';
import {xml,all,attr} from './chart-xml.mjs';

const cache=new Map(),inside=(b,p)=>p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2;
function parsed(source,read){if(!source)return null;if(cache.has(source))return cache.get(source);const value=read(xml(source));if(cache.size>=40)cache.delete(cache.keys().next().value);cache.set(source,value);return value;}
function range(text){const [first,last=first]=String(text).split(':'),a=parseAddress(first),b=parseAddress(last);return a&&b?{r1:Math.min(a.r,b.r),r2:Math.max(a.r,b.r),c1:Math.min(a.c,b.c),c2:Math.max(a.c,b.c)}:null;}
export function protection(sheet){return parsed(sheet.xlsxFeatures?.sheetProtection,doc=>Object.fromEntries(Array.from(doc.documentElement.attributes).map(a=>[a.name,a.value])));}
export const sheetIsProtected=sheet=>['1','true'].includes(protection(sheet)?.sheet);
export function protectionAllows(sheet,operation){if(sheet.kind==='chart')return operation==='objects'&&!parsed(sheet.xlsxChartSheet?.xml,doc=>all(doc,'sheetProtection').length);if(!sheetIsProtected(sheet))return true;const p=protection(sheet);return ['0','false'].includes(p[operation]);}
export function cellProtection(book,sheet,ref){
  const cell=sheet.cells[ref],p=parseAddress(ref),style=book.xlsxStyles?.[cell?.xlsxStyle]?.style;
  return cell?.xlsxProtection||style?.protection||sheet.xlsxProtectionDefaults?.rows?.[p?.r]||sheet.xlsxProtectionDefaults?.columns?.[p?.c]||{};
}
export const cellIsLocked=(book,sheet,ref)=>sheet.kind==='chart'||sheetIsProtected(sheet)&&cellProtection(book,sheet,ref).locked!==false;
export const cellFormulaHidden=(book,sheet,ref)=>sheetIsProtected(sheet)&&cellProtection(book,sheet,ref).hidden===true;
let arrayIndexes=new WeakMap();
export function resetConstraintIndexes(){arrayIndexes=new WeakMap();}
export function refreshConstraintIndexes(book,dirty){for(const sheet of book.sheets){const cells=unwrap(sheet.cells),index=arrayIndexes.get(cells),refs=dirty.get(cells);if(!index||!refs)continue;for(const ref of refs){if(cells[ref]?.xlsxFormula?.shareType==='array')index.add(ref);else index.delete(ref);}}}
export function arrayRegions(sheet){const cells=unwrap(sheet.cells);let index=arrayIndexes.get(cells);if(!index){index=new Set();for(const ref in cells)if(cells[ref]?.xlsxFormula?.shareType==='array')index.add(ref);arrayIndexes.set(cells,index);}const arrays=[];for(const ref of index){const cell=cells[ref];if(cell?.raw?.startsWith('=')&&cell.xlsxFormula?.shareType==='array'){const b=range(cell.xlsxFormula.ref||ref);if(b)arrays.push({ref,...b});}}return arrays;}

export function cellEditError(book,sheet,ref){
  if(sheet.kind==='chart')return '차트 전용 시트에서는 셀을 편집할 수 없습니다.';
  if(cellIsLocked(book,sheet,ref))return '보호된 셀입니다. 입력이 허용된 셀을 선택하거나 Excel에서 시트 보호를 해제하세요.';
  const p=parseAddress(ref),array=arrayRegions(sheet).find(a=>inside(a,p));if(array&&book.nativeSource)return 'XLSM의 배열 수식은 Excel에서 편집하거나 XLSX 사본을 사용하세요.';if(array&&array.ref!==ref)return '배열 수식 결과의 일부를 변경할 수 없습니다. 배열의 첫 셀에서 수식을 편집하세요.';
  return null;
}
export function assertCellUpdates(book,sheet,updates,{editAnchor=false,formatOnly=false}={}){
  if(formatOnly){if(!protectionAllows(sheet,'formatCells'))throw Error('이 시트에서는 셀 서식을 변경할 수 없습니다.');return;}
  const keys=new Set(updates.map(u=>u.key));
  for(const key of keys)if(cellIsLocked(book,sheet,key))throw Error('선택 범위에 보호된 셀이 있습니다. 변경을 적용하지 않았습니다.');
  for(const a of arrayRegions(sheet)){
    let count=0;for(const key of keys)if(inside(a,parseAddress(key)))count++;
    const anchor=updates.find(u=>u.key===a.ref),formula=anchor?.cell?.raw;
    if(count&&book.nativeSource)throw Error('XLSM의 배열 수식은 Excel에서 편집하거나 XLSX 사본을 사용하세요.');
    if(count&&count!==(a.r2-a.r1+1)*(a.c2-a.c1+1)&&!(editAnchor&&count===1&&keys.has(a.ref)&&(formula===undefined||formula.startsWith('='))))throw Error('배열 수식 결과의 일부를 변경할 수 없습니다. 배열 전체를 선택하세요.');
  }
  if(book.nativeSource&&updates.some(u=>u.cell?.xlsxFormula?.shareType==='array'))throw Error('XLSM에 배열 수식을 붙여넣는 작업은 Excel에서 진행하세요.');
}
export function assertStructureAllowed(sheet,axis,index,count,delta){
  if(!protectionAllows(sheet,(delta>0?'insert':'delete')+(axis==='r'?'Rows':'Columns')))throw Error('보호된 시트에서 이 행·열 변경은 허용되지 않습니다.');
  for(const a of arrayRegions(sheet)){const lo=axis==='r'?a.r1:a.c1,hi=axis==='r'?a.r2:a.c2;
    if(delta>0?index>lo&&index<=hi:index<=hi&&index+count-1>=lo&&!(index<=lo&&index+count-1>=hi))throw Error('배열 수식의 일부에 행·열을 삽입하거나 삭제할 수 없습니다.');
  }
}
export function assertSortAllowed(book,sheet,b){
  if(!protectionAllows(sheet,'sort'))throw Error('보호된 시트에서 정렬은 허용되지 않습니다.');
  if(arrayRegions(sheet).some(a=>a.r1<=b.r2&&a.r2>=b.r1&&a.c1<=b.c2&&a.c2>=b.c1))throw Error('배열 수식이 포함된 범위는 정렬할 수 없습니다.');
  if(sheetIsProtected(sheet))for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++)if(cellIsLocked(book,sheet,address(r,c)))throw Error('정렬 범위에 보호된 셀이 있습니다.');
}
export function validationRules(sheet){return parsed(sheet.xlsxFeatures?.dataValidations,doc=>all(doc,'dataValidation').map(node=>({
  ranges:attr(node,'sqref').split(/\s+/).map(range).filter(Boolean),anchor:parseAddress(attr(node,'sqref').split(/\s+/)[0]?.split(':')[0]),
  type:attr(node,'type','any'),operator:attr(node,'operator','between'),allowBlank:['1','true'].includes(attr(node,'allowBlank')),showError:['1','true'].includes(attr(node,'showErrorMessage')),showList:!['1','true'].includes(attr(node,'showDropDown')),
  severity:attr(node,'errorStyle','stop'),title:attr(node,'errorTitle','입력 제한'),message:attr(node,'error','입력한 값이 이 셀의 입력 제한 조건과 일치하지 않습니다.'),
  formulae:['formula1','formula2'].map(name=>all(node,name)[0]?.textContent).filter(v=>v!==undefined),
})))||[];}
const forCell=(sheet,ref)=>{const point=parseAddress(ref);return validationRules(sheet).filter(rule=>rule.ranges.some(b=>inside(b,point)));};
function relativeFormula(rule,formula,position){return shiftFormula('='+String(formula).replace(/^=/,''),position.r-(rule.anchor?.r||0),position.c-(rule.anchor?.c||0));}
function choices(book,sheet,ref,rule){
  const expression=rule.formulae[0]||'',literal=/^"(?:[^"]|"")*"$/.test(expression);
  if(literal)return expression.slice(1,-1).replace(/""/g,'"').split(',');
  const result=evaluateFormulaArray(book,sheet.id,relativeFormula(rule,expression,parseAddress(ref)),{position:parseAddress(ref)});
  if(result.error||result.values.some(v=>typeof v==='string'&&/^#(?:REF!|NAME\?|VALUE!|NUM!)/.test(v)))return null;
  return result.values.filter(v=>v!=='');
}
export function validationChoices(book,sheet,ref){const rule=forCell(sheet,ref).find(r=>r.type==='list'&&r.showList);if(!rule)return null;return {values:choices(book,sheet,ref,rule),rule};}
export function validateCellUpdates(book,sheet,updates){
  if(!sheet.xlsxFeatures?.dataValidations)return null;
  // A prototype can be a history Proxy: assigning through it would invoke its
  // setter and mutate the live workbook before validation has succeeded.
  const overrides=new Map(updates.map(update=>[update.key,update.cell]));
  const cells=new Proxy(sheet.cells,{get:(source,key)=>overrides.has(key)?overrides.get(key):Reflect.get(source,key),ownKeys:source=>[...new Set([...Reflect.ownKeys(source),...overrides.keys()])],getOwnPropertyDescriptor:(source,key)=>overrides.has(key)?{value:overrides.get(key),enumerable:true,configurable:true,writable:true}:Reflect.getOwnPropertyDescriptor(source,key)});
  const snapshot={...book,sheets:book.sheets.map(s=>s.id===sheet.id?{...s,cells}:s)},target=snapshot.sheets.find(s=>s.id===sheet.id);
  for(const update of updates){
    const ref=update.key,position=parseAddress(ref),value=evaluateCell(snapshot,sheet.id,ref);
    for(const rule of forCell(sheet,ref)){
      if(rule.type==='any'||!rule.showError||value===''&&rule.allowBlank)continue;
      let valid=true,unsupported=false;
      if(rule.type==='list'){const list=choices(snapshot,target,ref,rule);if(list)valid=list.some(v=>String(v).toLocaleLowerCase()===String(value).toLocaleLowerCase());else unsupported=true;}
      else if(rule.type==='custom'){const result=evaluateFormula(snapshot,sheet.id,relativeFormula(rule,rule.formulae[0]||'',position),{position});if(result==='#NAME?')unsupported=true;else valid=result===true||typeof result==='number'&&result!==0;}
      else if(['whole','decimal','date','time','textLength'].includes(rule.type)){
        const bounds=rule.formulae.map(f=>evaluateFormula(snapshot,sheet.id,relativeFormula(rule,f,position),{position}));
        let n=rule.type==='textLength'?String(value).length:value;
        if(bounds.some(b=>b==='#NAME?'))unsupported=true;
        else if(typeof n!=='number'||!Number.isFinite(n)||rule.type==='whole'&&!Number.isInteger(n))valid=false;
        else{if(rule.type==='date')n=Math.floor(n);if(rule.type==='time')n=((n%1)+1)%1;const [a,b]=bounds;valid=({between:()=>n>=a&&n<=b,notBetween:()=>n<a||n>b,equal:()=>n===a,notEqual:()=>n!==a,lessThan:()=>n<a,lessThanOrEqual:()=>n<=a,greaterThan:()=>n>a,greaterThanOrEqual:()=>n>=a})[rule.operator]?.()??false;}
      }else unsupported=true;
      if(unsupported)return {ref,severity:'unsupported',title:'입력 제한 확인',message:'이 셀의 입력 제한 수식은 아직 확인할 수 없습니다. Excel에서 조건을 확인하세요.'};
      if(!valid)return {ref,severity:rule.severity,title:rule.title,message:rule.message};
    }
  }
  return null;
}
