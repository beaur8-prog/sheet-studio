import {unwrap} from './clone.mjs';
import {tableBounds} from './table-schema.mjs';

// Derived calculation metadata is deliberately absent from workbook JSON,
// history patches and automatic-recovery records.
const regionsByCells=new WeakMap();
export function setCalculatedArrays(sheet,regions){
  if(!sheet?.cells)return;
  regionsByCells.set(unwrap(sheet.cells),regions.filter(a=>a.dynamic&&tableBounds(a.ref)&&Number.isInteger(a.r1)&&Number.isInteger(a.r2)&&Number.isInteger(a.c1)&&Number.isInteger(a.c2)).map(a=>({...a})));
}
export function calculatedArrayRegions(sheet){
  return calculatedArrayRecords(sheet).filter(a=>!a.blocked);
}
export function mayReturnArray(node){
  if(!node)return false;
  if(['array','range','whole','structured','spill','name','referenceOp'].includes(node.type))return true;
  if(node.type==='implicit')return false;
  if(node.type==='unary')return mayReturnArray(node.node);
  if(node.type==='binary')return mayReturnArray(node.left)||mayReturnArray(node.right);
  if(node.type==='call'){
    const name=node.name.replace(/^_XLFN\.(?:_XLWS\.)?/,'').replace(/^_XLWS\./,'');
    if(['FILTER','SORT','UNIQUE','SEQUENCE','ANCHORARRAY'].includes(name))return true;
    if(['ROW','COLUMN','TRANSPOSE'].includes(name))return node.args.some(mayReturnArray);
    if(name==='XLOOKUP')return mayReturnArray(node.args[0])||mayReturnArray(node.args[2])||mayReturnArray(node.args[3]);
    if(name==='INDEX')return node.args.slice(1,3).some(a=>a.type!=='literal'||a.value===0);
    if(name==='OFFSET')return mayReturnArray(node.args[0])||node.args.slice(3).some(a=>a.type!=='literal'||a.value>1);
    if(name==='INDIRECT')return node.args[0]?.type!=='literal'||typeof node.args[0].value==='string'&&node.args[0].value.includes(':');
    if(['IF','IFERROR','IFNA','IFS','SWITCH','CHOOSE','LET'].includes(name))return node.args.some(mayReturnArray);
    if(ARRAY_SCALAR_FUNCTIONS.has(name))return node.args.some(mayReturnArray);
  }
  return false;
}
// Aggregates and reference consumers must not be lifted element by element.
export const ARRAY_SCALAR_FUNCTIONS=new Set(`
ABS ACOS ACOSH ACOT ACOTH ASIN ASINH ATAN ATANH ATAN2 COS COSH COT COTH CSC CSCH
DEGREES EVEN EXP FACT FACTDOUBLE INT LN LOG LOG10 MOD MROUND ODD POWER QUOTIENT
RADIANS ROUND ROUNDDOWN ROUNDUP SIGN SIN SINH SQRT SQRTPI TAN TANH TRUNC
LEN LEFT RIGHT MID TRIM UPPER LOWER CLEAN EXACT FIND SEARCH SUBSTITUTE REPLACE REPT
TEXT VALUE NUMBERVALUE CHAR CODE UNICHAR UNICODE TEXTBEFORE TEXTAFTER
YEAR MONTH DAY DATE DATEVALUE TIME TIMEVALUE HOUR MINUTE SECOND WEEKDAY WEEKNUM
EOMONTH EDATE ISOWEEKNUM DAYS DAYS360 YEARFRAC
`.trim().split(/\s+/));

export function calculatedArrayRecords(sheet){return (regionsByCells.get(unwrap(sheet.cells))||[]).filter(a=>sheet.cells[a.ref]?.raw===a.raw);}
export function spillExplanation(issue){if(!issue)return '';return ({occupied:'결과 범위의 '+(issue.obstruction||'다른 셀')+'에 내용이 있습니다.',merge:'결과 범위에 병합 셀이 있습니다.',table:'동적 배열 결과는 표 안에 펼칠 수 없습니다.',edge:'결과가 시트의 마지막 행·열을 넘습니다.',capacity:'동적 배열 계산 결과가 문서의 1,000,000개 한도를 넘습니다.',anchors:'한 시트의 배열 수식 10,000개 한도를 넘습니다.',array:'결과 범위에 고정 배열 수식이 있습니다.',spill:'다른 동적 배열 결과와 범위가 겹칩니다.',cycle:'수식이 자신의 배열 결과를 참조합니다.'})[issue.reason]||'';}
