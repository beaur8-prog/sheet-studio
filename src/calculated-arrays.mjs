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
export function referenceDimensions(node){
  if(!node)return null;
  if(['literal','error','ref','implicit'].includes(node.type))return {rows:1,cols:1};
  if(node.type==='array')return {rows:node.rows,cols:node.cols};
  if(node.type==='range')return {rows:Math.abs(node.start.r-node.end.r)+1,cols:Math.abs(node.start.c-node.end.c)+1};
  if(node.type==='referenceOp'&&node.op===':'){
    const bounds=n=>n.type==='ref'?{r1:n.pos.r,r2:n.pos.r,c1:n.pos.c,c2:n.pos.c}:n.type==='range'?{r1:Math.min(n.start.r,n.end.r),r2:Math.max(n.start.r,n.end.r),c1:Math.min(n.start.c,n.end.c),c2:Math.max(n.start.c,n.end.c)}:null;
    const a=bounds(node.left),b=bounds(node.right);if(a&&b)return {rows:Math.max(a.r2,b.r2)-Math.min(a.r1,b.r1)+1,cols:Math.max(a.c2,b.c2)-Math.min(a.c1,b.c1)+1};
  }
  if(node.type==='whole')return {rows:node.axis==='c'?1048576:Math.abs(node.start-node.end)+1,cols:node.axis==='r'?16384:Math.abs(node.start-node.end)+1};
  return null;
}
export function mayReturnArray(node,resolveShape=referenceDimensions){
  if(!node)return false;
  const array=child=>mayReturnArray(child,resolveShape);
  if(['array','range','whole','structured','spill','name','referenceOp'].includes(node.type)){
    const shape=resolveShape(node);return shape?shape.rows*shape.cols>1: true;
  }
  if(node.type==='implicit')return false;
  if(node.type==='unary')return array(node.node);
  if(node.type==='binary')return array(node.left)||array(node.right);
  if(node.type==='call'){
    const name=node.name.replace(/^_XLFN\.(?:_XLWS\.)?/,'').replace(/^_XLWS\./,'');
    if(['FILTER','SORT','UNIQUE','SEQUENCE','ANCHORARRAY'].includes(name))return true;
    if(['ROW','COLUMN','TRANSPOSE'].includes(name))return node.args.some(array);
    if(name==='XLOOKUP'){
      if(array(node.args[0])||array(node.args[3]))return true;
      const lookup=resolveShape(node.args[1]),result=resolveShape(node.args[2]);
      if(!lookup||!result)return true;
      return lookup.cols===1?result.cols>1:lookup.rows===1&&result.rows>1;
    }
    if(name==='INDEX'){const source=resolveShape(node.args[0]);if(node.args.length===2&&(!source||source.rows>1&&source.cols>1))return true;return node.args.slice(1,3).some(a=>a.type==='call'&&['MATCH','XMATCH'].includes(a.name.replace(/^_XLFN\./,''))?false:a.type!=='literal'||a.value===0);}
    if(name==='OFFSET')return array(node.args[0])||node.args.slice(3).some(a=>a.type!=='literal'||a.value>1);
    if(name==='INDIRECT'){
      const arg=node.args[0];if(arg?.type!=='literal'||typeof arg.value!=='string')return true;
      const local=arg.value.replace(/^(?:'(?:[^']|'')*'|[\p{L}_][\p{L}\p{N}_.]*)!/u,'');
      return !tableBounds(local)||local.includes(':');
    }
    if(['IF','IFERROR','IFNA','IFS','SWITCH','CHOOSE','LET'].includes(name))return node.args.some(array);
    if(ARRAY_SCALAR_FUNCTIONS.has(name))return node.args.some(array);
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
ISERROR ISERR ISNA ISBLANK ISNUMBER ISTEXT ISNONTEXT ISLOGICAL ERROR.TYPE N T NOT
`.trim().split(/\s+/));

export function calculatedArrayRecords(sheet){return (regionsByCells.get(unwrap(sheet.cells))||[]).filter(a=>sheet.cells[a.ref]?.raw===a.raw);}
export function spillExplanation(issue){if(!issue)return '';return ({occupied:'결과 범위의 '+(issue.obstruction||'다른 셀')+'에 내용이 있습니다.',merge:'결과 범위에 병합 셀이 있습니다.',table:'동적 배열 결과는 표 안에 펼칠 수 없습니다.',edge:'결과가 시트의 마지막 행·열을 넘습니다.',capacity:'동적 배열 계산 결과가 문서의 1,000,000개 한도를 넘습니다.',anchors:'한 시트의 배열 수식 10,000개 한도를 넘습니다.',array:'결과 범위에 고정 배열 수식이 있습니다.',spill:'다른 동적 배열 결과와 범위가 겹칩니다.',cycle:'수식이 자신의 배열 결과를 참조합니다.'})[issue.reason]||'';}
