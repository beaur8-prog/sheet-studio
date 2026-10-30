import {calculatedArrayRegions} from './calculated-arrays.mjs';
import {storedArrayIndex} from './stored-arrays.mjs';
import {tableAt,tableBounds} from './table-schema.mjs';
import {address,parseAddress,createEvaluator} from './engine.mjs';

const contains=(b,r,c)=>r>=b.r1&&r<=b.r2&&c>=b.c1&&c<=b.c2;
const equal=(a,b)=>['r1','r2','c1','c2'].every(k=>a[k]===b[k]);
function add(map,key,value){const old=map.get(key);if(old===undefined)map.set(key,value);else if(typeof old==='number')map.set(key,[old,value]);else old.push(value);}
function intersects(values,lo,hi){
  if(values===undefined)return false;if(typeof values==='number')return values>=lo&&values<=hi;
  let left=0,right=values.length;while(left<right){const mid=(left+right)>>>1;if(values[mid]<lo)left=mid+1;else right=mid;}return left<values.length&&values[left]<=hi;
}
/** The occupied region around a cell, bounded by completely blank rows/columns.
 * Index each occupied cell once rather than repeatedly scanning the entire sheet.
 */
export function currentRegion(sheet,point){
  const r=point.r??point.r1,c=point.c??point.c1,rows=new Map(),cols=new Map(),arrays=[...storedArrayIndex(sheet).regions,...calculatedArrayRegions(sheet)];let found=arrays.some(a=>contains(a,r,c));
  for(const ref in sheet.cells){const cell=sheet.cells[ref];if(!cell?.raw||cell.raw==="'")continue;const p=parseAddress(ref);if(!p)continue;add(rows,p.r,p.c);add(cols,p.c,p.r);if(p.r===r&&p.c===c)found=true;}
  if(!found)return null;
  for(const values of rows.values())if(Array.isArray(values))values.sort((a,b)=>a-b);
  for(const values of cols.values())if(Array.isArray(values))values.sort((a,b)=>a-b);
  const b={r1:r,r2:r,c1:c,c2:c};let changed=true;
  while(changed){changed=false;
    for(const a of arrays)if(a.r1<=b.r2+1&&a.r2>=b.r1-1&&a.c1<=b.c2+1&&a.c2>=b.c1-1){const next={r1:Math.min(b.r1,a.r1),r2:Math.max(b.r2,a.r2),c1:Math.min(b.c1,a.c1),c2:Math.max(b.c2,a.c2)};if(!equal(b,next)){Object.assign(b,next);changed=true;}}
    if(b.r1>0&&intersects(rows.get(b.r1-1),b.c1-1,b.c2+1)){b.r1--;changed=true;}
    if(b.r2<1048575&&intersects(rows.get(b.r2+1),b.c1-1,b.c2+1)){b.r2++;changed=true;}
    if(b.c1>0&&intersects(cols.get(b.c1-1),b.r1-1,b.r2+1)){b.c1--;changed=true;}
    if(b.c2<16383&&intersects(cols.get(b.c2+1),b.r1-1,b.r2+1)){b.c2++;changed=true;}
  }
  return b;
}
function headerGuess(sheet,b,evaluate){
  if(b.r1===b.r2)return {headers:false,ambiguous:false};
  let text=0,typed=0,empty=0,bodyTyped=0,bold=0,bodyBold=0;
  for(let c=b.c1;c<=b.c2;c++){
    const ref=address(b.r1,c),cell=sheet.cells[ref],value=evaluate(ref);
    if(value==='')empty++;else if(typeof value==='string'&&!value.startsWith('#'))text++;else typed++;
    if(cell?.style?.bold)bold++;
    for(let r=b.r1+1;r<=Math.min(b.r2,b.r1+8);r++){
      const key=address(r,c),v=evaluate(key);if(typeof v==='number'||typeof v==='boolean')bodyTyped++;
      if(sheet.cells[key]?.style?.bold)bodyBold++;
    }
  }
  if(typed)return {headers:false,ambiguous:false};
  if(text&&bodyTyped||text&&bold&&!bodyBold)return {headers:true,ambiguous:false};
  return {headers:false,ambiguous:!!(text||empty)};
}
export function inferSortPlan(book,sheet,selection,column=selection.c1){
  const table=tableAt(sheet,selection.r1,Math.max(selection.c1,Math.min(selection.c2,column)));if(table){const b=tableBounds(table.range),range={...b,r2:b.r2-(table.totalsRow?1:0)};return {range,column:Math.max(b.c1,Math.min(b.c2,column)),headers:table.headerRow,review:false,expandedRange:null,tableId:table.id};}
  const evaluate=createEvaluator(book,sheet.id),single=selection.r1===selection.r2&&selection.c1===selection.c2;
  const region=currentRegion(sheet,{r:selection.r1,c:Math.max(selection.c1,Math.min(selection.c2,column))});
  let range={...selection};
  if(single||selection.r1===selection.r2){
    if(!region)throw Error('정렬할 데이터 안의 셀을 선택하세요.');range={...region};
  }else if(selection.r2===1048575||selection.c2===16383){
    let lastRow=range.r1,lastCol=range.c1,found=false;
    for(const ref in sheet.cells){const cell=sheet.cells[ref],p=parseAddress(ref);if(cell?.raw&&p&&contains(range,p.r,p.c)){found=true;lastRow=Math.max(lastRow,p.r);lastCol=Math.max(lastCol,p.c);}}
    if(!found)throw Error('선택한 범위에 데이터가 없습니다.');
    if(selection.r2===1048575)range.r2=lastRow;if(selection.c2===16383)range.c2=lastCol;
  }
  let expandedRange=null;
  if(!single&&selection.r1!==selection.r2&&region&&range.r1>=region.r1&&range.r2<=region.r2&&(range.c1>region.c1||range.c2<region.c2))expandedRange={...range,c1:Math.min(range.c1,region.c1),c2:Math.max(range.c2,region.c2)};
  if(expandedRange&&equal(expandedRange,range))expandedRange=null;
  const guess=headerGuess(sheet,range,evaluate),expandedGuess=expandedRange?headerGuess(sheet,expandedRange,evaluate):null;
  // A selection beginning inside a table cannot contain that table's header.
  if(region&&range.r1>region.r1){guess.headers=false;guess.ambiguous=false;if(expandedGuess){expandedGuess.headers=false;expandedGuess.ambiguous=false;}}
  return {range,expandedRange,column:Math.max(range.c1,Math.min(range.c2,column)),...guess,expandedGuess,review:!!expandedRange||guess.ambiguous};
}
