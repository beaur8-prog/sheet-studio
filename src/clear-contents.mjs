import {parseAddress,address} from './engine.mjs';
import {arrayRegions} from './cell-constraints.mjs';
import {MAX_EDIT_CELLS} from './limits.mjs';
const inside=(b,p)=>p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2;
const overlap=(a,b)=>a.r1<=b.r2&&a.r2>=b.r1&&a.c1<=b.c2&&a.c2>=b.c1;
const limit=()=>Error('한 번에 최대 100,000개의 원본 셀 내용을 지울 수 있습니다.');
/** Clearing a sparse selection edits its source records, not millions of
 * derived spill followers or blank cells. Formatting and notes stay in place.
 */
export function clearContentsPlan(sheet,selection,{editable=()=>true}={}){
  const arrays=arrayRegions(sheet).filter(a=>overlap(a,selection)),updates=new Map();
  const add=(ref,cell)=>{if(updates.has(ref))return;if(updates.size>=MAX_EDIT_CELLS)throw limit();updates.set(ref,{key:ref,cell:{...cell,raw:''}});};
  for(const a of arrays){
    if(a.dynamic){if(!inside(selection,{r:a.r1,c:a.c1})||!editable(a.r1,a.c1))throw Error('배열 수식 결과의 일부를 변경할 수 없습니다. 첫 셀을 포함해 선택하거나 첫 셀의 수식을 지우세요.');}
    else{
      if(a.r1<selection.r1||a.r2>selection.r2||a.c1<selection.c1||a.c2>selection.c2)throw Error('배열 수식 결과의 일부를 변경할 수 없습니다. 배열 전체를 선택하세요.');
      if((a.r2-a.r1+1)*(a.c2-a.c1+1)>MAX_EDIT_CELLS)throw limit();
      // Legacy CSE definitions require a complete range replacement. Include
      // absent followers so the existing constraint and metadata rules apply.
      for(let r=a.r1;r<=a.r2;r++)for(let c=a.c1;c<=a.c2;c++)if(editable(r,c))add(address(r,c),sheet.cells[address(r,c)]);
    }
  }
  for(const ref in sheet.cells){const cell=sheet.cells[ref];if(!cell.raw)continue;const p=parseAddress(ref);if(p&&inside(selection,p)&&editable(p.r,p.c))add(ref,cell);}
  return [...updates.values()];
}
