import {address} from './engine.mjs';
import {storedArrayIndex} from './stored-arrays.mjs';
/** Excel data-edge navigation depends on occupied cells in the current line,
 * not the last occupied cell elsewhere in the worksheet. Formula-empty strings
 * and declared array followers are occupied; formatting-only cells are not. */
export function dataEdge(sheet,point,axis,direction,{hiddenRow=()=>false,hiddenCol=()=>false}={}){
 const limit=axis==='r'?1048576:16384,other=axis==='r'?'c':'r',array=storedArrayIndex(sheet),occupied=index=>{const p={...point,[axis]:index};return !!sheet.cells[address(p.r,p.c)]?.raw||!!array.at(p);},visible=index=>!(axis==='r'?hiddenRow(index):hiddenCol(index));
 let current=point[axis],next=current+direction;while(next>=0&&next<limit&&!visible(next))next+=direction;if(next<0||next>=limit)return {...point};
 const seekGap=occupied(current)&&occupied(next);current=next;
 if(seekGap){for(let i=current+direction;i>=0&&i<limit;i+=direction){if(!visible(i))continue;if(!occupied(i))break;current=i;}}
 else if(!occupied(current)){for(let i=current+direction;i>=0&&i<limit;i+=direction){if(!visible(i))continue;current=i;if(occupied(i))break;}}
 return {[axis]:current,[other]:point[other]};
}
const contains=(b,p)=>p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2;
export function selectionStep(range,point,{axis='c',direction=1,hiddenRow=()=>false,hiddenCol=()=>false,merges=[]}={}){
 const visible=(i,lo,hi,step,hidden)=>{for(;i>=lo&&i<=hi;i+=step)if(!hidden(i))return i;return null;};
 const firstR=visible(direction>0?range.r1:range.r2,range.r1,range.r2,direction,hiddenRow),firstC=visible(direction>0?range.c1:range.c2,range.c1,range.c2,direction,hiddenCol);if(firstR===null||firstC===null)return point;
 const next=from=>{if(axis==='c'){let c=visible(from.c+direction,range.c1,range.c2,direction,hiddenCol),r=from.r;if(c===null){c=firstC;r=visible(from.r+direction,range.r1,range.r2,direction,hiddenRow)??firstR;}return {r,c};}let r=visible(from.r+direction,range.r1,range.r2,direction,hiddenRow),c=from.c;if(r===null){r=firstR;c=visible(from.c+direction,range.c1,range.c2,direction,hiddenCol)??firstC;}return {r,c};};
 let candidate=next(point);for(let tries=0;tries<merges.length*2+10;tries++){
  const merge=merges.find(b=>contains(b,candidate));if(!merge||candidate.r===merge.r1&&candidate.c===merge.c1)return candidate;
  const from={...candidate};if(axis==='c'){from.c=direction>0?merge.c2:merge.c1;if(merge.c1<=range.c1&&merge.c2>=range.c2)from.r=direction>0?merge.r2:merge.r1;}else{from.r=direction>0?merge.r2:merge.r1;if(merge.r1<=range.r1&&merge.r2>=range.r2)from.c=direction>0?merge.c2:merge.c1;}
  candidate=next(from);if(candidate.r===point.r&&candidate.c===point.c)return point;
 }
 return point;
}
