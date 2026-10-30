import {filteredRowsForSheet} from './table-schema.mjs';

// Sparse prefix sums avoid walking the million-row sheet when an object is far
// from the origin. Filtered and manually hidden dimensions occupy no print space.
function axis(normal,sizes,limit,hidden=new Set()){
 const overrides=new Map(Object.entries(sizes||{}).map(([key,value])=>[Number(key),value]));for(const index of hidden)overrides.set(index,0);
 const keys=[...overrides.keys()].filter(i=>i>=0&&i<limit).sort((a,b)=>a-b),prefix=[0];for(const i of keys)prefix.push(prefix.at(-1)+overrides.get(i)-normal);
 const before=index=>{let lo=0,hi=keys.length;while(lo<hi){const mid=(lo+hi)>>>1;if(keys[mid]<index)lo=mid+1;else hi=mid;}return lo;};
 const offset=index=>index*normal+prefix[before(index)],size=index=>overrides.get(index)??normal;
 const at=pixel=>{let lo=0,hi=limit;while(lo<hi){const mid=(lo+hi)>>>1;if(offset(mid+1)<=pixel)lo=mid+1;else hi=mid;}return Math.min(limit-1,lo);};
 return {offset,size,at};
}
export function printObjectAxes(sheet){return {rows:axis(sheet.defaultRowHeight??25,sheet.rowHeights,1048576,filteredRowsForSheet(sheet)),cols:axis(sheet.defaultColumnWidth??90,sheet.columnWidths,16384)};}
export function objectCellBounds(rect,axes){return {r1:axes.rows.at(rect.y),r2:axes.rows.at(Math.max(rect.y,rect.y+rect.height-.001)),c1:axes.cols.at(rect.x),c2:axes.cols.at(Math.max(rect.x,rect.x+rect.width-.001))};}

function slices(indices,originals,axis){
 const allowed=originals?new Set(originals):null,runs=[];let target=0;
 for(const index of indices){const size=axis.size(index),start=axis.offset(index);if(size>0&&(!allowed||allowed.has(index))){const previous=runs.at(-1);if(previous&&Math.abs(previous.end-start)<.001&&Math.abs(previous.target+previous.end-previous.start-target)<.001)previous.end=start+size;else runs.push({start,end:start+size,target});}target+=size;}
 return runs;
}
/** Clip a sheet object to the original cells on this page. Repeated title cells
 * consume space, but do not duplicate drawings on every subsequent page. */
export function printObjectFragments(rect,page,axes){
 const result=[];for(const y of slices(page.rows,page.objectRows,axes.rows))for(const x of slices(page.cols,page.objectCols,axes.cols)){
  const left=Math.max(rect.x,x.start),right=Math.min(rect.x+rect.width,x.end),top=Math.max(rect.y,y.start),bottom=Math.min(rect.y+rect.height,y.end);if(right<=left||bottom<=top)continue;
  result.push({x:x.target+left-x.start,y:y.target+top-y.start,width:right-left,height:bottom-top,contentX:rect.x-left,contentY:rect.y-top});
 }return result;
}
