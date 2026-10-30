import {tableBounds} from './table-schema.mjs';
import {unwrap} from './clone.mjs';
import {calculatedArrayRegions} from './calculated-arrays.mjs';
const inside=(a,p)=>p.r>=a.r1&&p.r<=a.r2&&p.c>=a.c1&&p.c<=a.c2;
/** Stored CSE and imported array ranges, indexed without expanding followers. */
export function storedArrayIndex(sheet,{validate=false,regions:provided}={}){
  const regions=provided?[...provided]:[];if(!provided)for(const key in sheet.cells){const cell=sheet.cells[key];if(cell.raw?.startsWith('=')&&cell.xlsxFormula?.shareType==='array'&&!cell.xlsxFormula.dynamic){const b=tableBounds(cell.xlsxFormula.ref||key);if(b)regions.push({ref:key,...b});}}
  if(regions.length>10000)throw Error('한 시트의 배열 수식 정의는 최대 10,000개까지 지원합니다.');regions.sort((a,b)=>a.r1-b.r1||a.c1-b.c1);
  if(validate){let active=[],row=-1;for(const region of regions){if(region.r1!==row){row=region.r1;active=active.filter(a=>a.r2>=row);}let lo=0,hi=active.length;while(lo<hi){const mid=(lo+hi)>>>1;if(active[mid].c1<region.c1)lo=mid+1;else hi=mid;}if(lo>0&&active[lo-1].c2>=region.c1||lo<active.length&&active[lo].c1<=region.c2)throw Error('배열 수식 범위가 서로 겹칩니다.');active.splice(lo,0,region);}}
  const ends=[];let maximum=-1,last=null;for(const region of regions){maximum=Math.max(maximum,region.r2);ends.push(maximum);}
  return {regions,at(position){if(last&&inside(last,position))return last;let lo=0,hi=regions.length;while(lo<hi){const mid=(lo+hi)>>>1;if(regions[mid].r1<=position.r)lo=mid+1;else hi=mid;}for(let i=lo-1;i>=0&&ends[i]>=position.r;i--)if(inside(regions[i],position))return last=regions[i];return null;}};
}
/** Once the complete original array is replaced, remove its old definition so
 * typing a later formula in the anchor cannot resurrect the deleted array.
 */
export function clearReplacedArrayDefinitions(sheet,updates,{editAnchor=false}={}){
  if(!updates.some(u=>u.cell?.xlsxFormula?.shareType==='array'))return;
  const keys=new Set(updates.map(u=>u.key));for(const array of storedArrayIndex(sheet).regions){let count=0;for(const key of keys){const b=tableBounds(key);if(b&&inside(array,{r:b.r1,c:b.c1}))count++;}if(count!==(array.r2-array.r1+1)*(array.c2-array.c1+1)||editAnchor&&count===1&&updates.find(u=>u.key===array.ref)?.cell?.raw?.startsWith('='))continue;const update=updates.find(u=>u.key===array.ref);if(update?.cell){const incoming=update.cell.xlsxFormula;if(incoming?.shareType==='array'&&unwrap(incoming)!==unwrap(sheet.cells[array.ref]?.xlsxFormula))continue;update.cell={...update.cell};delete update.cell.xlsxFormula;}}
}
export function normalizeArrayCopy(sheet,data,values,sourceRows,sourceCols,serialize){
  const index=storedArrayIndex(sheet,{regions:[...storedArrayIndex(sheet).regions,...calculatedArrayRegions(sheet)]}),rows=new Set(sourceRows),cols=new Set(sourceCols),complete=new Map();
  for(let r=0;r<sourceRows.length;r++)for(let c=0;c<sourceCols.length;c++){
    const array=index.at({r:sourceRows[r],c:sourceCols[c]});if(!array)continue;
    if(!complete.has(array.ref)){let all=true;for(let row=array.r1;row<=array.r2&&all;row++)if(!rows.has(row))all=false;for(let col=array.c1;col<=array.c2&&all;col++)if(!cols.has(col))all=false;complete.set(array.ref,all);}
    const anchorIncluded=rows.has(array.r1)&&cols.has(array.c1),isAnchor=sourceRows[r]===array.r1&&sourceCols[c]===array.c1;
    if(array.dynamic&&anchorIncluded){if(!isAnchor){data[r][c].raw='';delete data[r][c].xlsxFormula;}}
    else if(!isAnchor||!complete.get(array.ref)){data[r][c].raw=serialize(values[r][c],r,c);delete data[r][c].xlsxFormula;}
    delete data[r][c].cachedValue;delete data[r][c].cachedFormula;delete data[r][c].cachedError;delete data[r][c].xlsxMetadata;
  }
}
