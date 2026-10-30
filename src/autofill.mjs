import {cloneData} from './clone.mjs';
import {address,shiftFormula} from './engine.mjs';

const modulo=(n,size)=>(n%size+size)%size;
const numeric=/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i;
function numberSeed(cell){
 if(!cell||['text','@'].includes(cell.style?.format)||cell.style?.numberFormat==='@')return null;
 const raw=String(cell.raw??'').trim(),percent=raw.endsWith('%'),text=percent?raw.slice(0,-1):raw;
 if(!numeric.test(text))return null;
 const value=Number(text);return Number.isFinite(value)?{value,percent}:null;
}
function numericTrend(cells){
 if(cells.length<2)return null;
 const seeds=cells.map(numberSeed);if(seeds.some(s=>!s)||seeds.some(s=>s.percent!==seeds[0].percent))return null;
 const values=seeds.map(s=>s.value),n=values.length,middle=(n-1)/2,base=values[0],denominator=n*(n*n-1)/12;
 // Center the least-squares calculation to avoid summing large offsets.
 const meanDelta=values.reduce((sum,v)=>sum+(v-base)/n,0),step=values.reduce((sum,v,i)=>sum+(i-middle)*((v-base)/denominator),0);
 return index=>{
  const value=base+meanDelta+(index-middle)*step;
  if(!Number.isFinite(value))throw Error('자동 채우기 숫자가 계산 범위를 초과합니다.');
  return String(Number(value.toPrecision(15)))+(seeds[0].percent?'%':'');
 };
}
const visible=(first,last,hidden)=>Array.from({length:last-first+1},(_,i)=>first+i).filter(i=>!hidden(i));
const inside=(r,c,b)=>r>=b.r1&&r<=b.r2&&c>=b.c1&&c<=b.c2;
function validateRange(b){
 if(!b||!['r1','r2','c1','c2'].every(k=>Number.isInteger(b[k]))||b.r1<0||b.c1<0||b.r2>=1048576||b.c2>=16384||b.r2<b.r1||b.c2<b.c1)throw Error('자동 채우기 범위가 올바르지 않습니다.');
 if((b.r2-b.r1+1)*(b.c2-b.c1+1)>100000)throw Error('한 번에 최대 100,000개 셀을 채울 수 있습니다.');
}

export function autoFillPlan({sourceCells,sourceRange,targetRange,hiddenRow=()=>false,hiddenCol=()=>false,isEditable=()=>true}){
 validateRange(sourceRange);validateRange(targetRange);
 const s=sourceRange,t=targetRange;
 if(!inside(s.r1,s.c1,t)||!inside(s.r2,s.c2,t))throw Error('자동 채우기 범위에 원본 셀이 포함되어야 합니다.');
 const rows=visible(t.r1,t.r2,hiddenRow),cols=visible(t.c1,t.c2,hiddenCol),sourceRows=rows.filter(r=>r>=s.r1&&r<=s.r2),sourceCols=cols.filter(c=>c>=s.c1&&c<=s.c2);
 if(!sourceRows.length||!sourceCols.length)return [];
 const rowOffset=rows.indexOf(sourceRows[0]),colOffset=cols.indexOf(sourceCols[0]);
 const vertical=t.c1===s.c1&&t.c2===s.c2&&(t.r1<s.r1||t.r2>s.r2),horizontal=t.r1===s.r1&&t.r2===s.r2&&(t.c1<s.c1||t.c2>s.c2);
 const trends=new Map();
 if(vertical)for(const c of sourceCols)trends.set(c,numericTrend(sourceRows.map(r=>sourceCells[address(r,c)])));
 if(horizontal)for(const r of sourceRows)trends.set(r,numericTrend(sourceCols.map(c=>sourceCells[address(r,c)])));
 const updates=[];
 rows.forEach((r,ri)=>cols.forEach((c,ci)=>{
  if(inside(r,c,s)||!isEditable(r,c))return;
  const sr=sourceRows[modulo(ri-rowOffset,sourceRows.length)],sc=sourceCols[modulo(ci-colOffset,sourceCols.length)],cell=cloneData(sourceCells[address(sr,sc)]||{raw:''});
  const trend=trends.get(vertical?sc:sr);
  if(trend)cell.raw=trend(vertical?ri-rowOffset:ci-colOffset);
  else if(cell.raw?.startsWith('='))cell.raw=shiftFormula(cell.raw,r-sr,c-sc);
  delete cell.cachedValue;
  delete cell.cachedError;
  updates.push({key:address(r,c),cell});
 }));
 return updates;
}
