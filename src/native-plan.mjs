import {address, parseAddress, transformFormulaReferences} from './engine.mjs';

const clone=value=>structuredClone(value);
const equal=(a,b)=>JSON.stringify(a??null)===JSON.stringify(b??null);
export function attachNativeSource(workbook, data, filename) {
  const bytes=data instanceof Uint8Array?data:new Uint8Array(data);
  if(bytes.length>15*1024*1024)throw new Error('매크로 통합 문서는 현재 15 MB까지 지원합니다.');
  let binary='';for(let at=0;at<bytes.length;at+=32768)binary+=String.fromCharCode(...bytes.subarray(at,at+32768));
  workbook.nativeSource={format:'xlsm',filename,base64:btoa(binary),baseline:clone(workbook.sheets),operations:[]};
  return workbook;
}

// Mirror structural edits in the comparison baseline. Excel performs the real
// row/column operations first, preserving named ranges, controls and charts.
export function recordNativeStructure(workbook,sheetId,axis,index,delta){
  const source=workbook.nativeSource;if(!source)return;
  source.operations.push({sheetId,axis,index,delta});
  const target=source.baseline.find(s=>s.id===sheetId);if(!target)return;
  const cells={};for(const [key,cell] of Object.entries(target.cells)){
    const p=parseAddress(key);if(delta<0&&p[axis]===index)continue;
    if(p[axis]>=index)p[axis]+=delta;
    const next=address(p.r,p.c);if(next!=='#REF!')cells[next]=cell;
  }target.cells=cells;
  for(const owner of source.baseline)for(const cell of Object.values(owner.cells))if(cell.raw.startsWith('='))cell.raw=transformFormulaReferences(cell.raw,{axis,index,delta,targetSheet:target.name,ownerSheet:owner.name});
  target.merges=(target.merges||[]).map(value=>transformFormulaReferences(value,{axis,index,delta,targetSheet:target.name,ownerSheet:target.name})).filter(value=>!value.includes('#REF!'));
  const prop=axis==='r'?'rowHeights':'columnWidths',sizes={};for(const [key,value] of Object.entries(target[prop]||{})){let n=+key;if(delta<0&&n===index)continue;if(n>=index)n+=delta;sizes[n]=value;}target[prop]=sizes;
}

export function buildNativePlan(workbook){
  const source=workbook.nativeSource;if(source?.format!=='xlsm')throw new Error('먼저 XLSM 통합 문서를 여세요.');
  const baseline=new Map(source.baseline.map(s=>[s.id,s]));
  const sheets=workbook.sheets.map(sheet=>{
    const before=baseline.get(sheet.id),cells=[];
    for(const key of new Set([...Object.keys(before?.cells||{}),...Object.keys(sheet.cells)])){
      const old=before?.cells[key]||{raw:''},now=sheet.cells[key]||{raw:''},patch={address:key};
      if(old.raw!==now.raw)patch.raw=now.raw;
      const style={};for(const name of new Set([...Object.keys(old.style||{}),...Object.keys(now.style||{})]))if(!equal(old.style?.[name],now.style?.[name]))style[name]=now.style?.[name]??null;
      if(Object.keys(style).length)patch.style=style;
      if(!equal(old.note,now.note))patch.note=now.note||'';
      if(Object.keys(patch).length>1)cells.push(patch);
    }
    const dimensions=prop=>Object.fromEntries([...new Set([...Object.keys(before?.[prop]||{}),...Object.keys(sheet[prop]||{})])].filter(key=>!equal(before?.[prop]?.[key],sheet[prop]?.[key])).map(key=>[key,sheet[prop]?.[key]??null]));
    return {id:sheet.id,sourceName:before?.name??null,name:sheet.name,cells,columnWidths:dimensions('columnWidths'),rowHeights:dimensions('rowHeights'),unmerge:(before?.merges||[]).filter(v=>!(sheet.merges||[]).includes(v)),merge:(sheet.merges||[]).filter(v=>!(before?.merges||[]).includes(v))};
  });
  return {pivots:workbook.sheets.filter(s=>s.pivot).map(s=>({sheet:s.name,source:workbook.sheets.find(source=>source.id===s.pivot.sourceSheetId)?.name||null,spec:clone(s.pivot)})),sheets,deleted:source.baseline.filter(s=>!workbook.sheets.some(v=>v.id===s.id)).map(s=>s.name),operations:clone(source.operations||[]),activeSheetId:workbook.activeSheetId};
}

export function nativeSourceChanged(workbook){const p=buildNativePlan(workbook);return !!(p.deleted.length||p.operations.length||p.sheets.some(s=>!s.sourceName||s.name!==s.sourceName||s.cells.length||s.merge.length||s.unmerge.length||Object.keys(s.columnWidths).length||Object.keys(s.rowHeights).length));}

// This is a labelled display fallback, never a replacement for VBA evaluation.
export function nativeCachedValue(workbook,sheetId,key,computed){
  if(computed!=='#NAME?'||!workbook.nativeSource)return undefined;
  const cell=workbook.sheets.find(s=>s.id===sheetId)?.cells[key];
  const original=workbook.nativeSource.baseline.find(s=>s.id===sheetId)?.cells[key];
  return cell?.raw===original?.raw&&cell?.raw?.startsWith('=')?cell.cachedValue:undefined;
}
