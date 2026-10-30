import {stripRichTextFormat} from './cell-format-reset.mjs';
import {cloneData} from './clone.mjs';
import {address,parseAddress,shiftFormula,transposeFormula,numberFormatKind,formulaForXlsx} from './engine.mjs';
import {tableBounds,overlaps} from './table-schema.mjs';
import {isErrorText,entryRaw} from './cell-values.mjs';
export function scalarRaw(value,{error=false}={}) {
  if(value===null||value===undefined)return '';
  if(typeof value==='boolean')return value?'TRUE':'FALSE';
  if(typeof value==='number')return String(value);
  const text=String(value);
  // Preserve text identity (identifiers, formula-looking text and apostrophes).
  return text && (/^[=']/.test(text)||!Number.isNaN(Number(text))||/^(TRUE|FALSE)$/i.test(text)||!error&&isErrorText(text))?"'"+text:text;
}
export function formatCellRecord(old,source){
 const cell={...cloneData(old),style:cloneData(source.style||{})};
 for(const key of ['xlsxStyle','xlsxNormal','xlsxProtection']){if(source[key]!==undefined)cell[key]=cloneData(source[key]);else delete cell[key];}
 return stripRichTextFormat(cell);
}
export function pastePlan(sheet, clip, start, options={}) {
  const {mode='all',transpose=false,skipBlanks=false,operation='none',selection}=options;
  if(!clip?.data?.length || !clip.data[0]?.length)throw Error('복사한 셀이 없습니다.');
  if(clip.cut&&(mode!=='all'||transpose||skipBlanks||operation!=='none'))throw Error('선택하여 붙여넣기는 잘라내기 대신 복사한 셀에 사용하세요.');
  const sh=clip.data.length,sw=Math.max(...clip.data.map(r=>r.length)),h=transpose?sw:sh,w=transpose?sh:sw;
  let height=h,width=w;
  if(selection){const rh=selection.r2-selection.r1+1,rw=selection.c2-selection.c1+1;if(rh>=h&&rw>=w&&rh%h===0&&rw%w===0&&!clip.cut){height=rh;width=rw;}}
  if(height*width>100000||start.r+height>1048576||start.c+width>16384)throw Error('붙여넣기 범위가 시트 또는 100,000개 셀 제한을 초과합니다.');
  const targetRows=[],targetCols=[];
  for(let r=start.r;targetRows.length<height&&r<1048576;r++)if(!options.hiddenRow?.(r))targetRows.push(r);
  for(let c=start.c;targetCols.length<width&&c<16384;c++)if(!options.hiddenCol?.(c))targetCols.push(c);
  if(targetRows.length<height||targetCols.length<width)throw Error('가시 셀 붙여넣기 범위가 시트 경계를 초과합니다.');
  if(clip.cut&&(targetRows.some((r,i)=>r!==start.r+i)||targetCols.some((c,i)=>c!==start.c+i)))throw Error('숨김 셀을 가로지르는 범위로 잘라내어 이동할 수 없습니다. 숨김을 해제하세요.');
  const arrays=clip.data.some(row=>row.some(cell=>cell?.xlsxFormula?.shareType==='array'))&&['all','formulas'].includes(mode)&&operation==='none';if(arrays&&transpose)throw Error('배열 수식은 원래 방향으로 붙여넣거나 값으로 전치하세요.');if(arrays&&(skipBlanks||targetRows.some((r,i)=>r!==start.r+i)||targetCols.some((c,i)=>c!==start.c+i)))throw Error('배열 수식은 빈 셀 건너뛰기 없이 연속된 범위에 붙여넣으세요.');
  const rect={r1:start.r,c1:start.c,r2:targetRows.at(-1),c2:targetCols.at(-1)};
  const merges=(sheet.merges||[]).filter(m=>{const [a,b]=m.split(':').map(parseAddress),overlap=a.r<=rect.r2&&b.r>=rect.r1&&a.c<=rect.c2&&b.c>=rect.c1;if(overlap&&(a.r<rect.r1||b.r>rect.r2||a.c<rect.c1||b.c>rect.c2))throw Error('병합 셀 일부에 붙여넣을 수 없습니다. 병합 전체를 선택하거나 병합을 해제하세요.');return !overlap;});
  for(const m of clip.merges||[]){const [a,b]=m.split(':').map(parseAddress);if(transpose||targetRows.some((r,i)=>r!==start.r+i)||targetCols.some((c,i)=>c!==start.c+i))throw Error('병합 셀은 연속된 범위에 원래 방향으로 붙여넣으세요.');for(let r=0;r<height;r+=h)for(let c=0;c<width;c+=w)merges.push(address(start.r+r+a.r-clip.r,start.c+c+a.c-clip.c)+':'+address(start.r+r+b.r-clip.r,start.c+c+b.c-clip.c));}
  const updates=[];
  for(let r=0;r<height;r++)for(let c=0;c<width;c++){
    const sr=transpose?c%w:r%h,sc=transpose?r%h:c%w,source=clip.data[sr]?.[sc]||{raw:''};
    if(skipBlanks&&source.raw==='')continue;
    const key=address(targetRows[r],targetCols[c]),old=sheet.cells[key]||{raw:''};let cell;
    if(mode==='formats')cell=formatCellRecord(old,source);
    else {
      cell=mode==='all'&&!clip.plainText?cloneData(source):cloneData(old);
      let raw=mode==='values'?scalarRaw(clip.values?.[sr]?.[sc]??source.raw,{error:clip.errors?.[sr]?.[sc]}):source.raw;
      if(clip.plainText)raw=entryRaw(raw,options.styleAt?.(targetRows[r],targetCols[c])||old.style);
      const dateStyle=source.style?.format==='date'||numberFormatKind(source.style?.numberFormat)==='date';
      if(operation==='none'&&typeof clip.date1904==='boolean'&&clip.date1904!==(options.date1904===true)&&dateStyle&&!raw.startsWith("'")&&raw.trim()!==''&&Number.isFinite(Number(raw)))raw=String(Number(raw)+(clip.date1904?1462:-1462));
      if(mode!=='values'&&!clip.cut&&raw.startsWith('=')){const origin={r:clip.sourceRows?.[sr]??clip.r+sr,c:clip.sourceCols?.[sc]??clip.c+sc},dest={r:targetRows[r],c:targetCols[c]};raw=transpose?transposeFormula(raw,origin,dest):shiftFormula(raw,dest.r-origin.r,dest.c-origin.c);}
      if(operation!=='none'){
        const a=options.evaluate?.(key)??Number(old.raw||0),b=clip.values?.[sr]?.[sc]??Number(source.raw||0);
        if(typeof a!=='number'||typeof b!=='number')throw Error('연산 붙여넣기는 숫자 셀에 사용하세요.');
        raw=operation==='divide'&&b===0?'#DIV/0!':String(({add:()=>a+b,subtract:()=>a-b,multiply:()=>a*b,divide:()=>a/b})[operation]());
      }
      cell.raw=raw;delete cell.cachedValue;delete cell.cachedFormula;delete cell.cachedError;
      if(source.xlsxFormula?.shareType==='array'&&arrays){const range=tableBounds(source.xlsxFormula.ref||address(clip.r+sr,clip.c+sc)),origin={r:clip.sourceRows?.[sr]??clip.r+sr,c:clip.sourceCols?.[sc]??clip.c+sc},dest={r:targetRows[r],c:targetCols[c]},rows=range.r2-range.r1+1,cols=range.c2-range.c1+1,b={r1:dest.r,c1:dest.c,r2:dest.r+(transpose?cols:rows)-1,c2:dest.c+(transpose?rows:cols)-1};if((sheet.tables||[]).some(t=>overlaps(tableBounds(t.range),b)))throw Error('표 안에는 배열 수식을 붙여넣을 수 없습니다.');cell.xlsxFormula=cloneData(source.xlsxFormula);cell.xlsxFormula.ref=address(b.r1,b.c1)+':'+address(b.r2,b.c2);cell.xlsxFormula.raw=raw;cell.xlsxFormula.formula=formulaForXlsx(raw,source.xlsxFormula.formula).slice(1);}
      else if(source.xlsxFormula&&mode!=='values'&&operation==='none'&&raw.startsWith('='))cell.xlsxFormula={...cloneData(source.xlsxFormula),raw,formula:formulaForXlsx(raw,source.xlsxFormula.formula).slice(1)};
      else if(mode!=='formats')delete cell.xlsxFormula;
    }
    updates.push({key,cell});
  }
  const rowHeightUpdates=[];
  // External table rows are minimum heights. A shorter table frame must never
  // compress them, and an existing taller worksheet row must remain tall.
  if(mode==='all'&&!transpose&&operation==='none'&&Array.isArray(clip.rowHeights))for(let r=0;r<height;r++){
    const sourceHeight=clip.rowHeights[r%h],row=targetRows[r],current=sheet.rowHeights?.[row]??sheet.defaultRowHeight??25;
    if(typeof sourceHeight!=='number'||!Number.isFinite(sourceHeight)||sourceHeight<1||sourceHeight>10000||current===0)continue;
    if(skipBlanks&&!clip.data[r%h]?.some(cell=>(cell?.raw??'')!==''))continue;
    if(sourceHeight>current)rowHeightUpdates.push({row,height:sourceHeight});
  }
  return {updates,rowHeightUpdates,height:rect.r2-start.r+1,width:rect.c2-start.c+1,merges};
}
