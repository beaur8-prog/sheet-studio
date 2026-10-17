import {address,shiftFormula} from './engine.mjs';
export function scalarRaw(value) {
  if(value===null||value===undefined)return '';
  if(typeof value==='boolean')return value?'TRUE':'FALSE';
  if(typeof value==='number')return String(value);
  const text=String(value);
  // Preserve text identity (identifiers, formula-looking text and apostrophes).
  return text && (/^[=']/.test(text)||!Number.isNaN(Number(text))||/^(TRUE|FALSE)$/i.test(text))?"'"+text:text;
}
export function pastePlan(sheet, clip, start, options={}) {
  const {mode='all',transpose=false,skipBlanks=false,operation='none',selection}=options;
  if(!clip?.data?.length || !clip.data[0]?.length)throw Error('복사한 셀이 없습니다.');
  if(clip.cut&&(mode!=='all'||transpose||skipBlanks||operation!=='none'))throw Error('선택하여 붙여넣기는 잘라내기 대신 복사한 셀에 사용하세요.');
  const sh=clip.data.length,sw=Math.max(...clip.data.map(r=>r.length)),h=transpose?sw:sh,w=transpose?sh:sw;
  let height=h,width=w;
  if(selection){const rh=selection.r2-selection.r1+1,rw=selection.c2-selection.c1+1;if(rh>=h&&rw>=w&&rh%h===0&&rw%w===0&&!clip.cut){height=rh;width=rw;}}
  if(height*width>100000||start.r+height>1048576||start.c+width>16384)throw Error('붙여넣기 범위가 시트 또는 100,000개 셀 제한을 초과합니다.');
  const updates=[];
  for(let r=0;r<height;r++)for(let c=0;c<width;c++){
    const sr=transpose?c%w:r%h,sc=transpose?r%h:c%w,source=clip.data[sr]?.[sc]||{raw:''};
    if(skipBlanks&&source.raw==='')continue;
    const key=address(start.r+r,start.c+c),old=sheet.cells[key]||{raw:''};let cell;
    if(mode==='formats')cell={...structuredClone(old),style:structuredClone(source.style||{})};
    else {
      cell=mode==='all'?structuredClone(source):structuredClone(old);
      let raw=mode==='values'?scalarRaw(clip.values?.[sr]?.[sc]??source.raw):source.raw;
      if(mode!=='values'&&!clip.cut&&raw.startsWith('='))raw=shiftFormula(raw,start.r+r-(clip.sourceRows?.[sr]??clip.r+sr),start.c+c-(clip.c+sc));
      if(operation!=='none'){
        const a=options.evaluate?.(key)??Number(old.raw||0),b=clip.values?.[sr]?.[sc]??Number(source.raw||0);
        if(typeof a!=='number'||typeof b!=='number')throw Error('연산 붙여넣기는 숫자 셀에 사용하세요.');
        raw=operation==='divide'&&b===0?'#DIV/0!':String(({add:()=>a+b,subtract:()=>a-b,multiply:()=>a*b,divide:()=>a/b})[operation]());
      }
      cell.raw=raw;delete cell.cachedValue;
    }
    updates.push({key,cell});
  }
  return {updates,height,width};
}
