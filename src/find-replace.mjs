import {parseAddress,createEvaluator,formatValue,validateWorkbook,validateCell} from './engine.mjs';
import {cloneData,unwrap} from './clone.mjs';
import {worksheetCellStyle} from './cell-style.mjs';
import {assertCellUpdates,validateCellUpdates} from './cell-constraints.mjs';
import {assertNativeCellTransfers} from './native-cell-features.mjs';
import {prepareTableHeaderEdits,applyTableHeaderEdits} from './tables.mjs';
import {clearReplacedArrayDefinitions} from './stored-arrays.mjs';
import {preservedPivotTouches} from './native-pivots.mjs';
import {calculationDisplay} from './calculation-display.mjs';
import {MAX_DOCUMENT_TEXT_CHARS} from './limits.mjs';

const escaped=text=>text.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
const searchableRaw=raw=>raw.startsWith("'")?raw.slice(1):raw;
export function findOptions(book,options={}){
  const query=String(options.query??''),scope=options.scope||'sheet',lookIn=options.lookIn||'formulas';
  if(!query||query.length>32768)throw Error('찾을 내용은 1~32,768자로 입력하세요.');
  if(!['sheet','selection','workbook'].includes(scope)||!['formulas','values'].includes(lookIn))throw Error('검색 범위 또는 대상이 올바르지 않습니다.');
  if(!book.sheets.some(s=>s.id===options.sheetId))throw Error('검색할 시트를 찾을 수 없습니다.');
  const range=options.range;
  if(scope==='selection'&&(!range||!['r1','r2','c1','c2'].every(k=>Number.isInteger(range[k]))||range.r1<0||range.c1<0||range.r2<range.r1||range.c2<range.c1||range.r2>=1048576||range.c2>=16384))throw Error('검색할 선택 범위를 확인하세요.');
  const pattern=(options.wholeCell?'^(?:':'')+escaped(query)+(options.wholeCell?')(?![\\s\\S])':'');
  return {...options,query,scope,lookIn,range,matchCase:!!options.matchCase,wholeCell:!!options.wholeCell,limit:Math.max(1,Math.min(1000,Number.isInteger(options.limit)?options.limit:1000)),pattern};
}
export function replacementRaw(raw,options){
  const text=searchableRaw(raw),replacement=String(options.replacement??'');
  const result=text.replace(new RegExp(options.pattern,options.matchCase?'gu':'giu'),()=>replacement);
  return raw.startsWith("'")?"'"+result:result;
}
function sheets(book,o){return book.sheets.filter(s=>s.kind!=='chart'&&(o.scope==='workbook'?o.includeHiddenSheets||!s.state||s.state==='visible':s.id===o.sheetId));}
function inRange(ref,o){if(o.scope!=='selection')return true;const p=parseAddress(ref),b=o.range;return p&&p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2;}
function keepResult(heap,item,limit){
  if(heap.length===limit&&item.order>=heap[0].order)return;
  if(heap.length<limit){heap.push(item);let i=heap.length-1;while(i){const p=(i-1)>>1;if(heap[p].order>=heap[i].order)break;[heap[p],heap[i]]=[heap[i],heap[p]];i=p;}return;}
  heap[0]=item;let i=0;
  while(true){let next=i,left=i*2+1,right=left+1;if(left<heap.length&&heap[left].order>heap[next].order)next=left;if(right<heap.length&&heap[right].order>heap[next].order)next=right;if(next===i)break;[heap[next],heap[i]]=[heap[i],heap[next]];i=next;}
}

/** Read-only scan; a bounded heap keeps the first matches in workbook/row order
 * without allocating and sorting a second million-cell document. */
export function findWorkbook(book,options,providedEvaluator){
  const o=findOptions(book,options),heap=[],regex=new RegExp(o.pattern,o.matchCase?'u':'iu');let count=0;
  for(const sheet of sheets(book,o)){
    const index=book.sheets.indexOf(sheet),evaluate=providedEvaluator?.(sheet.id)||(o.lookIn==='values'?createEvaluator(book,sheet.id):null);
    const visit=ref=>{
      if(!inRange(ref,o))return;const cell=sheet.cells[ref],raw=cell?.raw||'';
      let text=searchableRaw(raw);
      if(o.lookIn==='values'){
        const result=calculationDisplay(evaluate.result(ref)),value=result.error?result.error:result.value,p=parseAddress(ref);
        text=String(formatValue(value,worksheetCellStyle(book,sheet,p.r,p.c),{date1904:book.date1904,error:result.error})??'');
      }
      if(!regex.test(text))return;count++;const p=parseAddress(ref),order=index*1048576*16384+p.r*16384+p.c;
      const item={sheetId:sheet.id,sheetName:sheet.name,ref,text:text.slice(0,240),order};
      if(o.lookIn==='formulas'&&o.replacement!==undefined)item.after=searchableRaw(replacementRaw(raw,o)).slice(0,240);
      keepResult(heap,item,o.limit);
    };
    for(const ref in sheet.cells)visit(ref);
    if(evaluate)for(const a of evaluate.spills({includeBlocked:true})){if(a.blocked)continue;for(let r=a.r1;r<=a.r2;r++)for(let c=a.c1;c<=a.c2;c++){let col='',n=c+1;while(n){col=String.fromCharCode(65+(n-1)%26)+col;n=Math.floor((n-1)/26);}const ref=col+(r+1);if(!Object.hasOwn(sheet.cells,ref))visit(ref);}}
  }
  return {count,results:heap.sort((a,b)=>a.order-b.order),truncated:count>o.limit};
}

/** Validate a copy before publishing any changes, including cross-sheet table
 * renames and dependent validation rules. The caller commits this as one undo. */
export function replaceWorkbookPlan(book,options){
  book=unwrap(book);
  const o=findOptions(book,options);if(o.lookIn!=='formulas')throw Error('표시 값은 검색 전용입니다. 바꾸려면 검색 대상을 값·수식 원문으로 선택하세요.');
  const regex=new RegExp(o.pattern,o.matchCase?'u':'iu'),groups=[];let count=0,totalText=0,removedText=0,addedText=0;
  for(const sheet of book.sheets)for(const ref in sheet.cells)totalText+=sheet.cells[ref].raw.length;
  for(const sheet of sheets(book,o)){
    const updates=[];
    for(const key in sheet.cells){const cell=sheet.cells[key];if(!inRange(key,o)||!regex.test(searchableRaw(cell.raw)))continue;const raw=replacementRaw(cell.raw,o);if(raw===cell.raw)continue;
      if(++count>100000)throw Error('모두 바꾸기는 한 번에 최대 100,000셀입니다. 검색 범위를 줄이세요.');
      if(sheet.pivot||preservedPivotTouches(book,sheet.id,key))throw Error(`${sheet.name}!${key}: 피벗 결과는 직접 바꿀 수 없습니다.`);
      validateCell(key,{...cell,raw});removedText+=cell.raw.length;addedText+=raw.length;
      if(addedText>MAX_DOCUMENT_TEXT_CHARS)throw Error('바꾸기 결과가 문서의 100,000,000자 한도를 초과합니다.');
      updates.push({key,cell:cloneData({...cell,raw})});
    }
    if(updates.length)groups.push({sheetId:sheet.id,updates});
  }
  if(totalText-removedText+addedText>MAX_DOCUMENT_TEXT_CHARS)throw Error('바꾸기 결과가 문서의 100,000,000자 한도를 초과합니다.');
  if(!count)return {book:null,count:0};
  const touched=new Set(groups.map(g=>g.sheetId)),hasTables=book.sheets.some(s=>s.tables?.length);
  const next={...book,sheets:book.sheets.map(source=>{
    const sheet={...source,cells:touched.has(source.id)||hasTables?{...source.cells}:source.cells};
    for(const field of ['tables','charts','conditionalRules','xlsxFeatures'])if(source[field]!==undefined)sheet[field]=cloneData(source[field]);
    // Table renames can change preservation baselines even when a current raw
    // formula no longer contains the old reference. Never mutate those on the source.
    if(hasTables)for(const key in sheet.cells){const cell=sheet.cells[key];if(cell.xlsxFormula||cell.cachedFormula||cell.xlsxExternal||cell.xlsxMetadata)sheet.cells[key]=cloneData(cell);}
    return sheet;
  }),...(book.chartNames?{chartNames:cloneData(book.chartNames)}:{})};next.referenceCachesValid=false;const applied=[];
  for(const {sheetId,updates} of groups){
    const sheet=next.sheets.find(s=>s.id===sheetId),plan=prepareTableHeaderEdits(next,sheet,updates);
    assertCellUpdates(next,sheet,plan.updates,{editAnchor:true});assertNativeCellTransfers(next,sheet,plan.updates);
    applied.push({sheet,plan,updates:plan.updates});
  }
  for(const {sheet,updates} of applied){
    for(const update of updates)if(update.cell.xlsxFormula&&sheet.cells[update.key]?.xlsxFormula)update.cell.xlsxFormula=sheet.cells[update.key].xlsxFormula;
    clearReplacedArrayDefinitions(sheet,updates,{editAnchor:true});
    for(const {key,cell} of updates){if(cell.xlsxFormula)cell.xlsxFormula=cloneData(cell.xlsxFormula);sheet.cells[key]=cell;}
  }
  for(const {sheet,plan} of applied)applyTableHeaderEdits(next,sheet,plan);
  for(const {sheet,updates} of applied){const invalid=validateCellUpdates(next,sheet,updates);if(invalid)throw Error(`${sheet.name}!${invalid.ref}: ${invalid.message}`);}
  validateWorkbook(next);return {book:next,count};
}

export function replacementPatch(before,after){
  const updates=[],sheets=[],metadata={};
  for(const next of after.sheets){
    const old=before.sheets.find(s=>s.id===next.id);if(!old)throw Error('바꾸기 대상 시트가 변경되었습니다.');
    for(const key in next.cells){const cell=next.cells[key],previous=old.cells[key];if(cell===previous)continue;if(previous?.raw!==cell.raw||['xlsxFormula','cachedFormula','xlsxExternal','xlsxMetadata'].some(field=>JSON.stringify(previous?.[field])!==JSON.stringify(cell[field])))updates.push({sheetId:next.id,key,cell});}
    const patch={};for(const field of ['tables','charts','conditionalRules','xlsxFeatures'])if(JSON.stringify(old[field])!==JSON.stringify(next[field]))patch[field]=next[field];
    if(Object.keys(patch).length)sheets.push({sheetId:next.id,metadata:patch});
  }
  for(const field of ['referenceCachesValid','chartNames'])if(JSON.stringify(before[field])!==JSON.stringify(after[field]))metadata[field]=after[field];
  return {updates,sheets,metadata};
}
export function applyReplacementPatch(book,patch){
  // Patches are produced by the worker's validated copy, never by a file.
  for(const update of patch.updates)if(!book.sheets.some(s=>s.id===update.sheetId))throw Error('바꾸기 대상 시트가 변경되었습니다.');
  for(const {sheetId,key,cell} of patch.updates)book.sheets.find(s=>s.id===sheetId).cells[key]=cell;
  for(const {sheetId,metadata} of patch.sheets)for(const [key,value] of Object.entries(metadata)){const sheet=book.sheets.find(s=>s.id===sheetId);if(value===undefined)delete sheet[key];else sheet[key]=value;}
  for(const [key,value] of Object.entries(patch.metadata))if(value===undefined)delete book[key];else book[key]=value;
}
