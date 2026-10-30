import {createTablePlan,updateTablePlan,calculatedColumnPlan,finalizeTablePlan} from './tables.mjs';
import {tableConversionPlan} from './table-conversion.mjs';
import {tableAt,tableBounds} from './table-schema.mjs';
import {assertSortAllowed,validationChoices,resetConstraintIndexes} from './cell-constraints.mjs';
import {currentRegion,inferSortPlan} from './data-region.mjs';
import {createWorkbook, parseWorkbook, serializeWorkbook, validateWorkbook, validateWorkbookMetadata, validateCell, createEvaluator, parseAddress,formulaReference} from './engine.mjs';
import {createConditionalEvaluator} from './conditional.mjs';
import {estimateDataBytes} from './data-size.mjs';
import {MAX_HISTORY_BYTES} from './limits.mjs';
import {readSpreadsheet, writeSpreadsheet} from './io.mjs';
import {openRecoveryStore, workbookMetadata} from './recovery-store.mjs';
import {applyFilters, filterOptions, filterPredicates, valueKey, sortFilterRange, sortRange, inferFilterRange} from './filters.mjs';

let book = createWorkbook(), revision = 0, store, prepared = null, evaluators = new Map(), conditionals = new Map();
let dirty = new Map(), replacements = new Set(), needsSave = false, metadataDirty = false;
function changed() { resetConstraintIndexes();evaluators = new Map(); conditionals = new Map(); needsSave = true; }
const evaluator = id => {
  if (!evaluators.has(id)) evaluators.set(id, createEvaluator(book, id));
  return evaluators.get(id);
};
const cellDescription = cells => {
  let textLength=0,count=0,bytes=4;
  for(const ref in cells){
    count++;textLength+=cells[ref].raw.length;
    if(bytes<=MAX_HISTORY_BYTES)bytes+=estimateDataBytes(cells[ref],MAX_HISTORY_BYTES-bytes)+(ref.length+4)*2;
  }
  return {count,textLength,bytes:Math.min(bytes,MAX_HISTORY_BYTES+1)};
};
const describe = book => {
  const sheets=book.sheets.map(sheet=>({id:sheet.id,...cellDescription(sheet.cells)}));
  return {bytes:Math.min(MAX_HISTORY_BYTES+1,estimateDataBytes(workbookMetadata(book),MAX_HISTORY_BYTES)+sheets.reduce((n,s)=>n+s.bytes+32,0)),sheets};
};
async function handle(type, data) {
  if (type === 'load') {
    store = await openRecoveryStore();
    const recovered = await store.load();
    if (recovered) book = validateWorkbook(recovered);
    else if (data.legacy) {
      book = parseWorkbook(data.legacy);
      await store.save(book, new Map(), new Set(book.sheets.map(sheet => sheet.id)));
    }
    return {book, description: describe(book), migrated: !!data.legacy, revision};
  }
  if (type === 'import') {
    if (data.bytes.byteLength > 100 * 1024 * 1024) throw Error('100 MB 이하의 파일을 선택하세요.');
    const loaded = /\.json$/i.test(data.name) ? parseWorkbook(new TextDecoder().decode(data.bytes)) : validateWorkbook(await readSpreadsheet(data.bytes, data.name));
    const token = crypto.randomUUID(); prepared = {token, book: loaded};
    return {book: loaded, description: describe(loaded), token};
  }
  if (type === 'sync') {
    const metadata=data.metadata||workbookMetadata(book);
    metadataDirty ||= !!data.metadata;
    if(data.metadata)validateWorkbookMetadata(metadata);
    for(const update of data.sheets)for(const [ref,cell] of update.entries||[])if(cell!==null)validateCell(ref,cell);
    const old = new Map(book.sheets.map(sheet => [sheet.id, sheet.cells]));
    if (data.token) {
      if (prepared?.token !== data.token) throw Error('파일 가져오기 작업이 만료되었습니다. 다시 열어 주세요.');
      for (const sheet of prepared.book.sheets) { if (old.get(sheet.id) !== sheet.cells) replacements.add(sheet.id); old.set(sheet.id, sheet.cells); }
      prepared = null;
    }
    const updates = new Map(data.sheets.map(sheet => [sheet.id, sheet]));
    if(data.sheets.some(sheet=>sheet.cells))validateWorkbook({...metadata,sheets:metadata.sheets.map(sheet=>({...sheet,cells:updates.get(sheet.id)?.cells||old.get(sheet.id)||{}}))});
    const next = {...metadata, sheets: metadata.sheets.map(metadata => {
      const update = updates.get(metadata.id);
      const cells = update?.cells || old.get(metadata.id) || {};
      if (update?.cells) replacements.add(metadata.id);
      if (update?.entries) {
        if (!dirty.has(metadata.id)) dirty.set(metadata.id, new Set());
        for (const [ref, cell] of update.entries) {
          if (cell === null) delete cells[ref]; else cells[ref] = cell;
          dirty.get(metadata.id).add(ref);
        }
      }
      return {...metadata, cells};
    })};
    book=next;
    revision = data.revision; changed(); return {revision};
  }
  if (type === 'flush') {
    if (!needsSave) return {revision, chunksWritten: 0, fullSheets: 0, metadataWritten: false};
    const result = await store.save(book, dirty, replacements, metadataDirty);
    dirty = new Map(); replacements = new Set(); needsSave = false; metadataDirty=false;
    return {...result, revision};
  }
  if (data.revision !== revision) throw Error('문서가 변경되었습니다. 작업을 다시 실행하세요.');
  if (type === 'recalculate') { evaluators=new Map();conditionals=new Map(); }
  if (type === 'calculate' || type === 'recalculate') return {revision, values: data.refs.map(([id, ref]) => {
    const sheet=book.sheets.find(sheet=>sheet.id===id),evaluate=evaluator(id),result=evaluate.result(ref),value=result.value,position=parseAddress(ref);
    if(!conditionals.has(id))conditionals.set(id,sheet?createConditionalEvaluator(book,sheet,evaluate):()=>({}));
    return [id, ref, value, position?conditionals.get(id)(position.r,position.c,value):{},result.error];
  })};
  if (type === 'export') {
    if (data.format === 'json') return {data: new TextEncoder().encode(serializeWorkbook(book)), mime: 'application/json', extension: 'sheet.json'};
    const snapshots=new WeakMap(),get=(snapshot,id)=>{
      if(snapshot===book)return evaluator(id);
      // Pivot export evaluates its refreshed snapshot rather than mixing those
      // cells with caches from the unrefreshed document.
      if(!snapshots.has(snapshot))snapshots.set(snapshot,new Map());
      const cache=snapshots.get(snapshot);if(!cache.has(id))cache.set(id,createEvaluator(snapshot,id));
      return cache.get(id);
    };
    return writeSpreadsheet(book,data.format,{...data.options,evaluate:(snapshot,id,ref)=>get(snapshot,id)(ref),evaluateDetailed:(snapshot,id,ref)=>get(snapshot,id).result(ref)});
  }
  const sheet = book.sheets.find(sheet => sheet.id === data.sheetId);
  if (!sheet) throw Error('시트를 찾을 수 없습니다.');
  if(type==='table-plan'){const table=sheet.tables?.find(t=>t.id===data.tableId);if(data.operation!=='create'&&!table)throw Error('표를 찾을 수 없습니다.');if(data.operation==='convert')return tableConversionPlan(book,sheet,table);const plan=data.operation==='create'?createTablePlan(book,sheet,data.values):data.operation==='calculate'?calculatedColumnPlan(book,sheet,table,data.values):updateTablePlan(book,sheet,table,data.values);return finalizeTablePlan(book,sheet,plan,data.values);}
  if (type === 'name-range') return formulaReference(book,sheet.id,data.name);
  if (type === 'validation-options') return validationChoices(book,sheet,data.ref);
  if (type === 'infer-region'){const table=tableAt(sheet,data.selection.r1,data.selection.c1);if(table){const b=tableBounds(table.range);return {...b,r2:b.r2-(table.totalsRow?1:0)};}return currentRegion(sheet,data.selection);}
  if (type === 'sort-plan') return inferSortPlan(book,sheet,data.selection,data.column);
  if (type === 'infer-filter') return inferFilterRange(sheet, data.selection);
  if (type === 'filter-options') {
    const table=data.tableId?sheet.tables?.find(t=>t.id===data.tableId):null,view=table?{...sheet,autoFilter:table.filter}:sheet;
    if(!view.autoFilter)throw Error('필터를 찾을 수 없습니다.');
    const options = filterOptions(book,view,data.column),rule=view.autoFilter.columns[data.column];
    const own=rule?.type!=='color'&&filterPredicates(book,view,evaluator(sheet.id)).find(predicate=>predicate.column===data.column);
    const original=rule?.type==='values'?rule.values.map(valueKey):options.values.filter(value=>!own||own.test(value.value)).map(value=>value.key);
    return {options,original};
  }
  if (type === 'filter') {
    const table=data.tableId?sheet.tables?.find(t=>t.id===data.tableId):null,filter=structuredClone(table?.filter||sheet.autoFilter);if(!filter)throw Error('필터를 찾을 수 없습니다.');
    if(data.column!==undefined){if(data.rule)filter.columns[data.column]=data.rule;else delete filter.columns[data.column];}
    const target=table?{...sheet,tables:sheet.tables.map(t=>t.id===table.id?{...t,filter}:t)}:{...sheet,autoFilter:filter};
    applyFilters({...book,sheets:book.sheets.map(item=>item.id===sheet.id?target:item)},{...target,autoFilter:filter});return filter;
  }
  if (type === 'sort-filter' || type === 'sort') {
    const table=data.tableId?sheet.tables?.find(t=>t.id===data.tableId):null;
    const range=type==='sort-filter'?{...(table?.filter||sheet.autoFilter).range,r1:(table?.filter||sheet.autoFilter).range.r1+1}:data.range;assertSortAllowed(book,sheet,range);
    const target = structuredClone(sheet);
    const next = {...book, sheets: book.sheets.map(item => item.id === sheet.id ? target : item)};
    if(type==='sort-filter'){const filter=table?target.tables.find(t=>t.id===data.tableId).filter:target.autoFilter;sortFilterRange(next,{...target,autoFilter:filter},data.column,data.desc);}
    else sortRange(next, target, data.range, data.column, data.desc);
    validateWorkbook(next);
    const token=crypto.randomUUID();prepared={token,book:next};
    return {cells: target.cells, autoFilter: target.autoFilter,...(table?{tableId:table.id,tableFilter:target.tables.find(t=>t.id===table.id).filter}:{}), token, description:cellDescription(target.cells)};
  }
  throw Error('알 수 없는 문서 작업입니다.');
}

let queue = Promise.resolve();
self.onmessage = ({data: {id, type, payload}}) => {
  queue = queue.then(async () => {
    try {
      const value = await handle(type, payload || {});
      const transfers = value?.data instanceof Uint8Array ? [value.data.buffer] : [];
      self.postMessage({id, value}, transfers);
    } catch (error) { self.postMessage({id, error: error.message || String(error)}); }
  });
};
