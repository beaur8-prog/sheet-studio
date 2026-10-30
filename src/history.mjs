import {plain, registerProxy, trustData} from './clone.mjs';
import {MAX_DOCUMENT_CELLS, MAX_DOCUMENT_TEXT_CHARS, MAX_HISTORY_BYTES} from './limits.mjs';
import {validateCell, validateWorkbookMetadata} from './engine.mjs';
import {estimateDataBytes} from './data-size.mjs';

const own = (value, key) => Object.hasOwn(value, key);
const index = key => typeof key === 'string' && /^(?:0|[1-9]\d*)$/.test(key) && +key < 4294967295;

// Keep references plus reversible property changes, rather than 60 entire books.
// Every model mutation, including nested formatting and array splices, goes
// through the journal. Large replacement operations share unchanged objects.
export function createHistory(initial, {maxEntries = 60, maxBytes = MAX_HISTORY_BYTES, maxCells = MAX_DOCUMENT_CELLS, maxText = MAX_DOCUMENT_TEXT_CHARS, description} = {}) {
  const holder = {value: description ? initial : plain(initial)}, undo = [], redo = [];
  let transaction = null, bytes = 0, replaying = false, dirty = new Map(), metadataChanged = false, metadataValidation = false;
  const counts = new WeakMap(), textCounts = new WeakMap(), sizes = new WeakMap(), containers = new WeakMap();
  const textLength = cells => { if(!textCounts.has(cells)) { let n=0;for(const ref in cells)n+=cells[ref]?.raw?.length||0;textCounts.set(cells,n); } return textCounts.get(cells); };
  function remember(book, description) {
    if (!description) return;
    trustData(book);
    sizes.set(book, description.bytes + 1024);
    for (const info of description.sheets) {
      const sheet = book.sheets.find(sheet => sheet.id === info.id);
      if (sheet) { counts.set(sheet.cells, info.count); if(info.textLength!==undefined)textCounts.set(sheet.cells,info.textLength);sizes.set(sheet.cells, info.bytes); }
    }
  }
  remember(holder.value, description);
  const count = cells => { if (!counts.has(cells)) counts.set(cells, Object.keys(cells).length); return counts.get(cells); };
  const totalCells = () => holder.value.sheets.reduce((n, s) => n + count(s.cells), 0);
  totalCells();
  for(const sheet of holder.value.sheets)textLength(sheet.cells);
  function marker(context, key) {
    return context?.dictionary ? {cells: context.cells, address: key} : context;
  }
  function touch(context) {
    if (!context?.cells || !context.address) { metadataChanged = true; metadataValidation = true; return; }
    if (!dirty.has(context.cells)) dirty.set(context.cells, new Set());
    dirty.get(context.cells).add(context.address);
  }
  function record(target, key, before, existed, after, exists, context) {
    touch(context);
    if (!transaction || replaying) return;
    let keys = transaction.targets.get(target);
    if (!keys) transaction.targets.set(target, keys = new Map());
    const previous = keys.get(key);
    if (previous) { previous.after = after; previous.exists = exists; return; }
    const op = {target, key, before, existed, after, exists, context};
    keys.set(key, op); transaction.ops.push(op);
  }
  function wrap(target, context = null) {
    if (!target || typeof target !== 'object') return target;
    // Sheets and workbook containers need stable identity for cut/move logic.
    // Cell descendants are intentionally uncached so reads do not retain a
    // million proxy objects beside the actual model.
    if (!context?.cells && containers.has(target)) return containers.get(target);
    const proxy=registerProxy(new Proxy(target, {
      get(object, key, receiver) {
        const value = Reflect.get(object, key, receiver);
        let child = context;
        if (key === 'cells' && object.id && object.cells === value) child = {cells: value, dictionary: true};
        else if (context?.dictionary) child = {cells: context.cells, address: key};
        return wrap(value, child);
      },
      set(object, key, value) {
        value = plain(value);
        const existed = own(object, key), before = object[key], mark = marker(context, key);
        if (existed && Object.is(before, value)) return true;
        if (Array.isArray(object)) {
          if (key === 'length' && value < object.length) {
            for (let i = value; i < object.length; i++) if (own(object, i)) record(object, String(i), object[i], true, undefined, false, mark);
          } else if (index(key) && +key >= object.length) {
            record(object, 'length', object.length, true, +key + 1, true, mark);
          }
        }
        if (context?.dictionary && !existed) counts.set(object, count(object) + 1);
        if(context?.dictionary)textCounts.set(object,textLength(object)+(value?.raw?.length||0)-(before?.raw?.length||0));
        else if(context?.cells&&key==='raw'&&context.cells[context.address]===object)textCounts.set(context.cells,textLength(context.cells)+(value?.length||0)-(before?.length||0));
        record(object, key, before, existed, value, true, mark);
        return Reflect.set(object, key, value);
      },
      deleteProperty(object, key) {
        if (!own(object, key)) return true;
        if (context?.dictionary) counts.set(object, count(object) - 1);
        if(context?.dictionary)textCounts.set(object,textLength(object)-(object[key]?.raw?.length||0));
        record(object, key, object[key], true, undefined, false, marker(context, key));
        return Reflect.deleteProperty(object, key);
      },
    }), target);
    if (!context?.cells) containers.set(target,proxy);
    return proxy;
  }
  function apply(op, reverse) {
    const exists = reverse ? op.existed : op.exists, value = reverse ? op.before : op.after;
    if (op.context?.cells === op.target) {
      counts.set(op.target, count(op.target) + Number(exists) - Number(own(op.target, op.key)));
      textCounts.set(op.target,textLength(op.target)+(exists?value?.raw?.length||0:0)-(op.target[op.key]?.raw?.length||0));
    }
    else if(op.context?.cells&&op.key==='raw'&&op.context.cells[op.context.address]===op.target)textCounts.set(op.context.cells,textLength(op.context.cells)+(exists?value?.length||0:0)-(op.target[op.key]?.length||0));
    if (exists) op.target[op.key] = value; else delete op.target[op.key];
    touch(op.context);
  }
  function rollback() {
    if (!transaction) return;
    replaying = true;
    for (const op of transaction.ops.toReversed()) apply(op, true);
    replaying = false; transaction = null;
  }
  function cost(ops) {
    const seen = new Set(); let size = ops.length * 96;
    for (const op of ops) for (const value of [op.before, op.after]) {
      if (value && typeof value === 'object') {
        if (seen.has(value)) continue;
        seen.add(value);
      }
      // A conservative UTF-16 payload budget, in addition to the entry limit.
      size += value && typeof value === 'object' && sizes.has(value) ? sizes.get(value) : estimateDataBytes(value,maxBytes-size);
      if (size > maxBytes) return size;
    }
    return size;
  }
  function commit() {
    // Imported unsupported formulas have no reliable dependency graph. A raw
    // value change invalidates their display caches, and the flag belongs to
    // the same undo entry as that edit. Loading a replacement workbook adopts
    // its own caches instead of immediately invalidating them.
    if(transaction&&!transaction.ops.some(op=>op.target===holder&&op.key==='value')&&transaction.ops.some(op=>op.context?.cells&&(op.key==='raw'&&!Object.is(op.before,op.after)||op.target===op.context.cells&&(op.before?.raw||'')!==(op.after?.raw||''))))wrap(holder.value).referenceCachesValid=false;
    try {
      if (totalCells() > maxCells) throw Error('통합 문서에는 최대 1,000,000개의 셀을 기록할 수 있습니다. 변경을 취소했습니다.');
      if(holder.value.sheets.reduce((n,s)=>n+textLength(s.cells),0)>maxText)throw Error('문서의 셀 텍스트 총량이 제한을 초과했습니다. 변경을 취소했습니다.');
      const current = new Set(holder.value.sheets.map(sheet => sheet.cells));
      for (const [cells, refs] of dirty) if (current.has(cells)) for (const ref of refs) if (own(cells,ref)) validateCell(ref,cells[ref]);
      if (metadataValidation) validateWorkbookMetadata(holder.value);
      metadataValidation = false;
    } catch (error) {
      rollback(); throw error;
    }
    if (!transaction) return {discarded: false};
    const ops = transaction.ops.filter(op => op.existed !== op.exists || !Object.is(op.before, op.after));
    transaction = null;
    if (!ops.length) return {discarded: false};
    for (const entry of redo) bytes -= entry.bytes;
    redo.length = 0;
    const entry = {ops, bytes: cost(ops)};
    if (!ops.some(op => op.target === holder && op.key === 'value') && sizes.has(holder.value)) sizes.set(holder.value, sizes.get(holder.value) + entry.bytes);
    for(const cells of new Set(ops.map(op=>op.context?.cells).filter(Boolean)))if(sizes.has(cells))sizes.set(cells,sizes.get(cells)+entry.bytes);
    if (entry.bytes > maxBytes) {
      undo.length = 0; bytes = 0;
      return {discarded: true};
    }
    undo.push(entry); bytes += entry.bytes;
    while (undo.length > maxEntries || bytes > maxBytes) bytes -= undo.shift().bytes;
    return {discarded: false};
  }
  return {
    get book() { return wrap(holder.value); },
    get raw() { return holder.value; },
    get canUndo() { return undo.length > 0; },
    get canRedo() { return redo.length > 0; },
    get stats() { return {undo: undo.length, redo: redo.length, bytes, cells: totalCells(), textChars:holder.value.sheets.reduce((n,s)=>n+textLength(s.cells),0)}; },
    begin() { commit(); transaction = {ops: [], targets: new WeakMap()}; },
    replace(value, description) {
      if (!description) { wrap(holder).value = value; return wrap(holder.value); }
      const before = holder.value;
      record(holder, 'value', before, true, value, true, null);
      holder.value = value; remember(value, description); return wrap(value);
    },
    commit, rollback,
    takeDirty() { const result = dirty; result.metadataChanged=metadataChanged; dirty = new Map(); metadataChanged=false; return result; },
    acceptCells(cells, description) { counts.set(cells, description.count); if(description.textLength!==undefined)textCounts.set(cells,description.textLength);sizes.set(cells, description.bytes); return trustData(cells); },
    undo(reverse = false) {
      commit(); const from = reverse ? redo : undo, to = reverse ? undo : redo;
      if (!from.length) return false;
      const entry = from.pop(); replaying = true;
      for (const op of reverse ? entry.ops : entry.ops.toReversed()) apply(op, !reverse);
      replaying = false; to.push(entry); return true;
    },
  };
}
