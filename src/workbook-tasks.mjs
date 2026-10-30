import {workbookMetadata} from './recovery-store.mjs';

export function createWorkbookTasks({onSave, onCalculated, onError}) {
  const worker = new Worker(new URL('./workbook-worker.mjs', import.meta.url), {type: 'module'});
  let sequence = 0, revision = 0, calculationEpoch = 0, synced = Promise.resolve(), flushTimer, saveRevision = 0, failed = null;
  let baseline = new Map(), values = new Map(), styles = new Map(), calculationErrors=new Map(),provenance=new Map(),pendingRefs = new Map(), calculateScheduled = false;
  const pending = new Map();
  worker.onmessage = ({data}) => {
    const item = pending.get(data.id);
    if (!item) return;
    pending.delete(data.id);
    if (data.error) item.reject(Error(data.error)); else item.resolve(data.value);
  };
  worker.onerror = () => {
    failed = Error('문서 작업 스레드를 사용할 수 없습니다. 앱을 다시 실행해 주세요.');
    for (const item of pending.values()) item.reject(failed);
    pending.clear(); onError(failed);
  };
  const call = (type, payload = {}, transfer = []) => new Promise((resolve, reject) => {
    if (failed) return reject(failed);
    const id = ++sequence; pending.set(id, {resolve, reject});
    try { worker.postMessage({id, type, payload}, transfer); } catch (error) { pending.delete(id); reject(error); }
  });
  const adopt = book => { baseline = new Map(book.sheets.map(sheet => [sheet.id, sheet.cells])); };
  async function flush() {
    clearTimeout(flushTimer);
    await synced;
    const result = await call('flush');
    saveRevision = result.revision;
    if (saveRevision === revision) onSave('saved');
    return result;
  }
  function sync(book, dirty, token) {
    const sheets = [], metadata = !baseline.size || dirty.metadataChanged || token ? workbookMetadata(book) : null;
    for (const sheet of book.sheets) {
      if (baseline.get(sheet.id) !== sheet.cells && !token) sheets.push({id: sheet.id, cells: sheet.cells});
      else if (dirty.has(sheet.cells)) sheets.push({id: sheet.id, entries: [...dirty.get(sheet.cells)].map(ref => [ref, sheet.cells[ref] || null])});
    }
    adopt(book); values = new Map(); styles = new Map();calculationErrors=new Map();provenance=new Map(); pendingRefs.clear(); ++revision; ++calculationEpoch;
    onSave('saving');
    synced = call('sync', {metadata, sheets, revision, token});
    synced.catch(error => { baseline=new Map();onSave('failed'); onError(error); });
    clearTimeout(flushTimer); flushTimer = setTimeout(() => flush().catch(error => { onSave('failed'); onError(error); }), 100);
    return synced;
  }
  function value(sheetId, ref) {
    const key = sheetId + '!' + ref;
    if (values.has(key)) return values.get(key);
    pendingRefs.set(key, [sheetId, ref]);
    if (!calculateScheduled) {
      calculateScheduled = true;
      queueMicrotask(async () => {
        calculateScheduled = false;
        const requestedRevision = revision, requestedEpoch = calculationEpoch, refs = [...pendingRefs.values()]; pendingRefs.clear();
        if (!refs.length) return;
        try {
          await synced;
          if (requestedRevision !== revision || requestedEpoch !== calculationEpoch) return;
          const result = await call('calculate', {refs, revision: requestedRevision,includeArrays:true});
          if (result.revision !== revision || requestedEpoch !== calculationEpoch) return;
          for (const [id, address, computed, style,error,origin] of result.values) {const key=id+'!'+address; values.set(key, computed); styles.set(key, style);calculationErrors.set(key,error||null);if(origin)provenance.set(key,origin); }
          onCalculated(result.arrays);
        } catch (error) { if (requestedRevision === revision && requestedEpoch === calculationEpoch) onError(error); }
      });
    }
    return '계산 중…';
  }
  async function recalculate(){
    const refs=new Map(pendingRefs);
    for(const key of values.keys()){
      const split=key.lastIndexOf('!');refs.set(key,[key.slice(0,split),key.slice(split+1)]);
    }
    const requestedEpoch=++calculationEpoch;
    values=new Map();styles=new Map();calculationErrors=new Map();provenance=new Map();pendingRefs.clear();
    // Queue this after already-posted edits and before subsequent edits. Do not
    // change the document revision merely because volatile formulas recalculate.
    const request=call('recalculate',{refs:[...refs.values()],revision,includeArrays:true});
    const ready=Promise.all([synced,request]);synced=ready.then(()=>undefined);
    const [,result]=await ready;
    if(result.revision!==revision||requestedEpoch!==calculationEpoch)return;
    for(const [id,ref,value,style,error,origin] of result.values){const key=id+'!'+ref;values.set(key,value);styles.set(key,style);calculationErrors.set(key,error||null);if(origin)provenance.set(key,origin);}
    onCalculated(result.arrays);
    return result;
  }
  return {
    get revision() { return revision; },
    get saving() { return saveRevision !== revision; },
    async load(legacy) { const result = await call('load', {legacy}); adopt(result.book); return result; },
    async import(bytes, name) { return call('import', {bytes, name}, [bytes.buffer]); },
    sync, flush, value, recalculate,
    result(sheetId,ref){return {value:value(sheetId,ref),error:calculationErrors.get(sheetId+'!'+ref)||null,...provenance.get(sheetId+'!'+ref)};},
    style(sheetId,ref) { const key=sheetId+'!'+ref;if(styles.has(key))return styles.get(key);value(sheetId,ref);return {}; },
    async run(type, payload = {}) { await synced; return call(type, {...payload, revision}); },
    destroy() { clearTimeout(flushTimer); worker.terminate(); },
  };
}
