import {importTask,abortable} from './import-task.mjs';
import {cellLiteralResult} from './engine.mjs';
import {workbookMetadata} from './recovery-store.mjs';

export function createWorkbookTasks({onSave, onCalculated, onError,onFatal,isComputed=()=>false}) {
  const worker = new Worker(new URL('./workbook-worker.mjs', import.meta.url), {type: 'module'});
  let sequence = 0, revision = 0, calculationEpoch = 0, synced = Promise.resolve(), flushTimer, saveRevision = 0, failed = null;
  let baseline = new Map(), values = new Map(), styles = new Map(), calculationErrors=new Map(),provenance=new Map(),pendingRefs = new Map(), calculateScheduled = false,calculating=0;
  const pending = new Map();let activeImport=null;
  let currentBook=null,needsArrays=true,presentationEnabled=false;
  const presentation=new Map(),presentationLimit=20000;
  function rememberPresentation(result){
    if(!presentationEnabled)return;
    for(const [id,ref,value,style,error,origin] of result.values){
      const key=id+'!'+ref,raw=baseline.get(id)?.[ref]?.raw??'';
      presentation.delete(key);presentation.set(key,{raw,result:{value,error:error||null,...origin},style,computed:isComputed(id,ref)});
      if(presentation.size>presentationLimit)presentation.delete(presentation.keys().next().value);
    }
  }
  function previousPresentation(id,ref){
    const last=presentation.get(id+'!'+ref);
    return last&&last.raw===(baseline.get(id)?.[ref]?.raw??'')&&(!last.result.cached||currentBook?.referenceCachesValid!==false)?last:null;
  }
  worker.onmessage = ({data}) => {
    const item = pending.get(data.id);
    if (!item) return;
    pending.delete(data.id);
    if (data.error) item.reject(Error(data.error)); else item.resolve(data.value);
  };
  worker.onerror = () => {
    failed = Error('문서 작업 스레드를 사용할 수 없습니다. 앱을 다시 실행해 주세요.');
    for (const item of pending.values()) item.reject(failed);
    pending.clear();activeImport?.abort(failed);onFatal?.(failed);onError(failed);
  };
  const call = (type, payload = {}, transfer = []) => new Promise((resolve, reject) => {
    if (failed) return reject(failed);
    const id = ++sequence; pending.set(id, {resolve, reject});
    try { worker.postMessage({id, type, payload}, transfer); } catch (error) { pending.delete(id); reject(error); }
  });
  const adopt = book => { currentBook=book;baseline = new Map(book.sheets.map(sheet => [sheet.id, sheet.cells])); };
  async function flush() {
    clearTimeout(flushTimer);
    await synced;
    const result = await call('flush');
    saveRevision = result.revision;
    if (saveRevision === revision) onSave('saved');
    return result;
  }
  function sync(book, dirty, token, documentId, {persist = true} = {}) {
    const sheets = [], metadata = !baseline.size || dirty.metadataChanged || token ? workbookMetadata(book) : null;
    for (const sheet of book.sheets) {
      if (baseline.get(sheet.id) !== sheet.cells && !token) sheets.push({id: sheet.id, cells: sheet.cells});
      else if (dirty.has(sheet.cells)) sheets.push({id: sheet.id, entries: [...dirty.get(sheet.cells)].map(ref => [ref, sheet.cells[ref] || null])});
    }
    if(token||(currentBook&&currentBook!==book))presentation.clear();
    else if(presentation.size){const ids=new Set();for(const [id,cells] of baseline)if(!book.sheets.some(sheet=>sheet.id===id&&sheet.cells===cells))ids.add(id);if(ids.size)for(const key of presentation.keys())if(ids.has(key.slice(0,key.lastIndexOf('!'))))presentation.delete(key);}
    needsArrays=true;
    adopt(book); values = new Map(); styles = new Map();calculationErrors=new Map();provenance=new Map(); pendingRefs.clear(); ++revision; ++calculationEpoch;
    if(persist)onSave('saving');
    synced = call('sync', {metadata, sheets, revision, token, documentId, persist});
    synced.catch(error => { baseline=new Map();onSave('failed'); onError(error); });
    clearTimeout(flushTimer);
    if(persist)flushTimer = setTimeout(() => flush().catch(error => { onSave('failed'); onError(error); }), 100);
    else saveRevision=revision;
    return synced;
  }
  function scheduleCalculation(){
    if(calculateScheduled)return;
    calculateScheduled=true;
    queueMicrotask(async()=>{
      calculateScheduled=false;
      const requestedRevision=revision,requestedEpoch=calculationEpoch,refs=[...pendingRefs.values()],discover=needsArrays;
      pendingRefs.clear();needsArrays=false;
      if(!refs.length&&!discover)return;
      calculating++;
      try{
        await synced;
        if(requestedRevision!==revision||requestedEpoch!==calculationEpoch)return;
        const result=await call('calculate',{refs,revision:requestedRevision,includeArrays:true});
        if(result.revision!==revision||requestedEpoch!==calculationEpoch)return;
        for(const [id,address,computed,style,error,origin] of result.values){const key=id+'!'+address;values.set(key,computed);styles.set(key,style);calculationErrors.set(key,error||null);if(origin)provenance.set(key,origin);}
        onCalculated(result.arrays);rememberPresentation(result);
      }catch(error){if(requestedRevision===revision&&requestedEpoch===calculationEpoch)onError(error);}
      finally{calculating--;}
    });
  }
  function value(sheetId,ref){
    const key=sheetId+'!'+ref;
    if(values.has(key))return values.get(key);
    pendingRefs.set(key,[sheetId,ref]);scheduleCalculation();return '계산 중…';
  }
  function result(sheetId,ref){const key=sheetId+'!'+ref,valueNow=value(sheetId,ref);return {value:valueNow,error:calculationErrors.get(key)||null,...provenance.get(key),pending:!values.has(key)};}
  // Presentation only: export, copy, validation and calculations still use the
  // worker's current revision. A pending preview is never a computed result.
  function displayResult(sheetId,ref){
    presentationEnabled=true;
    const raw=baseline.get(sheetId)?.[ref]?.raw??'',last=previousPresentation(sheetId,ref);
    if(!(typeof raw==='string'&&raw.startsWith('='))&&!isComputed(sheetId,ref)&&!last?.computed){
      if(needsArrays)scheduleCalculation();return {...cellLiteralResult(raw),pending:false};
    }
    const current=result(sheetId,ref);
    if(!current.pending)return current;
    if(last)return {...last.result,pending:true};
    const cell=baseline.get(sheetId)?.[ref];
    if(currentBook?.referenceCachesValid!==false&&cell?.cachedValue!==undefined&&(!cell.cachedFormula||cell.cachedFormula===raw))return {value:cell.cachedValue,error:cell.cachedError?String(cell.cachedValue):null,cachedValue:cell.cachedValue,cachedError:cell.cachedError?String(cell.cachedValue):null,cached:true,stale:false,pending:true};
    return {value:'',error:null,pending:true};
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
    onCalculated(result.arrays);rememberPresentation(result);
    return result;
  }
  return {
    get revision() { return revision; },
    get calculated() { return !calculateScheduled&&!pendingRefs.size&&!calculating; },
    get saving() { return saveRevision !== revision; },
    async load(legacy, options = {}) { const result = await call('load', {legacy, ...options}); presentation.clear();needsArrays=true;adopt(result.book); return result; },
    async import(bytes,name,options={}){const controller=new AbortController(),cancel=()=>controller.abort(options.signal?.reason);if(options.signal?.aborted)cancel();else options.signal?.addEventListener('abort',cancel,{once:true});activeImport=controller;try{await abortable(synced,controller.signal);return await importTask(bytes,name,{...options,signal:controller.signal,createWorker:()=>new Worker(new URL('./workbook-import-worker.mjs',import.meta.url),{type:'module'}),prepare:(token,port)=>call('import-port',{token,port},[port]),read:token=>call('import-preview',{token}),discard:token=>call('import-discard',{token}).catch(()=>{})});}finally{options.signal?.removeEventListener('abort',cancel);if(activeImport===controller)activeImport=null;}},
    discardImport:token=>call('import-discard',{token}).catch(()=>{}),
    sync, flush, value, recalculate, displayResult,
    result,
    style(sheetId,ref) { const key=sheetId+'!'+ref;if(styles.has(key))return styles.get(key);value(sheetId,ref);return {}; },
    displayStyle(sheetId,ref){presentationEnabled=true;const key=sheetId+'!'+ref;if(styles.has(key))return styles.get(key);value(sheetId,ref);return previousPresentation(sheetId,ref)?.style||{};},
    async run(type, payload = {}) { await synced; return call(type, {...payload, revision}); },
    destroy() { clearTimeout(flushTimer);activeImport?.abort(Error('문서 작업을 종료했습니다.'));worker.terminate(); },
  };
}
