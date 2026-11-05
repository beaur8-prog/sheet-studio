import test from 'node:test';
import assert from 'node:assert/strict';
import {abortable,importTask} from '../src/import-task.mjs';
function harness({hold=false,fail=false,readFail=false}={}){
  const state={created:0,terminated:0,discarded:[],read:0,prepared:null,events:[]};
  const options={
    prepare:async(token,port)=>{state.events.push('prepare');port.onmessage=({data})=>{state.prepared=data;port.postMessage({token,staged:true});};},
    createWorker:()=>{state.created++;const worker={terminate(){state.terminated++;},postMessage({token,port}){state.events.push('parse');if(hold)return;queueMicrotask(()=>{worker.onmessage({data:{token,progress:{phase:'decode'}}});if(fail){worker.onmessage({data:{token,error:'broken file'}});return;}port.onmessage=()=>worker.onmessage({data:{token,staged:true}});port.postMessage({token,book:{name:'candidate'},description:{sheets:[]}});});}};return worker;},
    onProgress:progress=>state.events.push(progress.phase),onReady:()=>state.events.push('lock'),
    read:async token=>{state.read++;state.events.push('read');if(readFail)throw Error('preview failed');return {...state.prepared,token};},
    discard:async token=>{state.discarded.push(token);if(state.prepared?.token===token)state.prepared=null;}
  };return {state,options};
}
test('a successful import terminates its parser before the candidate is read and leaves adoption to the document worker',async()=>{
  const {state,options}=harness(),result=await importTask(new Uint8Array([1]),'test.xlsx',options);assert.equal(result.book.name,'candidate');assert.equal(state.created,1);assert.equal(state.terminated,1);assert.equal(state.read,1);assert.deepEqual(state.discarded,[]);assert.deepEqual(state.events,['prepare','parse','decode','lock','read']);
});
test('canceling a running parser discards its candidate, skips adoption and releases the caller promptly',async()=>{
  const {state,options}=harness({hold:true}),controller=new AbortController();const job=importTask(new Uint8Array([1]),'test.xlsx',{...options,signal:controller.signal});await new Promise(resolve=>setTimeout(resolve,1));const start=performance.now();controller.abort();await assert.rejects(job,{name:'AbortError'});assert.ok(performance.now()-start<500);assert.equal(state.terminated,1);assert.equal(state.read,0);assert.equal(state.discarded.length,1);
});
test('cancel can return while the document worker is busy, with cleanup queued behind it',async()=>{
  const {options,state}=harness(),controller=new AbortController();let clean=false;
  const job=importTask(new Uint8Array([1]),'test.xlsx',{...options,signal:controller.signal,prepare:()=>new Promise(()=>{}),discard:()=>{clean=true;return new Promise(()=>{});}});controller.abort();await assert.rejects(job,{name:'AbortError'});assert.equal(clean,true);assert.equal(state.created,0);
});
test('parse or preview failures discard staged state and allow the current document to remain untouched',async()=>{
  for(const config of [{fail:true},{readFail:true}]){const {state,options}=harness(config);await assert.rejects(importTask(new Uint8Array([1]),'test.xlsx',options),/broken file|preview failed/);assert.equal(state.terminated,1);assert.equal(state.discarded.length,1);assert.equal(state.prepared,null);}
});
test('an already canceled operation never creates a parser or stages a document; abortable consumes late file-read failures',async()=>{
  const {state,options}=harness(),controller=new AbortController();controller.abort();await assert.rejects(importTask(new Uint8Array([1]),'test.xlsx',{...options,signal:controller.signal}),{name:'AbortError'});assert.equal(state.created,0);assert.deepEqual(state.events,[]);
  const next=new AbortController();let fail;const job=abortable(new Promise((_r,reject)=>{fail=reject;}),next.signal);next.abort();await assert.rejects(job,{name:'AbortError'});fail(Error('late read error'));const already=new AbortController();already.abort();let rejectAlready;const late=new Promise((_r,reject)=>{rejectAlready=reject;});await assert.rejects(abortable(late,already.signal),{name:'AbortError'});rejectAlready(Error('already-canceled read error'));await new Promise(resolve=>setTimeout(resolve,1));
});
