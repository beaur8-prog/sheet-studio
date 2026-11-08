import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator,cellLiteralResult} from '../src/engine.mjs';
import {createWorkbookTasks} from '../src/workbook-tasks.mjs';
const tick=()=>new Promise(resolve=>setTimeout(resolve,0));
async function fixture(cells,check){
 const original=globalThis.Worker,messages=[],book=createWorkbook();book.sheets[0].cells=Object.fromEntries(Object.entries(cells).map(([ref,raw])=>[ref,{raw}]));
 let worker;
 globalThis.Worker=class{constructor(){worker=this;}postMessage(message){if(message.type==='calculate')messages.push(message);else queueMicrotask(()=>this.onmessage({data:{id:message.id,value:message.type==='load'?{book,documentId:'doc'}:{revision:message.payload.revision}}}));}terminate(){}};
 let tasks;try{
  const computed=new Set();tasks=createWorkbookTasks({onSave(){},onCalculated(){},onError(error){throw error;},isComputed:(_id,ref)=>computed.has(ref)});await tasks.load('',{fresh:true});
  const reply=(message,values)=>worker.onmessage({data:{id:message.id,value:{revision:message.payload.revision,values:values.map(([ref,value,style={},error=null,origin=null])=>[book.sheets[0].id,ref,value,style,error,origin]),arrays:[]}}});
  await check({tasks,book,id:book.sheets[0].id,messages,reply,computed});
 }finally{tasks?.destroy();if(original===undefined)delete globalThis.Worker;else globalThis.Worker=original;}
}
test('synchronous literal presentation preserves evaluator types, text escapes and errors',()=>{
 const book=createWorkbook(),id=book.sheets[0].id;
 for(const raw of ['',null,undefined,'123','1e3','25%','TRUE','FALSE',"'123","'#N/A",'#N/A','1e999','TRUE ','한글']){book.sheets[0].cells.A1={raw};assert.deepEqual(cellLiteralResult(raw),createEvaluator(book,id).result('A1'),String(raw));}
});
test('literals need no cell calculation while formula previews never become authoritative values',async()=>fixture({A1:'1',B1:'=A1*2'},async({tasks,book,id,messages,reply})=>{
 assert.deepEqual(tasks.displayResult(id,'A1'),{value:1,error:null,pending:false});assert.equal(tasks.displayResult(id,'Z10000').value,'');
 assert.equal(tasks.displayResult(id,'B1').pending,true);await tick();assert.deepEqual(messages[0].payload.refs,[[id,'B1']]);reply(messages.shift(),[['B1',2,{color:'#ff0000'}]]);await tick();assert.equal(tasks.displayResult(id,'B1').value,2);
 book.sheets[0].cells.A1.raw='10';await tasks.sync(book,new Map([[book.sheets[0].cells,new Set(['A1'])]]),undefined,'doc',{persist:false});
 assert.equal(tasks.displayResult(id,'A1').value,10);assert.deepEqual(tasks.displayResult(id,'B1'),{value:2,error:null,pending:true});assert.equal(tasks.result(id,'B1').value,'계산 중…');assert.deepEqual(tasks.displayStyle(id,'B1'),{color:'#ff0000'});
 await tick();reply(messages.shift(),[['B1',20]]);await tick();assert.equal(tasks.displayResult(id,'B1').value,20);assert.equal(tasks.displayResult(id,'B1').pending,false);
}));
test('old revisions and other documents cannot populate the presentation cache',async()=>fixture({A1:'1',B1:'=A1*2'},async({tasks,book,id,messages,reply})=>{
 tasks.displayResult(id,'B1');await tick();const old=messages.shift();book.sheets[0].cells.A1.raw='10';await tasks.sync(book,new Map([[book.sheets[0].cells,new Set(['A1'])]]),undefined,'doc',{persist:false});tasks.displayResult(id,'B1');await tick();const fresh=messages.shift();reply(old,[['B1',999]]);await tick();assert.equal(tasks.displayResult(id,'B1').value,'');reply(fresh,[['B1',20]]);await tick();assert.equal(tasks.displayResult(id,'B1').value,20);
 const other=structuredClone(book);other.sheets[0].cells.A1.raw='3';await tasks.sync(other,new Map(),undefined,'other',{persist:false});assert.equal(tasks.displayResult(id,'B1').value,'');await tick();
}));
test('array followers remain worker results and stale file caches are not previewed as current',async()=>fixture({A1:'=SEQUENCE(2)',A2:'',B1:'=Unsupported()'},async({tasks,book,id,messages,reply,computed})=>{
 computed.add('A2');assert.equal(tasks.displayResult(id,'A2').pending,true);await tick();reply(messages.shift(),[['A2',2]]);await tick();assert.equal(tasks.displayResult(id,'A2').value,2);
 const cell=book.sheets[0].cells.B1;cell.cachedValue=42;cell.cachedFormula=cell.raw;assert.equal(tasks.displayResult(id,'B1').cachedValue,42);await tick();reply(messages.shift(),[['B1',42,{},null,{cached:true,stale:false,cachedValue:42}]]);await tick();
 book.referenceCachesValid=false;const dirty=new Map();dirty.metadataChanged=true;await tasks.sync(book,dirty,undefined,'doc',{persist:false});assert.equal(tasks.displayResult(id,'B1').value,'');await tick();reply(messages.shift(),[['B1','#NAME?',{},'#NAME?',{cached:true,stale:true,cachedValue:42}]]);await tick();assert.equal(tasks.displayResult(id,'B1').value,'#NAME?');assert.equal(tasks.displayResult(id,'B1').stale,true);
}));
