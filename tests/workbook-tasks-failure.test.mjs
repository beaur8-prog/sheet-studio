import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbookTasks} from '../src/workbook-tasks.mjs';
test('a document-worker failure interrupts a parser waiting to stage and notifies the file-read controller',async()=>{
  const OriginalWorker=globalThis.Worker,workers=[],errors=[],fatal=[];
  globalThis.Worker=class{
    constructor(url){this.url=String(url);this.stopped=false;workers.push(this);}
    postMessage(message){if(message.type==='import-port')queueMicrotask(()=>this.onmessage({data:{id:message.id,value:true}}));}
    terminate(){this.stopped=true;}
  };
  let tasks;try{
    tasks=createWorkbookTasks({onSave(){},onCalculated(){},onError:error=>errors.push(error),onFatal:error=>fatal.push(error)});
    const job=tasks.import(new Uint8Array([1]),'test.xlsx');await new Promise(resolve=>setTimeout(resolve,1));assert.equal(workers.length,2);
    const start=performance.now();workers[0].onerror();await assert.rejects(job,/작업 스레드.*다시 실행/);assert.ok(performance.now()-start<500);assert.equal(workers[1].stopped,true);assert.equal(fatal.length,1);assert.equal(errors.length,1);assert.equal(fatal[0],errors[0]);
    await assert.rejects(tasks.import(new Uint8Array([1]),'test.xlsx'),/작업 스레드/);assert.equal(workers.length,2,'failed primary cannot start another orphan parser');
  }finally{tasks?.destroy();if(OriginalWorker===undefined)delete globalThis.Worker;else globalThis.Worker=OriginalWorker;}
});
