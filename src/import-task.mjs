export function importCanceled(){const error=Error('파일 열기를 취소했습니다.');error.name='AbortError';return error;}
export function abortable(promise,signal){
  if(!signal)return Promise.resolve(promise);
  if(signal.aborted){void Promise.resolve(promise).catch(()=>{});return Promise.reject(importCanceled());}
  return new Promise((resolve,reject)=>{const cancel=()=>{cleanup();reject(importCanceled());},cleanup=()=>signal.removeEventListener('abort',cancel);signal.addEventListener('abort',cancel,{once:true});Promise.resolve(promise).then(value=>{cleanup();resolve(value);},error=>{cleanup();reject(error);});});
}
// The parser owns no recovery store or live document. Transfer its candidate to
// the document worker over a port; terminate the parser before returning the
// document worker's preview to the renderer. A canceled candidate is discarded.
export async function importTask(bytes,name,{signal,onProgress=()=>{},onReady=()=>{},createWorker,prepare,read,discard}){
  if(signal?.aborted)throw importCanceled();
  const token=crypto.randomUUID(),channel=new MessageChannel();let worker,success=false;
  try{
    await abortable(prepare(token,channel.port1),signal);
    if(signal?.aborted)throw importCanceled();
    worker=createWorker();
    await abortable(new Promise((resolve,reject)=>{
      worker.onmessage=({data})=>{if(data.token!==token)return;if(data.progress){onProgress(data.progress);return;}if(data.error)reject(Error(data.error));else if(data.staged)resolve();};
      worker.onerror=event=>{event?.preventDefault?.();reject(Error('파일을 읽는 작업 스레드를 사용할 수 없습니다. 다시 열어 주세요.'));};
      worker.postMessage({token,bytes,name,port:channel.port2},[bytes.buffer,channel.port2]);
    }),signal);
    if(signal?.aborted)throw importCanceled();worker.terminate();worker=null;
    onReady();const result=await read(token);success=true;return result;
  }finally{
    worker?.terminate();channel.port1.close();channel.port2.close();
    if(!success)void Promise.resolve(discard(token)).catch(()=>{});
  }
}
