export function androidHost(win = window) {
  const host=win.sheetStudioAndroid;
  return host?.protocolVersion===1&&host.platform==='android'?host:null;
}
export async function androidFile(result) {
  if(!result)return null;
  if(result instanceof File)return result;
  const bytes=result.data instanceof Blob?result.data:result.data instanceof ArrayBuffer||ArrayBuffer.isView(result.data)?result.data:null;
  if(!bytes)throw Error('파일 내용을 받지 못했습니다. 파일 선택기에서 다시 열어 주세요.');
  return new File([bytes],result.name||'통합 문서.xlsx',{type:result.mime||'application/octet-stream'});
}

// Source snapshots are separate from both the editable workbook and recovery.
// A File/Blob is immutable; sharing it never re-reads a possibly changed URI.
export async function openAndroidSources() {
  const db=await new Promise((resolve,reject)=>{
    const req=indexedDB.open('sheet-studio-android-sources',1);
    req.onupgradeneeded=()=>req.result.createObjectStore('sources');
    req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);
  });
  db.onversionchange=()=>db.close();
  function access(mode,operation){return new Promise((resolve,reject)=>{
    const tx=db.transaction('sources',mode),request=operation(tx.objectStore('sources'));let result;
    request.onsuccess=()=>{result=request.result;};
    tx.oncomplete=()=>resolve(result);tx.onerror=()=>{};tx.onabort=()=>reject(tx.error||Error('원본 스냅샷을 보관하지 못했습니다.'));
  });}
  return {get:id=>access('readonly',store=>store.get(id)),put:(id,file)=>access('readwrite',store=>store.put(file,id)),remove:id=>access('readwrite',store=>store.delete(id)),close:()=>db.close()};
}
