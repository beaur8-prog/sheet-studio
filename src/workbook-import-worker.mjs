import {encodeWorkbookWire} from './workbook-wire.mjs';
import {parseWorkbook,validateWorkbook} from './engine.mjs';
import {readSpreadsheet} from './io.mjs';
import {describeWorkbook} from './workbook-description.mjs';
self.onmessage=async({data:{token,bytes,name,port}})=>{
  const progress=stage=>self.postMessage({token,progress:stage});
  try{
    if(bytes.byteLength>100*1024*1024)throw Error('100 MB 이하의 파일을 선택하세요.');
    progress({phase:'inspect'});
    const loaded=/\.json$/i.test(name)?parseWorkbook(new TextDecoder().decode(bytes)):validateWorkbook(await readSpreadsheet(bytes,name,{onProgress:progress}));
    progress({phase:'validate'});const description=describeWorkbook(loaded);
    const staged=new Promise((resolve,reject)=>{port.onmessage=({data})=>{if(data.token!==token)return;if(data.error)reject(Error(data.error));else if(data.staged)resolve();};});
    progress({phase:'stage'});const wire=encodeWorkbookWire(loaded);port.postMessage({token,encoding:'sheet-json-v1',bytes:wire,description},[wire.buffer]);await staged;port.close();self.postMessage({token,staged:true});
  }catch(error){port?.close();self.postMessage({token,error:error.message||String(error)});}
};
