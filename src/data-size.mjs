// Bound estimates without ever creating a document-sized JSON string. Shared
// objects count once: this is a history memory budget, not a file-size promise.
export function estimateDataBytes(value, limit = Infinity, seen = new Set()) {
  let bytes=0;const stack=[value];
  const stringBytes=value=>{
    if(!/["\\\u0000-\u001f]/.test(value))return (value.length+2)*2;
    if(value.length>1_000_000)return value.length*12+4;
    return JSON.stringify(value).length*2;
  };
  while(stack.length){
    const current=stack.pop();
    if(current&&typeof current==='object'){
      if(seen.has(current))continue;seen.add(current);bytes+=4;
      for(const key of Object.keys(current)){
        if(!Array.isArray(current))bytes+=stringBytes(key)+2;
        bytes+=2;stack.push(current[key]);
        if(bytes>limit)return limit+1;
      }
    }else if(typeof current==='string')bytes+=stringBytes(current);
    else if(current!==undefined)bytes+=(JSON.stringify(current)?.length||0)*2;
    if(bytes>limit)return limit+1;
  }
  return bytes;
}
