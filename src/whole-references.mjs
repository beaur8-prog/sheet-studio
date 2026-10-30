export function wholeEndpoint(value,type='word'){
  const text=String(value),column=type==='word'&&/^(\$?)([A-Za-z]{1,3})$/.exec(text),row=/^(\$?)([1-9]\d*)$/.exec(text);
  if(column){let index=0;for(const letter of column[2].toUpperCase())index=index*26+letter.charCodeAt(0)-64;return {axis:'c',index:index-1,locked:!!column[1],limit:16384};}
  if(row)return {axis:'r',index:Number(row[2])-1,locked:!!row[1],limit:1048576};return null;
}
export function wholeNodeEndpoint(node){if(node.type==='name')return {...wholeEndpoint(node.name),sheet:node.sheet};if(['literal','axisEndpoint'].includes(node.type)&&typeof node.value==='number'&&Number.isInteger(node.value))return {...wholeEndpoint(node.value,'number'),sheet:node.sheet};return null;}
export function wholeReferences(tokens){
  const read=(start,inherited)=>{let at=start,explicitSheet=null;if(['word','sheet'].includes(tokens[at]?.type)&&tokens[at+1]?.value==='!'){explicitSheet=tokens[at].value;at+=2;}const token=tokens[at];if(!['word','number'].includes(token?.type)||['(','!'].includes(tokens[at+1]?.value))return null;const point=wholeEndpoint(token.value,token.type);return point?{...point,token,start:tokens[start].start,next:at+1,sheet:explicitSheet||inherited,explicitSheet}:null;},ranges=[];
  for(let i=0;i<tokens.length;i++){const first=read(i);if(!first||tokens[first.next]?.value!==':')continue;const last=read(first.next+1,first.sheet);if(last&&last.axis===first.axis){ranges.push({first,last,axis:first.axis,limit:first.limit});i=last.next-1;}}
  return ranges;
}
