import {xml} from './chart-xml.mjs';
export function validatePreservedData(value,limit=1_000_000){
  const stack=[[value,0]];let size=0;
  while(stack.length){const [v,depth]=stack.pop();if(depth>20)throw Error('보존된 Excel 정보가 너무 복잡합니다.');
    if(v===null||typeof v==='boolean')size+=8;
    else if(typeof v==='number'){if(!Number.isFinite(v))throw Error('보존된 Excel 숫자가 올바르지 않습니다.');size+=16;}
    else if(typeof v==='string')size+=v.length*2;
    else if(typeof v==='object'){for(const [key,item] of Object.entries(v)){if(['__proto__','constructor','prototype'].includes(key))throw Error('보존된 Excel 정보의 키가 올바르지 않습니다.');size+=key.length*2+8;stack.push([item,depth+1]);}}
    else throw Error('보존된 Excel 정보가 올바르지 않습니다.');
    if(size>limit)throw Error('보존된 Excel 정보가 크기 제한을 초과했습니다.');
  }
}
export function validatePreservedXML(value,root){
  if(typeof value!=='string'||value.length>6_000_000)throw Error('보존된 Excel XML이 너무 큽니다.');
  const doc=xml(value);if(doc.documentElement.localName!==root)throw Error('보존된 Excel XML 종류가 올바르지 않습니다.');return doc;
}
