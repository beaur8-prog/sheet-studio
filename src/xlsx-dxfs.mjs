import {xml,stringify,children,child} from './chart-xml.mjs';
import {validatePreservedXML} from './xlsx-preservation-schema.mjs';
function canonical(node){
  return JSON.stringify([node.namespaceURI,node.localName,Array.from(node.attributes||[]).filter(a=>a.namespaceURI!=='http://www.w3.org/2000/xmlns/').map(a=>[a.namespaceURI,a.localName,a.value]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b))),Array.from(node.childNodes).flatMap(n=>n.nodeType===1?[canonical(n)]:n.nodeType===3&&n.nodeValue.trim()?[n.nodeValue]:[])]);
}
/** Deduplicate shared differential formats instead of appending another copy
 * on every XLSX roundtrip. Tables, filters and preserved CF use the same map.
 */
export async function mergePreservedDxfs(zip,sources=[]){
  if(!sources.length)return [];
  const doc=xml(await zip.file('xl/styles.xml').async('string')),root=doc.documentElement;let dxfs=child(root,'dxfs');
  if(!dxfs){dxfs=doc.createElementNS(root.namespaceURI,'dxfs');root.insertBefore(dxfs,child(root,'tableStyles')||child(root,'colors')||child(root,'extLst')||null);}
  const nodes=children(dxfs,'dxf'),existing=new Map(nodes.map((node,i)=>[canonical(node),i])),map=[];let changed=false;
  for(const source of sources){const node=validatePreservedXML(source,'dxf').documentElement,key=canonical(node);let id=existing.get(key);if(id===undefined){id=children(dxfs,'dxf').length;dxfs.appendChild(doc.importNode(node,true));existing.set(key,id);changed=true;}map.push(id);}
  dxfs.setAttribute('count',String(children(dxfs,'dxf').length));if(changed)zip.file('xl/styles.xml',stringify(doc));return map;
}
export function remapDxfIds(root,map){
  const nodes=[root,...Array.from(root.getElementsByTagNameNS('*','*'))];
  for(const node of nodes)for(const attribute of Array.from(node.attributes||[]))if(/dxfId$/i.test(attribute.localName)&&/^\d+$/.test(attribute.value)&&map[Number(attribute.value)]!==undefined)node.setAttribute(attribute.name,String(map[Number(attribute.value)]));
}
