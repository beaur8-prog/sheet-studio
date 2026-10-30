import JSZip from 'jszip';
import {xml,stringify,child,children,all,attr,R,readRelationships,resolvePart,relationship,relationships} from './chart-xml.mjs';
import {capturePartTree,writeStoredParts,contentTypeWriter,validateStoredParts,validateStoredRelationships} from './xlsx-part-store.mjs';
import {validatePreservedXML} from './xlsx-preservation-schema.mjs';

export async function readExternalLinks(zip){
  const workbook=xml(await zip.file('xl/workbook.xml').async('string')),root=child(workbook.documentElement,'externalReferences');if(!root)return undefined;
  const workbookRels=await readRelationships(zip,'xl/workbook.xml'),rels=[];
  for(const ref of children(root,'externalReference')){
    const id=ref.getAttributeNS(R,'id')||attr(ref,'r:id'),rel=workbookRels.find(r=>r.id===id&&r.type.endsWith('/externalLink')&&r.mode!=='External');if(!rel)throw Error('Excel 외부 통합 문서 연결을 찾을 수 없습니다.');
    rels.push({...rel,target:resolvePart('xl/workbook.xml',rel.target)});
  }
  if(!rels.length)return undefined;
  return {xml:stringify(root),rels,parts:await capturePartTree(zip,rels.map(r=>r.target))};
}
export function validateExternalLinks(value){
  if(!value||typeof value!=='object')throw Error('외부 통합 문서 연결 정보가 올바르지 않습니다.');
  const root=validatePreservedXML(value.xml,'externalReferences').documentElement;validateStoredParts(value.parts);validateStoredRelationships(value.rels,value.parts);
  if(children(root,'externalReference').length!==value.rels.length||value.rels.some(r=>!r.type.endsWith('/externalLink')||r.mode==='External'))throw Error('외부 통합 문서 연결 목록이 올바르지 않습니다.');
  for(const ref of children(root,'externalReference'))if(!value.rels.some(r=>r.id===(ref.getAttributeNS(R,'id')||attr(ref,'r:id'))))throw Error('외부 참조의 연결 번호가 올바르지 않습니다.');
}
/** Preserve external targets as inert OOXML relationships. This never opens,
 * downloads or refreshes a referenced workbook, DDE server or OLE resource. */
export async function addExternalLinks(bytes,book){
  const source=book.xlsxExternalLinks;if(!source)return bytes;
  validateExternalLinks(source);const zip=await JSZip.loadAsync(bytes),workbook=xml(await zip.file('xl/workbook.xml').async('string')),rels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string')),types=xml(await zip.file('[Content_Types].xml').async('string')),content=contentTypeWriter(types);
  const map=path=>'xl/sheetStudioExternalLinks/'+path.replace(/^(?:xl\/sheetStudioExternalLinks\/)+/,'');await writeStoredParts(zip,source.parts,map,content);
  const root=workbook.importNode(xml(source.xml).documentElement,true),used=new Set(all(rels,'Relationship').map(n=>attr(n,'Id'))),ids=new Map();let sequence=0;
  for(const rel of source.rels){let id;do{id='rIdSheetStudioExternal'+ ++sequence;}while(used.has(id));used.add(id);ids.set(rel.id,id);const node=xml(relationships(relationship(id,rel.type,'/'+map(rel.target)))).documentElement.firstChild;rels.documentElement.appendChild(rels.importNode(node,true));}
  for(const ref of children(root,'externalReference'))ref.setAttributeNS(R,'r:id',ids.get(ref.getAttributeNS(R,'id')||attr(ref,'r:id')));
  const old=child(workbook.documentElement,'externalReferences');if(old)workbook.documentElement.removeChild(old);workbook.documentElement.insertBefore(root,children(workbook.documentElement).find(n=>['definedNames','calcPr','oleSize','customWorkbookViews','pivotCaches','smartTagPr','smartTagTypes','webPublishing','fileRecoveryPr','webPublishObjects','extLst'].includes(n.localName))||null);
  zip.file('xl/workbook.xml',stringify(workbook));zip.file('xl/_rels/workbook.xml.rels',stringify(rels));zip.file('[Content_Types].xml',stringify(types));return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
