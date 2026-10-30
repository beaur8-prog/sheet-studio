import {xml,stringify,children,child,all,attr,readRelationships,resolvePart,relPath,relationships,relationship} from './chart-xml.mjs';
import {capturePartTree,writeStoredParts,validateStoredParts,validateStoredRelationships} from './xlsx-part-store.mjs';
import {drawingRelationshipIds,remapDrawingRelationships} from './xlsx-drawings.mjs';
import {validatePreservedXML} from './xlsx-preservation-schema.mjs';
export const EXTRA_SHEET_TAGS=['sheetPr','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','ignoredErrors','legacyDrawingHF'];
export const DOCUMENT_PROPERTY_FIELDS=['creator','title','subject','description','keywords','category','company','manager'];
export async function captureExtraSheetFeatures(zip,path,root,feature){
  const nodes=[];for(const name of EXTRA_SHEET_TAGS){const node=child(root,name);if(node){feature[name]=stringify(node);nodes.push(node);}}
  const ids=new Set(nodes.flatMap(node=>[...drawingRelationshipIds(node)]));if(!ids.size)return;
  const rels=(await readRelationships(zip,path)).filter(r=>ids.has(r.id)).map(r=>({...r,target:r.mode==='External'?r.target:resolvePart(path,r.target)}));
  if(rels.length!==ids.size)throw Error('인쇄 설정의 연결 파일 정보를 찾을 수 없습니다.');
  feature.linkedParts={rels,parts:await capturePartTree(zip,rels.filter(r=>r.mode!=='External').map(r=>r.target))};
}
export function validateExtraSheetFeatures(feature){
  for(const name of EXTRA_SHEET_TAGS)if(feature[name]!==undefined)validatePreservedXML(feature[name],name);
  if(feature.linkedParts!==undefined){const value=feature.linkedParts;if(!value||typeof value!=='object')throw Error('시트 연결 파일이 올바르지 않습니다.');validateStoredParts(value.parts);validateStoredRelationships(value.rels,value.parts);}
}
export async function writeExtraSheetRelationships(zip,path,feature,index,content){
  if(!feature.linkedParts)return new Map();
  const {parts,rels}=feature.linkedParts,prefix='xl/sheetStudioFeatures/sheet'+(index+1)+'/',map=part=>prefix+part.replace(/^(?:xl\/sheetStudioFeatures\/sheet\d+\/)+/,'');
  await writeStoredParts(zip,parts,map,content);
  const original=zip.file(relPath(path)),doc=xml(original?await original.async('string'):relationships('')),used=new Set(all(doc,'Relationship').map(n=>attr(n,'Id'))),ids=new Map();let n=0;
  for(const rel of rels){let id;do{id='rIdPreservedSheet'+ ++n;}while(used.has(id));used.add(id);ids.set(rel.id,id);const node=xml(relationships(relationship(id,rel.type,rel.mode==='External'?rel.target:'/'+map(rel.target),rel.mode))).documentElement.firstChild;doc.documentElement.appendChild(doc.importNode(node,true));}
  zip.file(relPath(path),stringify(doc));return ids;
}
export function extraSheetNode(feature,name,ids){
  const node=validatePreservedXML(feature[name],name).documentElement;remapDrawingRelationships(node,ids);return node;
}
const columnName=index=>{let s='';for(let n=index+1;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const columnIndex=text=>{let n=0;for(const c of text)n=n*26+c.charCodeAt(0)-64;return n-1;};
export function transformExtraSheetFeatures(feature,reference){
  if(feature.ignoredErrors){const doc=xml(feature.ignoredErrors);for(const node of all(doc,'ignoredError')){const refs=attr(node,'sqref').split(/\s+/).filter(Boolean).map(reference).filter(r=>!r.includes('#REF!'));if(refs.length)node.setAttribute('sqref',refs.join(' '));else node.parentNode.removeChild(node);}if(all(doc,'ignoredError').length)feature.ignoredErrors=stringify(doc);else delete feature.ignoredErrors;}
  for(const [name,rows] of [['rowBreaks',true],['colBreaks',false]])if(feature[name]){
    const doc=xml(feature[name]);for(const node of all(doc,'brk')){const id=Number(attr(node,'id')),ref=rows?'A'+Math.max(1,id):columnName(Math.max(0,id-1))+'1',next=reference(ref),point=/^([A-Z]+)(\d+)$/.exec(next);if(!point)node.parentNode.removeChild(node);else node.setAttribute('id',String(rows?Number(point[2]):columnIndex(point[1])+1));}
    const breaks=all(doc,'brk');if(!breaks.length)delete feature[name];else {doc.documentElement.setAttribute('count',String(breaks.length));doc.documentElement.setAttribute('manualBreakCount',String(breaks.filter(n=>['1','true'].includes(attr(n,'man'))).length));feature[name]=stringify(doc);}
  }
}
function sourcePageSetup(sheet){if(sheet?.xlsxFeatures?.pageSetup)return sheet.xlsxFeatures.pageSetup;if(sheet?.kind==='chart'&&sheet.xlsxChartSheet?.xml){const node=child(xml(sheet.xlsxChartSheet.xml).documentElement,'pageSetup');return node?stringify(node):'';}return '';}
export function printOrientation(sheet){const source=sourcePageSetup(sheet);return source&&attr(xml(source).documentElement,'orientation')==='landscape'?'landscape':'portrait';}
export function setPrintOrientation(sheet,orientation){
  if(!['portrait','landscape'].includes(orientation))throw Error('인쇄 방향이 올바르지 않습니다.');const source=sourcePageSetup(sheet);sheet.xlsxFeatures||={};const doc=xml(source||'<pageSetup xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');doc.documentElement.setAttribute('orientation',orientation);sheet.xlsxFeatures.pageSetup=stringify(doc);
}
export function documentProperties(book){
  const result={};for(const name of DOCUMENT_PROPERTY_FIELDS)if(typeof book[name]==='string'&&book[name])result[name]=book[name];
  if(book.created instanceof Date&&Number.isFinite(book.created.getTime()))result.created=book.created.toISOString();return result;
}
export function validateDocumentProperties(value){
  if(!value||typeof value!=='object'||Array.isArray(value))throw Error('통합 문서 속성이 올바르지 않습니다.');
  for(const [name,text] of Object.entries(value))if(![...DOCUMENT_PROPERTY_FIELDS,'created'].includes(name)||typeof text!=='string'||text.length>32768||name==='created'&&!Number.isFinite(new Date(text).getTime()))throw Error('통합 문서 속성이 올바르지 않습니다.');
}
