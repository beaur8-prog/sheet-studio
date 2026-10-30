import JSZip from 'jszip';
import {xml,stringify,all,child,children,attr,readRelationships,resolvePart,relationships,relationship,relPath} from './chart-xml.mjs';
import {capturePartTree,writeStoredParts,contentTypeWriter,validateStoredParts,validateStoredRelationships} from './xlsx-part-store.mjs';
import {validatePreservedXML} from './xlsx-preservation-schema.mjs';
import {worksheetParts} from './xlsx-preservation.mjs';

const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XDA='http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const METADATA_TYPE='application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
const metadataRel='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';
const calculationStateURI='{D0BC62EC-B5BF-4DE5-AF5A-04AB17EB64CD}';
const calculationStateNS='urn:sheet-studio:calculation-state:1';
const truth=value=>value==='1'||value==='true';
const node=(doc,name)=>doc.createElementNS(NS,name);

export async function readArrayMetadata(zip){
  const rels=await readRelationships(zip,'xl/workbook.xml'),rel=rels.find(r=>r.type.endsWith('/sheetMetadata')&&r.mode!=='External'),cells=new Map();
  if(!rel)return {cells,collect:undefined};
  const path=resolvePart('xl/workbook.xml',rel.target),source=await zip.file(path)?.async('string');
  if(!source)throw Error('Excel 배열 메타데이터 파일을 찾을 수 없습니다.');
  const doc=validatePreservedXML(source,'metadata'),types=children(child(doc.documentElement,'metadataTypes'),'metadataType');
  const dynamicType=types.findIndex(n=>attr(n,'name')==='XLDAPR')+1,future=children(doc.documentElement,'futureMetadata').find(n=>attr(n,'name')==='XLDAPR');
  const dynamicFuture=new Set(children(future,'bk').flatMap((bk,index)=>all(bk,'dynamicArrayProperties').some(n=>truth(attr(n,'fDynamic')))?[index]:[]));
  const dynamicCells=new Set(children(child(doc.documentElement,'cellMetadata'),'bk').flatMap((bk,index)=>children(bk,'rc').some(rc=>Number(attr(rc,'t'))===dynamicType&&dynamicFuture.has(Number(attr(rc,'v'))))?[index+1]:[]));
  const related=rels.filter(r=>r!==rel&&/(?:\/rdRich|\/richValue)/i.test(r.type)).map(r=>({...r,target:r.mode==='External'?r.target:resolvePart('xl/workbook.xml',r.target)}));
  const parts=await capturePartTree(zip,[path,...related.filter(r=>r.mode!=='External').map(r=>r.target)]);delete parts[path];
  const preserved={xml:source,sourcePath:path,parts,rels:related};
  const collect=(name,cell)=>{
    const cm=cell.cm===undefined?undefined:Number(cell.cm),vm=cell.vm===undefined?undefined:Number(cell.vm);
    if([cm,vm].some(n=>n!==undefined&&(!Number.isInteger(n)||n<0||n>1000000)))throw Error('셀 메타데이터 번호가 올바르지 않습니다.');
    if(!cells.has(name))cells.set(name,new Map());
    cells.get(name).set(cell.r,{cm,vm,dynamic:dynamicCells.has(cm),ref:cell.formulaAttributes?.ref,formula:cell.formula});
  };
  // Retain stale-cache provenance when this app saves an unsupported dynamic
  // formula. A subsequent editor's recalculation owns its newly saved caches.
  const state=children(child(doc.documentElement,'extLst'),'ext').find(n=>attr(n,'uri')===calculationStateURI);
  const core=await zip.file('docProps/core.xml')?.async('string');
  const stale=!!state&&all(state,'calculationState').some(n=>n.namespaceURI===calculationStateNS&&truth(attr(n,'staleCaches')))&&!!core&&all(xml(core),'lastModifiedBy')[0]?.textContent==='Sheet Studio';
  return {cells,collect,preserved,stale};
}

export function validateArrayMetadata(value){
  if(!value||typeof value!=='object'||typeof value.sourcePath!=='string')throw Error('Excel 셀 메타데이터가 올바르지 않습니다.');
  validatePreservedXML(value.xml,'metadata');
  validateStoredParts({...value.parts,[value.sourcePath]:{data:'',contentType:METADATA_TYPE}});
  validateStoredRelationships(value.rels||[],value.parts||{});
}

function dynamicRecord(doc){
  const root=doc.documentElement;
  let types=child(root,'metadataTypes');if(!types){types=node(doc,'metadataTypes');root.insertBefore(types,root.firstChild);}
  let type=children(types,'metadataType').find(n=>attr(n,'name')==='XLDAPR');
  if(!type){type=node(doc,'metadataType');for(const [key,value] of Object.entries({name:'XLDAPR',minSupportedVersion:120000,copy:1,pasteAll:1,pasteValues:1,merge:1,splitFirst:1,rowColShift:1,clearFormats:1,clearComments:1,assign:1,coerce:1,cellMeta:1}))type.setAttribute(key,String(value));types.appendChild(type);}
  types.setAttribute('count',String(children(types,'metadataType').length));
  const typeId=children(types,'metadataType').indexOf(type)+1;
  let future=children(root,'futureMetadata').find(n=>attr(n,'name')==='XLDAPR');
  if(!future){future=node(doc,'futureMetadata');future.setAttribute('name','XLDAPR');root.insertBefore(future,child(root,'cellMetadata')||child(root,'valueMetadata')||child(root,'extLst')||null);}
  let bk=children(future,'bk').find(n=>all(n,'dynamicArrayProperties').some(p=>truth(attr(p,'fDynamic'))));
  if(!bk){bk=node(doc,'bk');const list=node(doc,'extLst'),ext=node(doc,'ext'),props=doc.createElementNS(XDA,'xda:dynamicArrayProperties');ext.setAttribute('uri','{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}');props.setAttribute('fDynamic','1');props.setAttribute('fCollapsed','0');ext.appendChild(props);list.appendChild(ext);bk.appendChild(list);future.appendChild(bk);}
  future.setAttribute('count',String(children(future,'bk').length));const valueId=children(future,'bk').indexOf(bk);
  let cellMetadata=child(root,'cellMetadata');if(!cellMetadata){cellMetadata=node(doc,'cellMetadata');root.insertBefore(cellMetadata,child(root,'valueMetadata')||child(root,'extLst')||null);}
  let record=children(cellMetadata,'bk').find(n=>children(n,'rc').length===1&&children(n,'rc').some(rc=>Number(attr(rc,'t'))===typeId&&Number(attr(rc,'v'))===valueId));
  if(!record){record=node(doc,'bk');const rc=node(doc,'rc');rc.setAttribute('t',String(typeId));rc.setAttribute('v',String(valueId));record.appendChild(rc);cellMetadata.appendChild(record);}
  cellMetadata.setAttribute('count',String(children(cellMetadata,'bk').length));return children(cellMetadata,'bk').indexOf(record)+1;
}

export async function addArrayMetadata(bytes,book,names,arrays){
  const hasArrays=[...arrays.values()].some(a=>a.length);
  if(!book.xlsxMetadata&&!hasArrays)return bytes;
  const zip=await JSZip.loadAsync(bytes),metadata=book.xlsxMetadata,doc=xml(metadata?.xml||'<metadata xmlns="'+NS+'" xmlns:xda="'+XDA+'"/>'),cm=hasArrays?dynamicRecord(doc):null;
  let extensions=child(doc.documentElement,'extLst');
  for(const old of children(extensions,'ext').filter(n=>attr(n,'uri')===calculationStateURI))extensions.removeChild(old);
  if(book.referenceCachesValid===false&&[...arrays.values()].some(list=>list.some(a=>a.cached))){
    if(!extensions){extensions=node(doc,'extLst');doc.documentElement.appendChild(extensions);}
    const ext=node(doc,'ext'),state=doc.createElementNS(calculationStateNS,'ss:calculationState');ext.setAttribute('uri',calculationStateURI);state.setAttribute('staleCaches','1');ext.appendChild(state);extensions.appendChild(ext);
  }
  const types=xml(await zip.file('[Content_Types].xml').async('string')),content=contentTypeWriter(types),target='xl/metadata.xml';
  const map=part=>part===metadata?.sourcePath?target:part===relPath(metadata?.sourcePath||target)?relPath(target):'xl/sheetStudioMetadata/'+part.replace(/^(?:xl\/sheetStudioMetadata\/)+/,'');
  if(metadata)await writeStoredParts(zip,metadata.parts,map,content);
  zip.file(target,stringify(doc));content(target,METADATA_TYPE);
  const relDoc=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string')),used=new Set(all(relDoc,'Relationship').map(n=>attr(n,'Id')));let id=0;
  for(const rel of [{type:metadataRel,target},...(metadata?.rels||[]).map(r=>({...r,target:r.mode==='External'?r.target:map(r.target)}))]){
    let current=all(relDoc,'Relationship').find(n=>attr(n,'Type')===rel.type);
    if(!current){let key;do{key='rIdSheetStudioMetadata'+ ++id;}while(used.has(key));used.add(key);current=xml(relationships(relationship(key,rel.type,rel.mode==='External'?rel.target:'/'+rel.target,rel.mode))).documentElement.firstChild;current=relDoc.importNode(current,true);relDoc.documentElement.appendChild(current);}
    else current.setAttribute('Target',rel.mode==='External'?rel.target:'/'+rel.target);
  }
  zip.file('xl/_rels/workbook.xml.rels',stringify(relDoc));
  for(const part of await worksheetParts(zip)){
    const sheet=book.sheets.find(s=>(names?.get(s.id)||s.name)===part.name);if(!sheet||!part.type.endsWith('/worksheet'))continue;
    const patches=new Map();
    for(const ref in sheet.cells){const cell=sheet.cells[ref],m=cell.xlsxMetadata;if(m?.raw===cell.raw)patches.set(ref,{...(m.cm!==undefined?{cm:m.cm}:{}),...(m.vm!==undefined?{vm:m.vm}:{})});}
    for(const array of arrays.get(sheet.id)||[]){const cell=sheet.cells[array.ref],m=cell?.xlsxMetadata,original=cell?.xlsxFormula?.dynamic&&m?.raw===cell.raw?m.cm:undefined;patches.set(array.ref,{...patches.get(array.ref),cm:original||cm});}
    if(!patches.size)continue;
    const text=await zip.file(part.path).async('string');
    zip.file(part.path,text.replace(/<((?:[\w.-]+:)?c)(?=\s|\/?>)[^>]*>/g,tag=>{
      const ref=/\br="([^"]+)"/.exec(tag)?.[1],patch=patches.get(ref);if(!patch)return tag;
      tag=tag.replace(/\s(?:cm|vm)="[^"]*"/g,'');return tag.replace(/\/?>(?=$)/,end=>Object.entries(patch).map(([k,v])=>' '+k+'="'+v+'"').join('')+end);
    }));
  }
  zip.file('[Content_Types].xml',stringify(types));return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
