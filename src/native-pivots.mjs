import {xml,stringify,all,child,children,attr,readRelationships,resolvePart,relPath,relationship,relationships} from './chart-xml.mjs';
import {worksheetParts} from './xlsx-preservation.mjs';
import {capturePartTree,writeStoredParts,contentTypeWriter,validateStoredParts} from './xlsx-part-store.mjs';
import {tableBounds} from './table-schema.mjs';
import {aggregations} from './pivot-schema.mjs';

const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const ignored=new Set(['xmlns','cacheId','createdVersion','updatedVersion','refreshedVersion','minRefreshableVersion']);
const canonical=node=>node?.nodeType===1?[node.localName,Array.from(node.attributes).filter(a=>!a.name.startsWith('xmlns')&&!ignored.has(a.localName)).map(a=>[a.localName,a.value]).sort((a,b)=>a[0].localeCompare(b[0])),Array.from(node.childNodes).filter(n=>n.nodeType===1||n.nodeType===3&&n.data.trim()).map(n=>n.nodeType===1?canonical(n):n.data.trim())]:null;
const same=(a,b)=>JSON.stringify(a)===JSON.stringify(b);

export async function nativePivotIndex(zip,book){
  const entries=[];
  for(const part of await worksheetParts(zip)){
    const sheet=book.sheets.find(s=>s.name===part.name);if(!sheet||!part.type.endsWith('/worksheet'))continue;
    for(const rel of await readRelationships(zip,part.path))if(rel.type.endsWith('/pivotTable')&&rel.mode!=='External'){
      const path=resolvePart(part.path,rel.target),file=zip.file(path);if(!file)throw Error('피벗 정의 파일을 찾을 수 없습니다.');
      const table=xml(await file.async('string')).documentElement;
      const cacheRel=(await readRelationships(zip,path)).find(r=>r.type.endsWith('/pivotCacheDefinition')&&r.mode!=='External');
      if(!cacheRel)throw Error('피벗 캐시 연결을 찾을 수 없습니다.');
      const cache=resolvePart(path,cacheRel.target),cacheFile=zip.file(cache);if(!cacheFile)throw Error('피벗 캐시 정의를 찾을 수 없습니다.');
      entries.push({sheetId:sheet.id,path,cache,cacheId:attr(table,'cacheId'),table,definition:xml(await cacheFile.async('string')).documentElement});
      if(entries.length>200)throw Error('피벗테이블은 통합 문서당 최대 200개까지 보존할 수 있습니다.');
    }
  }
  return entries;
}

export function nativePivotLayout(entry){
  const fields=children(child(entry.definition,'cacheFields'),'cacheField');
  const items=children(child(entry.table,'pivotFields'),'pivotField').map((field,i)=>children(child(field,'items'),'item').map(item=>{
    const shared=children(child(fields[i],'sharedItems'))[Number(attr(item,'x','-1'))];
    return [attr(item,'h','0'),attr(item,'t'),shared?canonical(shared):null];
  }));
  return JSON.stringify({table:canonical(entry.table),source:canonical(child(entry.definition,'cacheSource')),fields:fields.map(field=>[attr(field,'name'),children(field).filter(n=>n.localName!=='sharedItems').map(canonical)]),items});
}

// Older app files have no layout proof. Compare the supported definition and
// reject differences before restoring Custom XML. This is intentionally strict.
export function legacyPivotMatches(entry,stored){
  if(!entry)return false;
  const {table,definition}=entry,spec=stored.spec;
  const fields=children(child(definition,'cacheFields'),'cacheField'),names=fields.map(f=>attr(f,'name'));
  const source=child(child(definition,'cacheSource'),'worksheetSource');
  if(attr(source,'sheet')!==stored.source||attr(source,'ref').replace(/\$/g,'')!==spec.sourceRange||attr(source,'name'))return false;
  const axis=name=>children(child(table,name),'field').map(f=>Number(attr(f,'x')));
  if(!same(axis('rowFields').map(i=>names[i]),spec.rows)||!same(axis('colFields').filter(i=>i!==-2).map(i=>names[i]),spec.columns))return false;
  if(axis('colFields').includes(-2)!==(spec.values.length>1))return false;
  const data=children(child(table,'dataFields'),'dataField').map(f=>({field:names[Number(attr(f,'fld'))],aggregate:attr(f,'subtotal','sum')==='countNums'?'countNumbers':attr(f,'subtotal','sum')}));
  if(!same(data,spec.values)||['rowGrandTotals','colGrandTotals'].some(k=>!['0','false'].includes(attr(table,k,'1'))!==spec.totals))return false;
  if(attr(table,'dataCaption','값')!=='값'||children(child(table,'dataFields'),'dataField').some((field,i)=>attr(field,'name')!==`${aggregations[spec.values[i].aggregate]}: ${spec.values[i].field} (${i+1})`||attr(field,'numFmtId','0')!=='0'))return false;
  const style=child(table,'pivotTableStyleInfo');if(style&&(attr(style,'name')!=='PivotStyleMedium9'||attr(style,'showRowStripes','1')!=='1'||attr(style,'showColStripes','0')!=='0'))return false;
  if(['compact','compactData','outline','outlineData'].some(k=>!['0','false'].includes(attr(table,k,'0'))))return false;
  if(children(table).some(n=>['formats','conditionalFormats','chartFormats','pivotFilters','rowHierarchiesUsage','colHierarchiesUsage','extLst'].includes(n.localName)))return false;
  const expectedPages=spec.filters.filter(f=>!spec.rows.includes(f.field)&&!spec.columns.includes(f.field)).map(f=>f.field);
  if(!same(children(child(table,'pageFields'),'pageField').map(f=>names[Number(attr(f,'fld'))]),expectedPages))return false;
  if(children(child(table,'pageFields'),'pageField').some(f=>attr(f,'item')!==''))return false;
  const typed=n=>n.localName==='m'?'':n.localName==='n'?Number(attr(n,'v')):n.localName==='b'?attr(n,'v')==='1':n.localName==='s'?attr(n,'v'):undefined;
  const pivots=children(child(table,'pivotFields'),'pivotField');
  for(let i=0;i<fields.length;i++){
    const field=pivots[i],filter=spec.filters.find(f=>f.field===names[i]),shared=children(child(fields[i],'sharedItems'));
    if(!field)return false;
    const order={asc:'ascending',desc:'descending',source:'manual'}[spec.sort];
    if((spec.rows.includes(names[i])||spec.columns.includes(names[i]))&&attr(field,'sortType','manual')!==order)return false;
    const items=children(child(field,'items'),'item');
    for(const item of items){const index=Number(attr(item,'x','-1'));if(index<0||index>=shared.length)return false;const value=typed(shared[index]);if(value===undefined)return false;const hidden=['1','true'].includes(attr(item,'h'));if(hidden!==(filter?!filter.values.some(v=>same([typeof v,v],[typeof value,value])):false))return false;}
  }
  return true;
}

export async function preserveNativePivots(zip,book,entries){
  if(!entries.length)return;
  const roots=[...new Set(entries.map(e=>e.path))],parts=await capturePartTree(zip,roots,{maxBytes:24_000_000,maxFiles:1000});
  const styles=zip.file('xl/styles.xml')?xml(await zip.file('xl/styles.xml').async('string')):null;
  const formats=Object.fromEntries(all(styles,'numFmt').map(n=>[attr(n,'numFmtId'),attr(n,'formatCode')]));
  book.xlsxPivotArchive={version:1,parts,formats,sheets:entries.map(e=>({sheetId:e.sheetId,path:e.path,cache:e.cache,cacheId:e.cacheId,...(attr(child(e.table,'location'),'ref')?{range:attr(child(e.table,'location'),'ref').replace(/\$/g,'')}:{} )})),names:Object.fromEntries(book.sheets.map(s=>[s.id,s.name]))};
  if(entries.some(e=>attr(child(e.definition,'cacheSource'),'type')!=='worksheet'||all(e.table,'formats').length||all(e.table,'conditionalFormats').length||all(e.table,'extLst').length||all(e.definition,'extLst').length)||Object.keys(zip.files).some(p=>/^xl\/(?:slicerCaches|slicers|timelineCaches|timelines|model)\//.test(p)))book.xlsxPivotArchive.blockReason='외부 데이터·슬라이서·고급 서식이 연결된 원본 피벗의 XLSX 보존 저장은 아직 지원하지 않습니다. Excel에서 저장하거나 일반 범위로 변환한 사본을 저장하세요. 앱 문서는 원본 정보를 유지합니다.';
}

export function validateNativePivotArchive(value){
  if(!value||value.version!==1||!Array.isArray(value.sheets)||!value.sheets.length||value.sheets.length>200||!value.names||typeof value.names!=='object'||Array.isArray(value.names)||!value.formats||typeof value.formats!=='object'||Array.isArray(value.formats))throw Error('보존된 피벗 정보가 올바르지 않습니다.');
  validateStoredParts(value.parts);
  if(value.blockReason!==undefined&&(typeof value.blockReason!=='string'||value.blockReason.length>1000))throw Error('피벗 저장 제한 안내가 올바르지 않습니다.');
  if(Object.entries(value.names).length>100||Object.entries(value.names).some(([id,name])=>!id||typeof name!=='string'||name.length>31))throw Error('피벗 원본 시트 정보가 올바르지 않습니다.');
  if(Object.keys(value.formats).length>65536||Object.entries(value.formats).some(([id,code])=>!/^\d+$/.test(id)||typeof code!=='string'||code.length>8192))throw Error('피벗 숫자 형식 정보가 올바르지 않습니다.');
  for(const item of value.sheets)if(!item||!Object.hasOwn(value.names,item.sheetId)||!Object.hasOwn(value.parts,item.path)||!Object.hasOwn(value.parts,item.cache)||typeof item.cacheId!=='string'||!/^\d+$/.test(item.cacheId))throw Error('피벗 정의의 연결이 올바르지 않습니다.');
  for(const item of value.sheets)if(item.range!==undefined&&(typeof item.range!=='string'||item.range.length>100||!tableBounds(item.range)))throw Error('피벗 결과 범위가 올바르지 않습니다.');
}

export const preservedPivots=(book,id)=>book.xlsxPivotArchive?.sheets.filter(s=>s.sheetId===id)||[];
export function preservedPivotTouches(book,id,selection){
  const b=typeof selection==='string'?tableBounds(selection):selection;
  return preservedPivots(book,id).some(item=>{const a=item.range&&tableBounds(item.range);return !a||!b||a.r1<=b.r2&&a.r2>=b.r1&&a.c1<=b.c2&&a.c2>=b.c1;});
}
export function pivotHasExtraCells(target,entry,actual){
  const b=actual&&tableBounds(attr(child(actual.table,'location'),'ref'));if(!b)return false;
  if(target.merges?.length||target.tables?.length)return true;
  const pages=entry.rowOffset?entry.rowOffset-1:0;
  const filters=entry.spec.filters.filter(f=>!entry.spec.rows.includes(f.field)&&!entry.spec.columns.includes(f.field));
  for(const ref in target.cells){
    const p=tableBounds(ref),cell=target.cells[ref];if(p.r1>=b.r1&&p.r1<=b.r2&&p.c1>=b.c1&&p.c1<=b.c2)continue;
    if(pages&&p.r1<pages&&p.c1<=1){const f=filters[p.r1],expected=p.c1===0?f?.field:f?.values.length===1?String(f.values[0]):'(여러 항목)';if(String(cell.raw||'').replace(/^'/,'')===expected&&!cell.note)continue;}
    return true;
  }
  return false;
}
export function nativePivotNames(book){return new Set((book.xlsxPivotArchive?.sheets||[]).map(item=>attr(xml(new TextDecoder().decode(Uint8Array.from(atob(book.xlsxPivotArchive.parts[item.path].data),c=>c.charCodeAt(0)))).documentElement,'name')));}
export function removePreservedPivots(book,id){
  if(!book.xlsxPivotArchive)return;
  book.xlsxPivotArchive.sheets=book.xlsxPivotArchive.sheets.filter(s=>s.sheetId!==id);
  if(!book.xlsxPivotArchive.sheets.length)delete book.xlsxPivotArchive;
  const sheet=book.sheets.find(s=>s.id===id);if(sheet)delete sheet.pivotImportNotice;
}
export function assertPreservedPivotStructure(book){
  if(book.xlsxPivotArchive)throw Error('원본 피벗을 보존하는 문서의 행·열·시트 구성 변경은 Excel에서 진행하세요. 일반 셀 값과 서식은 편집할 수 있습니다.');
}

export async function addNativePivotParts(bytes,book,names){
  const archive=book.xlsxPivotArchive;if(!archive)return bytes;validateNativePivotArchive(archive);
  if(archive.blockReason)throw Error(archive.blockReason);
  for(const [id,name] of Object.entries(archive.names))if(!book.sheets.some(s=>s.id===id&&s.name===name))throw Error('원본 피벗의 시트 연결을 보존할 수 없습니다. 시트 구성은 Excel에서 변경하세요.');
  const {default:JSZip}=await import('jszip'),zip=await JSZip.loadAsync(bytes),types=xml(await zip.file('[Content_Types].xml').async('string')),content=contentTypeWriter(types);
  const wb=xml(await zip.file('xl/workbook.xml').async('string')),wbRels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
  let caches=child(wb.documentElement,'pivotCaches');if(!caches){caches=wb.createElementNS(NS,'pivotCaches');wb.documentElement.insertBefore(caches,children(wb.documentElement).find(n=>['smartTagPr','smartTagTypes','webPublishing','fileRecoveryPr','webPublishObjects','extLst'].includes(n.localName))||null);}
  let nextId=Math.max(0,...children(caches,'pivotCache').map(n=>Number(attr(n,'cacheId'))))+1;
  const prefix='xl/sheetStudioPivotArchive/',map=path=>prefix+path.replace(/^(?:xl\/sheetStudioPivotArchive\/)+/,'');
  await writeStoredParts(zip,archive.parts,map,content);
  const styles=xml(await zip.file('xl/styles.xml').async('string'));let numFmts=child(styles.documentElement,'numFmts');
  const formatId=old=>{
    const code=archive.formats[old];if(!code)return old;
    if(!numFmts){numFmts=styles.createElementNS(NS,'numFmts');styles.documentElement.insertBefore(numFmts,styles.documentElement.firstChild);}
    let found=children(numFmts,'numFmt').find(n=>attr(n,'formatCode')===code);if(!found){found=styles.createElementNS(NS,'numFmt');found.setAttribute('numFmtId',String(Math.max(163,...children(numFmts,'numFmt').map(n=>Number(attr(n,'numFmtId'))))+1));found.setAttribute('formatCode',code);numFmts.appendChild(found);numFmts.setAttribute('count',String(children(numFmts,'numFmt').length));}return attr(found,'numFmtId');
  };
  const cacheIds=new Map(),usedNames=new Set();
  for(const file of zip.file(/^xl\/pivotTables\/[^/]+\.xml$/))usedNames.add(attr(xml(await file.async('string')).documentElement,'name'));
  const parts=await worksheetParts(zip);let index=0;
  for(const item of archive.sheets){
    const part=parts.find(p=>p.name===(names?.get(item.sheetId)||book.sheets.find(s=>s.id===item.sheetId)?.name));if(!part)throw Error('피벗 결과 시트를 찾을 수 없습니다.');
    if(!cacheIds.has(item.cache)){
      const id=String(nextId++),rid='rIdPreservedPivotCache'+id;cacheIds.set(item.cache,id);
      const node=wb.createElementNS(NS,'pivotCache');node.setAttribute('cacheId',id);node.setAttributeNS(R.slice(0,-1),'r:id',rid);caches.appendChild(node);
      wbRels.documentElement.appendChild(wbRels.importNode(xml(relationships(relationship(rid,R+'pivotCacheDefinition','/'+map(item.cache)))).documentElement.firstChild,true));
    }
    const table=xml(await zip.file(map(item.path)).async('string'));table.documentElement.setAttribute('cacheId',cacheIds.get(item.cache));
    let name=attr(table.documentElement,'name')||'PreservedPivot';while(usedNames.has(name))name='PreservedPivot'+ ++index;usedNames.add(name);table.documentElement.setAttribute('name',name);
    for(const node of [table.documentElement,...all(table,'dataField'),...all(table,'pivotField')])if(node.hasAttribute('numFmtId'))node.setAttribute('numFmtId',formatId(attr(node,'numFmtId')));
    zip.file(map(item.path),stringify(table));
    const rp=relPath(part.path),rels=xml(zip.file(rp)?await zip.file(rp).async('string'):relationships('')),ids=new Set(all(rels,'Relationship').map(n=>attr(n,'Id')));let rid='rIdPreservedPivot'+ ++index;while(ids.has(rid))rid+='1';
    rels.documentElement.appendChild(rels.importNode(xml(relationships(relationship(rid,R+'pivotTable','/'+map(item.path)))).documentElement.firstChild,true));zip.file(rp,stringify(rels));
  }
  for(const path of cacheIds.keys()){
    const cache=xml(await zip.file(map(path)).async('string'));for(const node of all(cache,'cacheField'))if(node.hasAttribute('numFmtId'))node.setAttribute('numFmtId',formatId(attr(node,'numFmtId')));zip.file(map(path),stringify(cache));
  }
  zip.file('xl/styles.xml',stringify(styles));zip.file('xl/workbook.xml',stringify(wb));zip.file('xl/_rels/workbook.xml.rels',stringify(wbRels));zip.file('[Content_Types].xml',stringify(types));
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
