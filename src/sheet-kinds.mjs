import {validateStoredParts,validateStoredRelationships,capturePartTree,writeStoredParts,contentTypeWriter} from './xlsx-part-store.mjs';
import {validatePreservedXML} from './xlsx-preservation-schema.mjs';
import {worksheetParts} from './xlsx-preservation.mjs';
import {xml,stringify,children,child,all,attr,R,relPath,readRelationships,resolvePart,relationship,relationships} from './chart-xml.mjs';
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main',TYPE='application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml';
export function validateSheetKind(sheet){
  if(sheet.kind!==undefined&&!['worksheet','chart'].includes(sheet.kind))throw Error('시트 종류가 올바르지 않습니다.');
  if(sheet.state!==undefined&&!['visible','hidden','veryHidden'].includes(sheet.state))throw Error('시트 표시 상태가 올바르지 않습니다.');
  if(sheet.kind==='chart'&&(Object.keys(sheet.cells).length||sheet.tables?.length||sheet.merges?.length||sheet.pivot||sheet.autoFilter))throw Error('차트 전용 시트에는 셀이나 표를 넣을 수 없습니다.');
  const source=sheet.xlsxChartSheet;if(source!==undefined){if(!source||sheet.kind!=='chart')throw Error('차트 시트의 원본 정보가 올바르지 않습니다.');validatePreservedXML(source.xml,'chartsheet');if(source.parts)validateStoredParts(source.parts);if(source.rels)validateStoredRelationships(source.rels,source.parts);}
}
export async function restoreSheetKinds(zip,book,newSheet){
  const parts=await worksheetParts(zip),doc=xml(await zip.file('xl/workbook.xml').async('string')),entries=all(doc,'sheet'),ordered=[];
  for(const part of parts){
    let sheet=book.sheets.find(s=>s.name===part.name);const node=entries.find(n=>attr(n,'name')===part.name),state=attr(node,'state','visible');
    if(part.type.endsWith('/chartsheet')){
      sheet||=newSheet(part.name);sheet.kind='chart';const rels=(await readRelationships(zip,part.path)).filter(r=>!r.type.endsWith('/drawing')).map(r=>({...r,target:r.mode==='External'?r.target:resolvePart(part.path,r.target)}));
      sheet.xlsxChartSheet={xml:await zip.file(part.path).async('string'),...(rels.length?{rels,parts:await capturePartTree(zip,rels.filter(r=>r.mode!=='External').map(r=>r.target))}:{})};
    }
    if(!sheet)continue;if(state!=='visible')sheet.state=state;ordered.push(sheet);
  }
  if(ordered.length)book.sheets=ordered;const active=Number(attr(all(doc,'workbookView')[0],'activeTab','0'));book.activeSheetId=book.sheets[active]?.id||book.sheets[0].id;
}
/** Convert ExcelJS's temporary blank worksheet into an actual chart sheet only
 * after the drawing exporter has attached its chart relationship.
 */
export async function addSheetKindParts(bytes,book,names){
  if(!book.sheets.some(s=>s.kind==='chart'||s.state&&s.state!=='visible'))return bytes;
  const {default:JSZip}=await import('jszip'),zip=await JSZip.loadAsync(bytes),doc=xml(await zip.file('xl/workbook.xml').async('string')),rels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string')),types=xml(await zip.file('[Content_Types].xml').async('string')),content=contentTypeWriter(types);
  let index=0;
  for(const node of all(doc,'sheet')){
    const sheet=book.sheets.find(s=>(names?.get(s.id)||s.name)===attr(node,'name'));if(!sheet)continue;if(sheet.state&&sheet.state!=='visible')node.setAttribute('state',sheet.state);else node.removeAttribute('state');if(sheet.kind!=='chart')continue;
    const rel=all(rels,'Relationship').find(r=>attr(r,'Id')===attr(node,'r:id')),old=resolvePart('xl/workbook.xml',attr(rel,'Target')),path=`xl/chartsheets/sheet${++index}.xml`,source=sheet.xlsxChartSheet;
    const target=source?xml(source.xml):xml(`<chartsheet xmlns="${NS}" xmlns:r="${R}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>`),root=target.documentElement;
    for(const view of all(root,'sheetView'))if(sheet.id===book.activeSheetId)view.setAttribute('tabSelected','1');else view.removeAttribute('tabSelected');
    for(const drawing of children(root,'drawing'))root.removeChild(drawing);
    const drawing=(await readRelationships(zip,old)).find(r=>r.type.endsWith('/drawing')),out=[];
    if(drawing){const ids=new Set((source?.rels||[]).map(r=>r.id));let id='rIdSheetStudioDrawing';while(ids.has(id))id+='1';const node=target.createElementNS(NS,'drawing');node.setAttributeNS(R,'r:id',id);root.insertBefore(node,children(root).find(n=>['legacyDrawing','legacyDrawingHF','picture','webPublishItems','extLst'].includes(n.localName))||null);out.push(relationship(id,drawing.type,'/'+resolvePart(old,drawing.target),drawing.mode));}
    if(source?.parts){const map=p=>`xl/sheetStudioChartSheets/sheet${index}/`+p.replace(/^(?:xl\/sheetStudioChartSheets\/sheet\d+\/)+/,'');await writeStoredParts(zip,source.parts,map,content);for(const r of source.rels||[])out.push(relationship(r.id,r.type,r.mode==='External'?r.target:'/'+map(r.target),r.mode));}
    zip.file(path,stringify(target));zip.file(relPath(path),relationships(out.join('')));zip.remove(old);zip.remove(relPath(old));for(const type of all(types,'Override'))if(attr(type,'PartName')==='/'+old)types.documentElement.removeChild(type);content(path,TYPE);rel.setAttribute('Type',R+'/chartsheet');rel.setAttribute('Target','/'+path);
  }
  zip.file('xl/workbook.xml',stringify(doc));zip.file('xl/_rels/workbook.xml.rels',stringify(rels));zip.file('[Content_Types].xml',stringify(types));return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
