import {validatePreservedXML} from './xlsx-preservation-schema.mjs';
import JSZip from 'jszip';
import {xml,stringify,children,child,all,attr,readRelationships,resolvePart} from './chart-xml.mjs';
import {splitWorksheetXML} from './worksheet-metadata.mjs';

export const plainCopy=value=>JSON.parse(JSON.stringify(value));
export async function worksheetParts(zip){
  const book=xml(await zip.file('xl/workbook.xml').async('string')),rels=await readRelationships(zip,'xl/workbook.xml');
  return all(book,'sheet').map(node=>{const rel=rels.find(r=>r.id===attr(node,'r:id'));return {name:attr(node,'name'),path:rel&&rel.mode!=='External'?resolvePart('xl/workbook.xml',rel.target):null,type:rel?.type};}).filter(item=>item.path);
}
/** Read range validation XML without ExcelJS expanding entire-column rules into
 * millions of JavaScript objects. Keep the original bounds, formulas and flags.
 */
export async function prepareWorksheetFeatures(zip){
  const features=new Map();let stripped=false;
  for(const part of await worksheetParts(zip)){
    if(!part.type?.endsWith('/worksheet'))continue;
    const text=await zip.file(part.path).async('string');
    if(!/<(?:\w+:)?(?:dataValidations|sheetProtection|conditionalFormatting|hyperlinks)\b/.test(text))continue;
    const split=splitWorksheetXML(text),doc=xml(split.metadata),root=doc.documentElement,feature={};
    const validations=child(root,'dataValidations'),protection=child(root,'sheetProtection'),conditions=children(root,'conditionalFormatting');
    if(validations){feature.dataValidations=stringify(validations);root.removeChild(validations);stripped=true;zip.file(part.path,split.restore(stringify(doc)));}
    if(protection)feature.sheetProtection=stringify(protection);
    if(conditions.length)feature.conditionalFormatting=conditions.map(stringify);
    const tips=children(child(root,'hyperlinks'),'hyperlink').filter(n=>attr(n,'tooltip')).map(n=>[attr(n,'ref'),attr(n,'tooltip')]);if(tips.length)feature.hyperlinkTips=Object.fromEntries(tips);
    features.set(part.name,feature);
  }
  const styles=zip.file('xl/styles.xml')?xml(await zip.file('xl/styles.xml').async('string')):null;
  return {features,dxfs:children(child(styles?.documentElement,'dxfs'),'dxf').map(stringify),data:stripped?await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}):null};
}
const order=['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'];
function insert(doc,node){const index=order.indexOf(node.localName),before=children(doc.documentElement).find(n=>order.indexOf(n.localName)>index);doc.documentElement.insertBefore(doc.importNode(node,true),before||null);}
export async function addPreservedFeatures(bytes,book,names){
  if(!book.sheets.some(s=>s.xlsxFeatures))return bytes;
  const zip=await JSZip.loadAsync(bytes);let dxfOffset=0,styles;
  if(book.sheets.some(s=>s.unsupportedConditionalCount&&s.xlsxFeatures?.conditionalFormatting?.length)&&book.xlsxDxfs?.length){
    styles=xml(await zip.file('xl/styles.xml').async('string'));let dxfs=child(styles.documentElement,'dxfs');
    if(!dxfs){dxfs=styles.createElementNS(styles.documentElement.namespaceURI,'dxfs');const before=child(styles.documentElement,'tableStyles')||child(styles.documentElement,'colors')||child(styles.documentElement,'extLst');styles.documentElement.insertBefore(dxfs,before||null);}
    dxfOffset=children(dxfs,'dxf').length;for(const source of book.xlsxDxfs)dxfs.appendChild(styles.importNode(validatePreservedXML(source,'dxf').documentElement,true));dxfs.setAttribute('count',String(children(dxfs,'dxf').length));zip.file('xl/styles.xml',stringify(styles));
  }
  for(const part of await worksheetParts(zip)){
    const sheet=book.sheets.find(s=>(names?.get(s.id)||s.name)===part.name),feature=sheet?.xlsxFeatures;if(!feature)continue;
    const split=splitWorksheetXML(await zip.file(part.path).async('string')),doc=xml(split.metadata);
    for(const name of ['sheetProtection','dataValidations'])if(feature[name]){for(const node of children(doc.documentElement,name))doc.documentElement.removeChild(node);insert(doc,validatePreservedXML(feature[name],name).documentElement);}
    if(sheet.unsupportedConditionalCount&&feature.conditionalFormatting?.length){
      for(const node of children(doc.documentElement,'conditionalFormatting'))doc.documentElement.removeChild(node);
      for(const source of feature.conditionalFormatting){const node=validatePreservedXML(source,'conditionalFormatting').documentElement;for(const rule of all(node,'cfRule'))if(attr(rule,'dxfId')!=='')rule.setAttribute('dxfId',String(Number(attr(rule,'dxfId'))+dxfOffset));insert(doc,node);}
    }
    zip.file(part.path,split.restore(stringify(doc)));
  }
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
