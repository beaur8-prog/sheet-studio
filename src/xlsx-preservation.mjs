import {mergePreservedDxfs,remapDxfIds} from './xlsx-dxfs.mjs';
import {validatePreservedXML} from './xlsx-preservation-schema.mjs';
import JSZip from 'jszip';
import {xml,stringify,children,child,all,attr,readRelationships,resolvePart,relPath} from './chart-xml.mjs';
import {splitWorksheetXML} from './worksheet-metadata.mjs';
import {captureExtraSheetFeatures,writeExtraSheetRelationships,extraSheetNode,EXTRA_SHEET_TAGS} from './worksheet-extra.mjs';
import {contentTypeWriter} from './xlsx-part-store.mjs';

export const plainCopy=value=>JSON.parse(JSON.stringify(value));
export async function worksheetParts(zip){
  const book=xml(await zip.file('xl/workbook.xml').async('string')),rels=await readRelationships(zip,'xl/workbook.xml');
  return all(book,'sheet').map(node=>{const rel=rels.find(r=>r.id===attr(node,'r:id'));return {name:attr(node,'name'),path:rel&&rel.mode!=='External'?resolvePart('xl/workbook.xml',rel.target):null,type:rel?.type};}).filter(item=>item.path);
}
/** Read range validation XML without ExcelJS expanding entire-column rules into
 * millions of JavaScript objects. Keep the original bounds, formulas and flags.
 */
export async function prepareWorksheetFeatures(zip){
  const features=new Map(),styleRows=new Map(),normalCells=new Map(),loader=zip.clone();loader.files={...zip.files};let stripped=false,hasTables=false;
  // XML boolean 'true' is valid, but ExcelJS 4.4 only recognizes '1' here.
  // Normalize the private loader before it converts any date serials to Date.
  const workbook=xml(await zip.file('xl/workbook.xml').async('string')),properties=child(workbook.documentElement,'workbookPr');
  if(attr(properties,'date1904')==='true'){properties.setAttribute('date1904','1');loader.file('xl/workbook.xml',stringify(workbook));stripped=true;}
  for(const part of await worksheetParts(zip)){
    if(!part.type?.endsWith('/worksheet'))continue;
    const text=await zip.file(part.path).async('string');
    const rows=[],normal=new Set(),columns=new Set();let styledRow=false;
    const split=splitWorksheetXML(text,{onColumn:col=>{if(Number(col.style)>0)for(let c=Math.max(1,Number(col.min));c<=Math.min(16384,Number(col.max));c++)columns.add(c);},onAnyRow:row=>{styledRow=row.s!==undefined;if(styledRow)rows.push(Number(row.r));},wantCell:cell=>{if(cell.s!==undefined&&cell.s!=='0')return false;const name=/^[A-Z]+/i.exec(cell.r||'')?.[0];let c=0;for(const letter of name||'')c=c*26+letter.toUpperCase().charCodeAt(0)-64;return styledRow||columns.has(c);},onCell:cell=>normal.add(cell.r)}),doc=xml(split.metadata),root=doc.documentElement,feature={};
    if(rows.length)styleRows.set(part.name,rows);if(normal.size)normalCells.set(part.name,normal);
    const validations=child(root,'dataValidations'),protection=child(root,'sheetProtection'),conditions=children(root,'conditionalFormatting');
    if(validations){feature.dataValidations=stringify(validations);root.removeChild(validations);stripped=true;}
    const tableParts=child(root,'tableParts');if(tableParts){root.removeChild(tableParts);stripped=true;hasTables=true;}
    const drawing=child(root,'drawing');if(drawing){root.removeChild(drawing);stripped=true;for(const rel of await readRelationships(zip,part.path))if(rel.type.endsWith('/drawing')&&rel.mode!=='External'){const path=resolvePart(part.path,rel.target);loader.remove(path);loader.remove(relPath(path));}}
    if(validations||tableParts||drawing)loader.file(part.path,split.restore(stringify(doc)));
    if(protection)feature.sheetProtection=stringify(protection);
    if(conditions.length)feature.conditionalFormatting=conditions.map(stringify);
    await captureExtraSheetFeatures(zip,part.path,root,feature);
    const tips=children(child(root,'hyperlinks'),'hyperlink').filter(n=>attr(n,'tooltip')).map(n=>[attr(n,'ref'),attr(n,'tooltip')]);if(tips.length)feature.hyperlinkTips=Object.fromEntries(tips);
    features.set(part.name,feature);
  }
  if(hasTables)for(const path of Object.keys(loader.files))if(/(?:^|\/)tables\/[^/]+\.xml$/.test(path))loader.remove(path);
  const styles=zip.file('xl/styles.xml')?xml(await zip.file('xl/styles.xml').async('string')):null;
  return {features,styleRows,normalCells,dxfs:children(child(styles?.documentElement,'dxfs'),'dxf').map(stringify),...(child(styles?.documentElement,'colors')?{colors:stringify(child(styles.documentElement,'colors'))}:{}),data:stripped?await loader.generateAsync({type:'uint8array',compression:'DEFLATE'}):null};
}
const order=['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'];
function insert(doc,node){const index=order.indexOf(node.localName),before=children(doc.documentElement).find(n=>order.indexOf(n.localName)>index);doc.documentElement.insertBefore(doc.importNode(node,true),before||null);}
export async function addPreservedFeatures(bytes,book,names){
  if(!book.sheets.some(s=>s.xlsxFeatures)&&!book.xlsxColors)return bytes;
  const zip=await JSZip.loadAsync(bytes),dxfMap=book.sheets.some(s=>s.unsupportedConditionalCount&&s.xlsxFeatures?.conditionalFormatting?.length)?await mergePreservedDxfs(zip,book.xlsxDxfs):[],types=xml(await zip.file('[Content_Types].xml').async('string')),content=contentTypeWriter(types);
  for(const [index,part] of (await worksheetParts(zip)).entries()){
    const sheet=book.sheets.find(s=>(names?.get(s.id)||s.name)===part.name),feature=sheet?.xlsxFeatures;if(!feature)continue;
    const split=splitWorksheetXML(await zip.file(part.path).async('string')),doc=xml(split.metadata);
    const ids=await writeExtraSheetRelationships(zip,part.path,feature,index,content);
    for(const name of EXTRA_SHEET_TAGS)if(feature[name]){for(const node of children(doc.documentElement,name))doc.documentElement.removeChild(node);const node=extraSheetNode(feature,name,ids);if(name==='sheetPr'){const filters=[sheet.autoFilter,...(sheet.tables||[]).map(t=>t.filter)].filter(Boolean);node.setAttribute('filterMode',filters.some(f=>Object.keys(f.columns||{}).length||f.unsupported?.length)?'1':'0');}insert(doc,node);}
    for(const name of ['sheetProtection','dataValidations'])if(feature[name]){for(const node of children(doc.documentElement,name))doc.documentElement.removeChild(node);insert(doc,validatePreservedXML(feature[name],name).documentElement);}
    if(sheet.unsupportedConditionalCount&&feature.conditionalFormatting?.length){
      for(const node of children(doc.documentElement,'conditionalFormatting'))doc.documentElement.removeChild(node);
      for(const source of feature.conditionalFormatting){const node=validatePreservedXML(source,'conditionalFormatting').documentElement;remapDxfIds(node,dxfMap);insert(doc,node);}
    }
    zip.file(part.path,split.restore(stringify(doc)));
  }
  if(book.xlsxColors){const styles=xml(await zip.file('xl/styles.xml').async('string')),root=styles.documentElement,old=child(root,'colors');if(old)root.removeChild(old);root.insertBefore(styles.importNode(validatePreservedXML(book.xlsxColors,'colors').documentElement,true),child(root,'extLst')||null);zip.file('xl/styles.xml',stringify(styles));}
  zip.file('[Content_Types].xml',stringify(types));
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
