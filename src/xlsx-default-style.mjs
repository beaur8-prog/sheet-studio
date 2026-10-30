import JSZip from 'jszip';
import {xml,stringify,children,child,all,attr,resolvePart,relPath,relationships,relationship,R} from './chart-xml.mjs';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',P='http://schemas.openxmlformats.org/package/2006/content-types';
export const normalDigitWidth=font=>Math.max(5,Math.round((/맑은 고딕|Malgun|Arial|Aptos/i.test(font?.name||'')?8:7)*(Number(font?.size)||11)/11));
/** ExcelJS skips cell style 0. Read it through a tiny public-API workbook whose
 * one cell refers to a duplicate XF. Never add a fake cell to the user's sheet.
 */
export async function readDefaultStyleModel(Workbook,stylesXML,themeXML){
  if(!stylesXML)return {};
  const doc=xml(stylesXML),xfs=child(doc.documentElement,'cellXfs'),first=children(xfs,'xf')[0];if(!first)return {};
  const id=children(xfs,'xf').length;xfs.appendChild(first.cloneNode(true));xfs.setAttribute('count',String(id+1));
  const zip=new JSZip();zip.file('[Content_Types].xml','<Types xmlns="'+P+'"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'+(themeXML?'<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>':'')+'</Types>');
  zip.file('_rels/.rels',relationships(relationship('rId1',R+'/officeDocument','xl/workbook.xml')));zip.file('xl/workbook.xml','<workbook xmlns="'+S+'" xmlns:r="'+R+'"><sheets><sheet name="Style" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels',relationships(relationship('rId1',R+'/worksheet','worksheets/sheet1.xml')+relationship('rId2',R+'/styles','styles.xml')+(themeXML?relationship('rId3',R+'/theme','theme/theme1.xml'):'')));
  zip.file('xl/styles.xml',stringify(doc));if(themeXML)zip.file('xl/theme/theme1.xml',themeXML);
  zip.file('xl/worksheets/sheet1.xml','<worksheet xmlns="'+S+'"><sheetData><row r="1"><c r="A1" s="'+id+'"><v>1</v></c></row></sheetData></worksheet>');
  const book=new Workbook();await book.xlsx.load(await zip.generateAsync({type:'uint8array'}));return book.worksheets[0].getCell('A1').style;
}
/** Serialize the original default with ExcelJS, then make its generated XF the
 * actual Normal/default XF and remove the temporary export-only worksheet.
 */
export async function applyDefaultStyleProbe(bytes,name){
  if(!name)return bytes;
  const zip=await JSZip.loadAsync(bytes),wb=xml(await zip.file('xl/workbook.xml').async('string')),entry=all(wb,'sheet').find(n=>attr(n,'name')===name);if(!entry)throw Error('기본 서식 저장 정보를 찾을 수 없습니다.');
  const relDoc=xml(await zip.file(relPath('xl/workbook.xml')).async('string')),rel=all(relDoc,'Relationship').find(n=>attr(n,'Id')===attr(entry,'r:id')),path=resolvePart('xl/workbook.xml',attr(rel,'Target')),sheet=xml(await zip.file(path).async('string')),styleId=Number(attr(all(sheet,'c')[0],'s','0')),styles=xml(await zip.file('xl/styles.xml').async('string')),xfs=child(styles.documentElement,'cellXfs'),source=children(xfs,'xf')[styleId];if(!source)throw Error('기본 셀 서식을 저장할 수 없습니다.');
  xfs.replaceChild(source.cloneNode(true),children(xfs,'xf')[0]);const styleXfs=child(styles.documentElement,'cellStyleXfs');if(styleXfs&&children(styleXfs,'xf')[0]){const normal=source.cloneNode(true);normal.removeAttribute('xfId');styleXfs.replaceChild(normal,children(styleXfs,'xf')[0]);}
  entry.parentNode.removeChild(entry);rel.parentNode.removeChild(rel);zip.remove(path);zip.remove(relPath(path));
  const appPart=zip.file('docProps/app.xml');if(appPart){const app=xml(await appPart.async('string'));for(const vector of all(app,'vector')){for(const node of children(vector))if(node.localName==='lpstr'&&node.textContent===name)vector.removeChild(node);if(attr(vector,'baseType')==='lpstr')vector.setAttribute('size',String(children(vector).length));}
    const pairs=children(child(all(app,'HeadingPairs')[0],'vector'));for(let i=0;i+1<pairs.length;i++)if(all(pairs[i],'lpstr').some(n=>n.textContent==='Worksheets')){const count=all(pairs[i+1],'i4')[0];if(count)count.textContent=String(Math.max(0,Number(count.textContent)-1));}zip.file('docProps/app.xml',stringify(app));}
  const types=xml(await zip.file('[Content_Types].xml').async('string'));for(const node of all(types,'Override'))if(attr(node,'PartName')==='/'+path)node.parentNode.removeChild(node);
  zip.file('xl/styles.xml',stringify(styles));zip.file('xl/workbook.xml',stringify(wb));zip.file(relPath('xl/workbook.xml'),stringify(relDoc));zip.file('[Content_Types].xml',stringify(types));return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
