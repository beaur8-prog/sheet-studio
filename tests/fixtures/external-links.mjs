import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {xml,stringify,child} from '../../src/chart-xml.mjs';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** Synthetic links only; no referenced workbook is opened or downloaded. */
export async function externalLinksFixture(){
 const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Input');
 const values=[42,true,'001','#N/A',{error:'#N/A'}];for(let i=0;i<values.length;i++)sheet.getCell('A'+(i+1)).value={formula:"'["+(i<3?1:2)+"]Data'!A"+(i+1),result:values[i]};sheet.getCell('C1').value={formula:'A1*2',result:84};sheet.getCell('E1').value='Local';
 const zip=await JSZip.loadAsync(await book.xlsx.writeBuffer()),workbook=xml(await zip.file('xl/workbook.xml').async('string')),refs=workbook.createElementNS(S,'externalReferences');
 for(let i=1;i<=2;i++){const ref=workbook.createElementNS(S,'externalReference');ref.setAttributeNS(R,'r:id','rIdExternal'+i);refs.appendChild(ref);}
 workbook.documentElement.insertBefore(refs,child(workbook.documentElement,'definedNames')||child(workbook.documentElement,'calcPr')||null);zip.file('xl/workbook.xml',stringify(workbook));
 const rels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string')),types=xml(await zip.file('[Content_Types].xml').async('string'));
 for(let i=1;i<=2;i++){
  const rel=rels.createElementNS(rels.documentElement.namespaceURI,'Relationship');rel.setAttribute('Id','rIdExternal'+i);rel.setAttribute('Type',R+'/externalLink');rel.setAttribute('Target','externalLinks/externalLink'+i+'.xml');rels.documentElement.appendChild(rel);
  const part='xl/externalLinks/externalLink'+i+'.xml';zip.file(part,'<externalLink xmlns="'+S+'" xmlns:r="'+R+'"><externalBook r:id="rIdSource"><sheetNames><sheetName val="Data"/></sheetNames><definedNames><definedName name="OriginalName" refersTo="Data!$A$1"/></definedNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1" t="n"><v>'+ (i===1?42:7) +'</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>');
  zip.file('xl/externalLinks/_rels/externalLink'+i+'.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSource" Type="'+R+'/externalLinkPath" Target="'+(i===1?'file:///virtual/source.xlsx':'https://example.invalid/source.xlsx')+'" TargetMode="External"/></Relationships>');
  const type=types.createElementNS(types.documentElement.namespaceURI,'Override');type.setAttribute('PartName','/'+part);type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml');types.documentElement.appendChild(type);
 }
 zip.file('xl/_rels/workbook.xml.rels',stringify(rels));zip.file('[Content_Types].xml',stringify(types));return zip.generateAsync({type:'uint8array'});
}
