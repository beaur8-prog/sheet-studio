import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {R,D,A,relationships,relationship} from '../../src/chart-xml.mjs';
export const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==';
export async function drawingFixture(){
  const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Pictures');sheet.getCell('A1').value='그림과 도형';sheet.getCell('A2').value=5;const id=book.addImage({base64:png,extension:'png'});sheet.addImage(id,{tl:{col:3,row:2},ext:{width:140,height:80}});
  const zip=await JSZip.loadAsync(await book.xlsx.writeBuffer()),part='xl/drawings/drawing1.xml';let drawing=await zip.file(part).async('string');
  drawing=drawing.replace(/r:embed=/g,'link:embed=').replace('<xdr:wsDr ',`<xdr:wsDr xmlns:link="${R}" `);
  const shape=`<xdr:absoluteAnchor xmlns:xdr="${D}" xmlns:a="${A}" xmlns:r="${R}"><xdr:pos x="4000000" y="500000"/><xdr:ext cx="1600000" cy="500000"/><xdr:sp macro=""><xdr:nvSpPr><xdr:cNvPr id="99" name="Preserved shape"><a:hlinkClick r:id="rIdLink"/></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F0C040"/></a:solidFill></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>원본 도형</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:absoluteAnchor>`;
  zip.file(part,drawing.replace('</xdr:wsDr>',shape+'</xdr:wsDr>'));const rels='xl/drawings/_rels/drawing1.xml.rels';zip.file(rels,(await zip.file(rels).async('string')).replace('</Relationships>',relationship('rIdLink',R+'/hyperlink','https://example.com','External')+'</Relationships>'));return zip.generateAsync({type:'uint8array'});
}
