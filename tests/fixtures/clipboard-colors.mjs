import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {xml,stringify,all,children} from '../../src/chart-xml.mjs';
export async function clipboardColorFixture(){
 const wb=new ExcelJS.Workbook(),sheet=wb.addWorksheet('Source');sheet.getCell('A1').value=42;sheet.getCell('A1').font={name:'Arial',size:13,color:{theme:4,tint:.2}};sheet.getCell('A1').border={diagonal:{style:'double',up:true,color:{indexed:8}}};sheet.getCell('A1').fill={type:'gradient',gradient:'angle',degree:45,stops:[{position:0,color:{theme:4}},{position:1,color:{indexed:8}}]};sheet.getCell('A2').value={richText:[{text:'Theme',font:{color:{theme:4}}},{text:'Index',font:{color:{indexed:8}}}]};
 const zip=await JSZip.loadAsync(await wb.xlsx.writeBuffer()),theme=xml(await zip.file('xl/theme/theme1.xml').async('string'));children(all(theme,'accent1')[0])[0].setAttribute('val','CC2200');zip.file('xl/theme/theme1.xml',stringify(theme));const styles=xml(await zip.file('xl/styles.xml').async('string')),root=styles.documentElement,NS=root.namespaceURI,colors=xml(`<colors xmlns="${NS}"><indexedColors>${Array.from({length:64},(_,i)=>`<rgbColor rgb="FF${i===8?'0088AA':'000000'}"/>`).join('')}</indexedColors></colors>`).documentElement;root.appendChild(styles.importNode(colors,true));zip.file('xl/styles.xml',stringify(styles));return zip.generateAsync({type:'uint8array'});
}
