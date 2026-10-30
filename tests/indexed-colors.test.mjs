import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {tableCellStyle} from '../src/table-style.mjs';
import {xml,stringify,child,children} from '../src/chart-xml.mjs';
test('custom indexed palettes remain consistent in ordinary cells, borders, table styles and repeated XLSX saves',async()=>{
  const wb=new ExcelJS.Workbook(),sheet=wb.addWorksheet('Indexed');sheet.addTable({name:'Palette',ref:'A1',headerRow:true,columns:[{name:'Name'}],rows:[['Item']]});sheet.getCell('D1').value=10;sheet.getCell('D1').font={color:{indexed:8}};sheet.getCell('D1').fill={type:'pattern',pattern:'solid',fgColor:{indexed:9}};sheet.getCell('D1').border={bottom:{style:'thin',color:{indexed:8}}};
  let zip=await JSZip.loadAsync(await wb.xlsx.writeBuffer()),styles=xml(await zip.file('xl/styles.xml').async('string')),root=styles.documentElement,NS=root.namespaceURI;
  const colors=xml(`<colors xmlns="${NS}"><indexedColors>${Array.from({length:64},(_,i)=>`<rgbColor rgb="FF${i===8?'123456':i===9?'ABCDEF':'000000'}"/>`).join('')}</indexedColors></colors>`).documentElement;root.appendChild(styles.importNode(colors,true));
  const dxfs=child(root,'dxfs'),dxf=xml(`<dxf xmlns="${NS}"><fill><patternFill patternType="solid"><fgColor indexed="8"/></patternFill></fill></dxf>`).documentElement;dxfs.appendChild(styles.importNode(dxf,true));dxfs.setAttribute('count','1');child(root,'tableStyles').appendChild(styles.importNode(xml(`<tableStyle xmlns="${NS}" name="CustomIndexed" table="1" count="1"><tableStyleElement type="headerRow" dxfId="0"/></tableStyle>`).documentElement,true));child(root,'tableStyles').setAttribute('count','1');zip.file('xl/styles.xml',stringify(styles));zip.file('xl/tables/table1.xml',(await zip.file('xl/tables/table1.xml').async('string')).replace(/name="TableStyleMedium2"/,'name="CustomIndexed"'));let bytes=await zip.generateAsync({type:'uint8array'});
  for(let round=0;round<3;round++){
    const book=await readSpreadsheet(bytes,'indexed.xlsx'),s=book.sheets[0];assert.ok(book.xlsxColors);assert.equal(s.cells.D1.style.color.toLowerCase(),'#123456');assert.equal(s.cells.D1.style.background.toLowerCase(),'#abcdef');assert.equal(s.cells.D1.style.borders.bottom.color.toLowerCase(),'#123456');assert.equal(tableCellStyle(book,s,0,0).background,'#123456');
    s.cells.D1.raw=String(round+11);bytes=(await writeSpreadsheet(book)).data;zip=await JSZip.loadAsync(bytes);styles=xml(await zip.file('xl/styles.xml').async('string'));assert.equal(children(child(child(styles.documentElement,'colors'),'indexedColors'),'rgbColor').length,64);
  }
});
