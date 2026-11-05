import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,all,children,attr} from '../src/chart-xml.mjs';
import {normalizeXlsxFontOrder} from '../src/xlsx-default-style.mjs';

test('XLSX font properties conform to CT_Font while retaining advanced properties and rich-text runs',async()=>{
 const original=new ExcelJS.Workbook(),sheet=original.addWorksheet('Fonts');sheet.getCell('A1').value='한글';sheet.getCell('A1').font={name:'Arial',size:13,bold:true,italic:true,strike:true,underline:'double',vertAlign:'superscript',color:{argb:'FF224466'},family:2,charset:129,scheme:'minor'};sheet.getCell('A2').value={richText:[{font:{bold:true,color:{argb:'FFFF0000'}},text:'강조 '},{font:{italic:true},text:'보존'}]};
 let bytes=new Uint8Array(await original.xlsx.writeBuffer());
 for(let round=0;round<3;round++){
  const book=await readSpreadsheet(bytes,'fonts.xlsx');bytes=(await writeSpreadsheet(book)).data;const zip=await JSZip.loadAsync(bytes),doc=xml(await zip.file('xl/styles.xml').async('string')),fonts=all(doc,'font'),font=fonts.find(n=>children(n,'name').some(n=>attr(n,'val')==='Arial'));
  assert.deepEqual(children(font).map(n=>n.localName),['b','i','strike','u','vertAlign','sz','color','name','family','charset','scheme']);
  const read=new ExcelJS.Workbook();await read.xlsx.load(bytes);assert.deepEqual(read.worksheets[0].getCell('A1').font,sheet.getCell('A1').font);assert.deepEqual(read.worksheets[0].getCell('A2').value,sheet.getCell('A2').value);
 }
 assert.strictEqual(await normalizeXlsxFontOrder(bytes),bytes,'a schema-ordered archive needs no additional rewrite');
});

test('font normalization retains unknown extension payloads and differential format IDs',async()=>{
 const zip=new JSZip();zip.file('xl/styles.xml','<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:other="urn:test"><fonts count="1"><font><scheme val="minor"/><sz val="12"/><other:extra value="keep"/></font></fonts><dxfs count="1"><dxf><font><color rgb="FF123456"/><sz val="14"/><b/></font></dxf></dxfs></styleSheet>');zip.file('unrelated.bin',new Uint8Array([1,2,3,4]));const result=await JSZip.loadAsync(await normalizeXlsxFontOrder(await zip.generateAsync({type:'uint8array'}))),doc=xml(await result.file('xl/styles.xml').async('string'));
 assert.deepEqual(all(doc,'font').map(n=>children(n).map(n=>n.localName)),[['sz','scheme','extra'],['b','sz','color']]);assert.equal(attr(all(doc,'extra')[0],'value'),'keep');assert.equal(attr(all(doc,'dxfs')[0],'count'),'1');assert.deepEqual(await result.file('unrelated.bin').async('uint8array'),new Uint8Array([1,2,3,4]));
});
