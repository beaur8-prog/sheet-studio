import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {validateWorkbook,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {worksheetCellStyle} from '../src/cell-style.mjs';
import {cellIsLocked} from '../src/cell-constraints.mjs';
import {structuralWorkbookPlan} from '../src/workbook-structure.mjs';
import {styleDefaultsFixture} from './fixtures/style-defaults.mjs';
import {xml,stringify,all,children,child,attr} from '../src/chart-xml.mjs';
test('Normal, styled empty rows/columns and explicit style-zero cells retain distinct inheritance without allocating blank cells',async()=>{
  const book=validateWorkbook(await readSpreadsheet(await styleDefaultsFixture(),'defaults.xlsx')),s=book.sheets[0];
  assert.ok(Number.isInteger(book.xlsxDefaultStyle));assert.equal(Object.keys(s.cells).length,3);assert.ok(Number.isInteger(s.xlsxStyleDefaults.rows[4]));assert.ok(Number.isInteger(s.xlsxStyleDefaults.columns[1]));
  const normal=worksheetCellStyle(book,s,1,0),column=worksheetCellStyle(book,s,2,1),row=worksheetCellStyle(book,s,4,1),reset=worksheetCellStyle(book,s,3,1);
  assert.equal(normal.fontFamily,'Arial');assert.equal(normal.fontSize,14);assert.equal(normal.color,'#004488');assert.equal(normal.italic,true);
  assert.equal(column.fontSize,12);assert.equal(column.color.toLowerCase(),'#ff0000');assert.equal(column.bold,true);assert.equal(column.italic,false);assert.equal(column.numberFormat,'00000');assert.equal(column.background.toLowerCase(),'#ffffcc');
  assert.equal(row.fontFamily,'Georgia');assert.equal(row.color.toLowerCase(),'#0000ff');assert.equal(row.bold,false);assert.equal(row.italic,true);assert.equal(row.numberFormat,'0.0%');assert.equal(reset.color,'#004488');assert.equal(reset.fontSize,14);assert.equal(s.cells.B4.xlsxNormal,true);
});
test('default and axis styles survive three XLSX roundtrips, new input, ordinary formatting edits and JSON',async()=>{
  let book=await readSpreadsheet(await styleDefaultsFixture(),'defaults.xlsx'),counts;
  for(let round=0;round<3;round++){
    const s=book.sheets[0];s.cells.B3={raw:'12',style:{}};s.cells.C5={raw:'.25',style:{}};s.cells.A2={raw:'new normal',style:{color:'#008800'}};book=parseWorkbook(serializeWorkbook(book));
    const out=await writeSpreadsheet(book),zip=await JSZip.loadAsync(out.data),styles=xml(await zip.file('xl/styles.xml').async('string')),xf=children(child(styles.documentElement,'cellXfs'),'xf')[0],font=children(child(styles.documentElement,'fonts'),'font')[Number(attr(xf,'fontId'))];
    assert.equal(attr(child(font,'name'),'val'),'Arial');assert.equal(attr(child(font,'sz'),'val'),'14');assert.equal(attr(child(font,'color'),'rgb'),'FF004488');assert.ok(child(font,'i'));
    const wb=xml(await zip.file('xl/workbook.xml').async('string'));assert.deepEqual(all(wb,'sheet').map(n=>attr(n,'name')),['Defaults']);assert.ok(!Object.keys(zip.files).some(p=>p.includes('sheet2.xml')));assert.doesNotMatch(await zip.file('docProps/app.xml').async('string'),/_SheetStudio_Normal/);const count=Object.keys(zip.files).filter(p=>!zip.files[p].dir).length;if(counts===undefined)counts=count;else assert.equal(count,counts);
    const excel=new ExcelJS.Workbook();await excel.xlsx.load(out.data);const check=excel.worksheets[0];assert.equal(check.getColumn(2).numFmt,'00000');assert.equal(check.getColumn(2).border.diagonal.style,'double');assert.equal(check.getRow(5).font.name,'Georgia');assert.equal(check.getRow(5).alignment.textRotation,30);assert.equal(check.getCell('B3').numFmt,'00000');assert.equal(check.getCell('C5').numFmt,'0.0%');assert.equal(check.getCell('B3').border.diagonal.up,true);assert.equal(check.getCell('A2').font.size,14);assert.equal(check.getCell('A2').font.italic,true);
    book=validateWorkbook(await readSpreadsheet(out.data,'defaults.xlsx'));assert.equal(worksheetCellStyle(book,book.sheets[0],2,1).color.toLowerCase(),'#ff0000');assert.equal(worksheetCellStyle(book,book.sheets[0],4,2).fontFamily,'Georgia');assert.equal(worksheetCellStyle(book,book.sheets[0],3,1).color,'#004488');
  }
});
test('row and column templates follow bulk structure and invalid style references are refused',async()=>{
  let book=await readSpreadsheet(await styleDefaultsFixture(),'defaults.xlsx');book=structuralWorkbookPlan(book,book.activeSheetId,{axis:'r',index:2,count:3,delta:1});assert.ok(Number.isInteger(book.sheets[0].xlsxStyleDefaults.rows[7]));book=structuralWorkbookPlan(book,book.activeSheetId,{axis:'c',index:0,count:2,delta:1});assert.ok(Number.isInteger(book.sheets[0].xlsxStyleDefaults.columns[3]));assert.equal(worksheetCellStyle(book,book.sheets[0],7,3).fontFamily,'Georgia');
  book.sheets[0].xlsxStyleDefaults.rows[10]=99999;assert.throws(()=>validateWorkbook(book),/서식 참조/);
});
test('Normal protection and style-zero overrides are honored in protected sheets',async()=>{
  const zip=await JSZip.loadAsync(await styleDefaultsFixture()),styles=xml(await zip.file('xl/styles.xml').async('string')),xf=children(child(styles.documentElement,'cellXfs'),'xf')[0],protection=styles.createElementNS(xf.namespaceURI,'protection');protection.setAttribute('locked','0');xf.appendChild(protection);xf.setAttribute('applyProtection','1');zip.file('xl/styles.xml',stringify(styles));
  const sheet=xml(await zip.file('xl/worksheets/sheet1.xml').async('string')),p=sheet.createElementNS(sheet.documentElement.namespaceURI,'sheetProtection');p.setAttribute('sheet','1');sheet.documentElement.insertBefore(p,child(sheet.documentElement,'pageMargins')||null);zip.file('xl/worksheets/sheet1.xml',stringify(sheet));
  const book=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'protected-defaults.xlsx'),s=book.sheets[0];assert.equal(cellIsLocked(book,s,'A2'),false);assert.equal(cellIsLocked(book,s,'B2'),true);assert.equal(cellIsLocked(book,s,'B4'),false);assert.equal(cellIsLocked(book,s,'C5'),false);
});
