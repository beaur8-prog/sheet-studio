import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {validateWorkbook,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {structuralWorkbookPlan} from '../src/workbook-structure.mjs';
import {printOrientation,setPrintOrientation} from '../src/worksheet-extra.mjs';
import {worksheetExtrasFixture,printerBytes} from './fixtures/worksheet-extras.mjs';
import {chartsheetBook} from './fixtures/sheet-kinds.mjs';
import {xml,all,child,attr,readRelationships,resolvePart,relPath} from '../src/chart-xml.mjs';
test('worksheet print settings, header/footer pictures, printer binaries and original document properties survive repeated saves',async()=>{
  let bytes=await worksheetExtrasFixture(),count;
  for(let round=0;round<3;round++){
    let book=validateWorkbook(await readSpreadsheet(bytes,'print.xlsx'));book=parseWorkbook(serializeWorkbook(book));const s=book.sheets[0];assert.equal(printOrientation(s),'landscape');assert.equal(book.xlsxProperties.creator,'Original Author');assert.equal(book.xlsxProperties.created,'2018-04-03T02:01:00.000Z');s.cells.A2.raw=String(round+4);bytes=(await writeSpreadsheet(book)).data;
    const zip=await JSZip.loadAsync(bytes),doc=xml(await zip.file('xl/worksheets/sheet1.xml').async('string')),setup=child(doc.documentElement,'pageSetup'),rels=await readRelationships(zip,'xl/worksheets/sheet1.xml');assert.equal(attr(setup,'orientation'),'landscape');assert.equal(attr(setup,'fitToWidth'),'1');assert.equal(attr(child(doc.documentElement,'pageMargins'),'right'),'0.4');assert.equal(all(doc,'oddHeader')[0].textContent,'&LCompany&CReport&R&P / &N');assert.equal(attr(child(doc.documentElement,'sheetPr'),'codeName'),'OriginalCode');assert.equal(attr(all(doc,'tabColor')[0],'rgb'),'FF008800');assert.equal(attr(all(doc,'ignoredError')[0],'sqref'),'A2:C20');
    const printer=rels.find(r=>r.id===attr(setup,'r:id')),printerPath=resolvePart('xl/worksheets/sheet1.xml',printer.target);assert.deepEqual(await zip.file(printerPath).async('uint8array'),printerBytes);
    const header=rels.find(r=>r.id===attr(all(doc,'legacyDrawingHF')[0],'r:id')),headerPath=resolvePart('xl/worksheets/sheet1.xml',header.target),picture=(await readRelationships(zip,headerPath))[0];assert.ok(zip.file(resolvePart(headerPath,picture.target)));assert.match(await zip.file(headerPath).async('string'),/HeaderPicture/);
    const current=Object.keys(zip.files).filter(p=>!zip.files[p].dir).length;if(count===undefined)count=current;else assert.equal(current,count);
    const excel=new ExcelJS.Workbook();await excel.xlsx.load(bytes);assert.equal(excel.creator,'Original Author');assert.equal(excel.created.toISOString(),'2018-04-03T02:01:00.000Z');assert.equal(excel.lastModifiedBy,'Sheet Studio');assert.equal(excel.company,'Example company');assert.equal(excel.worksheets[0].getCell('B2').result,(round+4)*3);
  }
});
test('orientation changes preserve original printer links and page breaks/ignored ranges follow bulk edits',async()=>{
  let book=await readSpreadsheet(await worksheetExtrasFixture(),'print.xlsx'),s=book.sheets[0];setPrintOrientation(s,'portrait');assert.equal(printOrientation(s),'portrait');book=structuralWorkbookPlan(book,s.id,{axis:'r',index:1,count:3,delta:1});s=book.sheets[0];assert.equal(attr(all(xml(s.xlsxFeatures.rowBreaks),'brk')[0],'id'),'13');assert.equal(attr(all(xml(s.xlsxFeatures.ignoredErrors),'ignoredError')[0],'sqref'),'A5:C23');
  book=structuralWorkbookPlan(book,s.id,{axis:'c',index:0,count:2,delta:1});s=book.sheets[0];assert.equal(attr(all(xml(s.xlsxFeatures.colBreaks),'brk')[0],'id'),'4');const out=await readSpreadsheet((await writeSpreadsheet(book)).data,'changed.xlsx');assert.equal(printOrientation(out.sheets[0]),'portrait');assert.equal(out.sheets[0].xlsxFeatures.linkedParts.rels.length,2);
});
test('chart-sheet orientation overrides survive XLSX export and invalid feature metadata is rejected',async()=>{
  let book=chartsheetBook(),sheet=book.sheets[0];setPrintOrientation(sheet,'landscape');book=await readSpreadsheet((await writeSpreadsheet(book)).data,'chart.xlsx');assert.equal(book.sheets[0].kind,'chart');assert.equal(printOrientation(book.sheets[0]),'landscape');setPrintOrientation(book.sheets[0],'portrait');book=await readSpreadsheet((await writeSpreadsheet(book)).data,'chart.xlsx');assert.equal(printOrientation(book.sheets[0]),'portrait');
  book.sheets[1].xlsxFeatures.pageSetup='<headerFooter/>';assert.throws(()=>validateWorkbook(book),/XML 종류/);
});
