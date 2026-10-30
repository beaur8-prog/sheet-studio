import test from 'node:test';
import {chartsheetBook} from './fixtures/sheet-kinds.mjs';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createEvaluator,validateWorkbook} from '../src/engine.mjs';
import {chartData} from '../src/chart-model.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {attachNativeSource,buildNativePlan,nativeSourceChanged} from '../src/native-plan.mjs';
import {xml,stringify,all,attr,readRelationships} from '../src/chart-xml.mjs';
test('chart sheets retain their type, active tab, source data, hidden states, local names and print metadata through repeated XLSX saves',async()=>{
  let book=chartsheetBook(),bytes=(await writeSpreadsheet(book)).data,zip=await JSZip.loadAsync(bytes);zip.file('xl/chartsheets/sheet1.xml',(await zip.file('xl/chartsheets/sheet1.xml').async('string')).replace('<drawing','<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><headerFooter><oddHeader>Confidential</oddHeader></headerFooter><drawing'));
  bytes=await zip.generateAsync({type:'uint8array'});let counts;
  for(let round=0;round<3;round++){
    book=validateWorkbook(await readSpreadsheet(bytes,'dashboard.xlsx'));assert.deepEqual(book.sheets.map(s=>s.name),['Dashboard','Source','Tail']);assert.equal(book.sheets[0].kind,'chart');assert.equal(book.activeSheetId,book.sheets[0].id);assert.equal(book.sheets[1].state,'hidden');assert.equal(book.chartNames[0].sheetId,book.sheets[2].id);assert.equal(createEvaluator(book,book.sheets[2].id)('A1'),20);assert.deepEqual(chartData(book.sheets[0].charts[0],book.sheets[0],book)[0].values,[10,20]);assert.match(book.sheets[0].xlsxChartSheet.xml,/Confidential/);
    bytes=(await writeSpreadsheet(book)).data;zip=await JSZip.loadAsync(bytes);const rels=await readRelationships(zip,'xl/workbook.xml');assert.equal(rels.filter(r=>r.type.endsWith('/chartsheet')).length,1);assert.equal(rels.filter(r=>r.type.endsWith('/worksheet')).length,2);const current=Object.keys(zip.files).filter(p=>!zip.files[p].dir).length;if(!counts)counts=current;else assert.equal(current,counts);const names=all(xml(await zip.file('xl/workbook.xml').async('string')),'definedName');assert.equal(attr(names.find(n=>attr(n,'name')==='Local'),'localSheetId'),'2');
  }
  await assert.rejects(writeSpreadsheet(book,'csv'),/차트 전용/);book.sheets[0].cells.A1={raw:'1'};assert.throws(()=>validateWorkbook(book),/차트 전용/);
});
test('native save plans keep chart sheets in the original file and refuse unsupported edits instead of treating them as worksheets',async()=>{
  const book=chartsheetBook();attachNativeSource(book,(await writeSpreadsheet(book)).data,'dashboard.xlsm');assert.equal(nativeSourceChanged(book),false);const plan=buildNativePlan(book);assert.equal(plan.sheets.length,2);assert.ok(!plan.sheets.some(s=>s.name==='Dashboard'));book.sheets[0].charts[0].title='changed';assert.throws(()=>buildNativePlan(book),/차트 전용/);
});
