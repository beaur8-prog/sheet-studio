import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {externalLinksFixture} from './fixtures/external-links.mjs';
import {createEvaluator,validateWorkbook,serializeWorkbook,parseWorkbook,formulaHasExternalReferences,createWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {createHistory} from '../src/history.mjs';
import {calculationDisplay} from '../src/calculation-display.mjs';
import {assertCellUpdates} from '../src/cell-constraints.mjs';
import {xml,stringify,all,child,attr,readRelationships,resolvePart,R} from '../src/chart-xml.mjs';

async function links(bytes){const zip=await JSZip.loadAsync(bytes),book=xml(await zip.file('xl/workbook.xml').async('string')),rels=await readRelationships(zip,'xl/workbook.xml'),out=[];for(const ref of all(book,'externalReference')){const rel=rels.find(r=>r.id===ref.getAttributeNS(R,'id')),path=resolvePart('xl/workbook.xml',rel.target),source=await zip.file(path).async('string'),targets=await readRelationships(zip,path);out.push({path,source,targets});}return out;}

test('external workbook definitions, ordered indices, cache types and targets survive JSON and repeated XLSX round trips',async()=>{
 let bytes=await externalLinksFixture();for(let round=0;round<3;round++){
  const b=parseWorkbook(serializeWorkbook(validateWorkbook(await readSpreadsheet(bytes,'external.xlsx')))),s=b.sheets[0],e=createEvaluator(b,s.id);assert.equal(b.xlsxExternalLinks.rels.length,2);assert.equal(e('A1'),'#NAME?');assert.equal(e.result('A1').cacheSource,'external');assert.equal(calculationDisplay(e.result('A1')).value,42);assert.equal(calculationDisplay(e.result('A2')).value,true);assert.equal(calculationDisplay(e.result('A3')).value,'001');assert.equal(calculationDisplay(e.result('A4')).error,null);assert.equal(calculationDisplay(e.result('A4')).value,'#N/A');assert.equal(calculationDisplay(e.result('A5')).error,'#N/A');assert.equal(calculationDisplay(e.result('C1')).value,84);
  bytes=(await writeSpreadsheet(b)).data;const out=await links(bytes);assert.equal(out.length,2);assert.equal(out[0].targets[0].target,'file:///virtual/source.xlsx');assert.equal(out[1].targets[0].target,'https://example.invalid/source.xlsx');assert.equal(out[0].targets[0].mode,'External');assert.match(out[0].source,/OriginalName/);assert.match(out[0].source,/>42</);assert.ok(!out[0].path.includes('sheetStudioExternalLinks/xl/sheetStudioExternalLinks'));
  const check=new ExcelJS.Workbook();await check.xlsx.load(bytes);assert.equal(check.worksheets[0].getCell('A1').formula,"'[1]Data'!A1");assert.equal(check.worksheets[0].getCell('A1').result,42);assert.equal(check.worksheets[0].getCell('A4').result,'#N/A');assert.deepEqual(check.worksheets[0].getCell('A5').result,{error:'#N/A'});
 }
});
test('raw edits invalidate unsupported caches atomically while XLSX retains the original payload and stale provenance',async()=>{
 const h=createHistory(await readSpreadsheet(await externalLinksFixture(),'external.xlsx'));h.begin();h.book.sheets[0].cells.E1.raw='changed';h.commit();let b=h.raw,e=createEvaluator(b,b.activeSheetId);assert.equal(e.result('A1').stale,true);assert.equal(calculationDisplay(e.result('A1')).value,'#NAME?');
 const csv=new TextDecoder().decode((await writeSpreadsheet(b,'csv')).data);assert.match(csv,/#NAME\?/);let bytes=(await writeSpreadsheet(b)).data;const check=new ExcelJS.Workbook();await check.xlsx.load(bytes);assert.equal(check.worksheets[0].getCell('A1').result,42);assert.equal(check.worksheets[0].getCell('C1').result,84);assert.equal((await links(bytes)).length,2);
 b=await readSpreadsheet(bytes,'stale.xlsx');assert.equal(b.referenceCachesValid,false);assert.equal(calculationDisplay(createEvaluator(b,b.activeSheetId).result('A1')).value,'#NAME?');h.undo();assert.equal(calculationDisplay(createEvaluator(h.raw,h.raw.activeSheetId).result('A1')).value,42);
 const zip=await JSZip.loadAsync(bytes),core=xml(await zip.file('docProps/core.xml').async('string'));all(core,'lastModifiedBy')[0].textContent='Another editor';zip.file('docProps/core.xml',stringify(core));b=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'external-owner.xlsx');assert.equal(b.referenceCachesValid,undefined);
});
test('external markers skip strings, refuse missing definitions and reject broken relationships rather than dropping them',async()=>{
 for(const formula of ["='[1]Data'!A1",'=[1]Data!A1',"=SUM('[Book.xlsx]Data'!A1)"])assert.equal(formulaHasExternalReferences(formula),true,formula);for(const formula of ['="[1]Data!A1"','=SUM(A1:A3)','=Table1[Column]'])assert.equal(formulaHasExternalReferences(formula),false,formula);
 const b=createWorkbook();assert.throws(()=>assertCellUpdates(b,b.sheets[0],[{key:'A1',cell:{raw:"='[1]Data'!A1"}}]),/외부/);
 const zip=await JSZip.loadAsync(await externalLinksFixture()),rels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));const bad=all(rels,'Relationship').find(n=>attr(n,'Id')==='rIdExternal1');bad.parentNode.removeChild(bad);zip.file('xl/_rels/workbook.xml.rels',stringify(rels));const badBytes=await zip.generateAsync({type:'uint8array'});await assert.rejects(()=>readSpreadsheet(badBytes,'bad.xlsx'),/외부 통합 문서 연결/);
});
test('stored unsupported results never replace fresh arithmetic errors or text results after the formula changes',async()=>{
 const source=new ExcelJS.Workbook(),s=source.addWorksheet('Future');s.getCell('A1').value={formula:'_xlfn.FUTUREFOO()',result:7};let b=await readSpreadsheet(await source.xlsx.writeBuffer(),'future.xlsx'),cell=b.sheets[0].cells.A1;assert.equal(calculationDisplay(createEvaluator(b,b.activeSheetId).result('A1')).value,7);cell.raw='=1/0';assert.equal(calculationDisplay(createEvaluator(b,b.activeSheetId).result('A1')).value,'#DIV/0!');cell.raw='="#NAME?"';assert.equal(calculationDisplay(createEvaluator(b,b.activeSheetId).result('A1')).error,null);
});
