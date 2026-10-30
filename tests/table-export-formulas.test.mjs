import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {formulaForXlsx,createEvaluator} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {tableFixture} from './tables-fixture.mjs';
const table={name:'Sales',columns:['Qty','Price',"O'Clock",'Net Sales','@Flag','Part[A]'].map(name=>({name}))};
test('XLSX formulas qualify implicit table selectors and use explicit current-row syntax without changing strings or external links',()=>{
 for(const [input,expected] of [
  ['=SUM([Qty])','=SUM(Sales[Qty])'],
  ['=[@Qty]*[@Price]','=Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]'],
  ['=SUM([@[Qty]:[Price]])','=SUM(Sales[[#This Row],[Qty]:[Price]])'],
  ["=[@[O''Clock]]","=Sales[[#This Row],[O''Clock]]"],
  ["=SUM(['@Flag])+SUM([Part'[A']])","=SUM(Sales['@Flag])+SUM(Sales[Part'[A']])"],
  ['=Other[@[Net Sales]]','=Other[[#This Row],[Net Sales]]'],
  ['="[@Qty]"&"Sales[@Price]"','="[@Qty]"&"Sales[@Price]"'],
  ['=SUM([Missing])','=SUM([Missing])'],
  ['=[1]Sheet1!A1','=[1]Sheet1!A1'],
 ])assert.equal(formulaForXlsx(input,undefined,table),expected,input);
 assert.equal(formulaForXlsx('=[Qty]'),'=[Qty]');
});
test('future-function prefixes and parameter names survive table-reference canonicalization',()=>{
 const formula='=LET(q,[@Qty],q*[@Price])',original='_xlfn.LET(_xlpm.q,[@Qty],_xlpm.q*[@Price])';
 assert.equal(formulaForXlsx(formula,original,table),'=_xlfn.LET(q,Sales[[#This Row],[Qty]],q*Sales[[#This Row],[Price]])');
 assert.equal(formulaForXlsx('='+original,original,table),'=_xlfn.LET(_xlpm.q,Sales[[#This Row],[Qty]],_xlpm.q*Sales[[#This Row],[Price]])');
});
test('worksheet and calculated-column formulas agree through repeated XLSX saves including a preserved future formula',async()=>{
 let book=await readSpreadsheet(await tableFixture(),'structured.xlsx'),s=book.sheets[0];s.cells.C2.raw='=LET(q,[@Qty],q*[@Price])';s.cells.C2.xlsxFormula={raw:s.cells.C2.raw,formula:'_xlfn.LET(_xlpm.q,[@Qty],_xlpm.q*[@Price])'};
 for(let round=0;round<3;round++){
  const data=(await writeSpreadsheet(book)).data,excel=new ExcelJS.Workbook();await excel.xlsx.load(data);const formula=excel.worksheets[0].getCell('C2').formula;assert.match(formula,/_xlfn.LET/);assert.match(formula,/Sales\[\[#This Row\],\[Qty\]\]/);assert.equal(excel.worksheets[0].getCell('C2').result,6);
  const zip=await JSZip.loadAsync(data);assert.match(await zip.file('xl/tables/table1.xml').async('string'),/Sales\[\[#This Row\],\[Price\]\]/);
  book=await readSpreadsheet(data,'roundtrip.xlsx');s=book.sheets[0];assert.equal(createEvaluator(book,s.id)('C2'),6);assert.equal(createEvaluator(book,s.id)('E1'),26);
 }
});
