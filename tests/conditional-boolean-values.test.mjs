import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {createWorkbook,createEvaluator} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';

function fixture(){
 const book=createWorkbook(),sheet=book.sheets[0];
 sheet.cells={A1:{raw:'1'},A2:{raw:'1'},A3:{raw:'1'},A4:{raw:'1'},A5:{raw:'1'},A6:{raw:'0'},B1:{raw:'TRUE'},B2:{raw:'FALSE'},B3:{raw:'3'},B4:{raw:"'9"},B5:{raw:''},B6:{raw:'=NA()'}};
 return {book,sheet};
}
test('SUMIFS and AVERAGEIFS include matching Boolean result values as 1 and 0 in bounded and sparse ranges',()=>{
 const {book,sheet}=fixture();
 for(const [i,formula] of ['=SUMIFS(B1:B6,A1:A6,1)','=SUMIFS(B:B,A:A,1)','=AVERAGEIFS(B1:B6,A1:A6,1)','=AVERAGEIFS(B:B,A:A,1)'].entries())sheet.cells['D'+(i+1)]={raw:formula};
 const e=createEvaluator(book,sheet.id);assert.equal(e('D1'),4);assert.equal(e('D2'),4);assert.equal(e('D3'),4/3);assert.equal(e('D4'),4/3);
 // The single-condition average has a distinct documented Boolean rule.
 sheet.cells.D5={raw:'=AVERAGEIF(A1:A6,1,B1:B6)'};assert.equal(createEvaluator(book,sheet.id)('D5'),3);
 sheet.cells.D6={raw:'=SUMIFS(B:B,A:A,0)'};assert.equal(createEvaluator(book,sheet.id)('D6'),'#N/A');
});
test('conditional Boolean arithmetic is retained in current XLSX caches and recalculation after reopen',async()=>{
 let {book,sheet}=fixture();sheet.cells.D1={raw:'=SUMIFS(B1:B6,A1:A6,1)'};sheet.cells.D2={raw:'=AVERAGEIFS(B:B,A:A,1)'};
 for(let pass=0;pass<2;pass++){
  const output=await writeSpreadsheet(book),excel=new ExcelJS.Workbook();await excel.xlsx.load(output.data);assert.equal(excel.worksheets[0].getCell('D1').result,4);assert.equal(excel.worksheets[0].getCell('D2').result,4/3);
  book=await readSpreadsheet(output.data,'booleans.xlsx');sheet=book.sheets[0];const e=createEvaluator(book,sheet.id);assert.equal(e('D1'),4);assert.equal(e('D2'),4/3);
 }
});
