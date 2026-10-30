import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {createWorkbook,createEvaluator} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {pastePlan,scalarRaw} from '../src/cell-paste.mjs';
import {normalizeArrayCopy} from '../src/stored-arrays.mjs';
import {chartData,refValues} from '../src/chart-model.mjs';
const fixture=()=>{const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'="#N/A"'},A2:{raw:'=NA()'},A3:{raw:'="#NAME?"'},A4:{raw:'=1/0'},A5:{raw:"'#N/A"},A6:{raw:'#N/A'}};return {b,s};};
test('detailed results distinguish literal error-looking text from actual formula errors through repeated XLSX saves',async()=>{
  let {b,s}=fixture(),bytes;
  for(let i=0;i<3;i++){
    const e=createEvaluator(b,s.id);assert.deepEqual(e.result('A1'),{value:'#N/A',error:null});assert.deepEqual(e.result('A2'),{value:'#N/A',error:'#N/A'});assert.deepEqual(e.result('A3'),{value:'#NAME?',error:null});assert.equal(e.result('A4').error,'#DIV/0!');assert.equal(e.result('A5').error,null);assert.equal(e.result('A6').error,'#N/A');
    bytes=(await writeSpreadsheet(b)).data;const out=new ExcelJS.Workbook();await out.xlsx.load(bytes);const sheet=out.worksheets[0];assert.equal(sheet.getCell('A1').result,'#N/A');assert.deepEqual(sheet.getCell('A2').result,{error:'#N/A'});assert.equal(sheet.getCell('A3').result,'#NAME?');assert.equal(sheet.getCell('A5').value,'#N/A');assert.deepEqual(sheet.getCell('A6').value,{error:'#N/A'});b=await readSpreadsheet(bytes,'types.xlsx');s=b.sheets[0];
  }
});
test('unsupported formula caches retain whether an error-looking cached string is actual text or an error',async()=>{
  const source=new ExcelJS.Workbook(),s=source.addWorksheet('Cached');s.getCell('A1').value={formula:'_xlfn.FUTUREFOO()',result:'#N/A'};s.getCell('A2').value={formula:'_xlfn.FUTUREFOO()',result:{error:'#N/A'}};let bytes=await source.xlsx.writeBuffer();
  for(let i=0;i<3;i++){const book=await readSpreadsheet(bytes,'cached.xlsx');assert.equal(book.sheets[0].cells.A1.cachedError,false);assert.equal(book.sheets[0].cells.A2.cachedError,true);bytes=(await writeSpreadsheet(book)).data;const out=new ExcelJS.Workbook();await out.xlsx.load(bytes);assert.equal(out.worksheets[0].getCell('A1').result,'#N/A');assert.deepEqual(out.worksheets[0].getCell('A2').result,{error:'#N/A'});}
});
test('value paste, array result copy and CSV preserve text identity independently of visible error spelling',async()=>{
  const {b,s}=fixture();assert.equal(scalarRaw('#N/A'),"'#N/A");assert.equal(scalarRaw('#N/A',{error:'#N/A'}),'#N/A');
  const clip={r:0,c:0,data:[[{raw:'="#N/A"'}],[{raw:'=NA()'}]],values:[['#N/A'],['#N/A']],errors:[[null],['#N/A']]},plan=pastePlan(s,clip,{r:0,c:2},{mode:'values'});assert.equal(plan.updates[0].cell.raw,"'#N/A");assert.equal(plan.updates[1].cell.raw,'#N/A');
  s.cells.D1={raw:'={"#N/A";#N/A}',xlsxFormula:{raw:'={"#N/A";#N/A}',formula:'{"#N/A";#N/A}',shareType:'array',ref:'D1:D2'}};const e=createEvaluator(b,s.id),data=[[{raw:'999'}]];normalizeArrayCopy(s,data,[[e('D2')]],[1],[3],(v)=>scalarRaw(v,{error:e.result('D2').error}));assert.equal(data[0][0].raw,'#N/A');
  const csv=(await writeSpreadsheet(b,'csv')).data,loaded=await readSpreadsheet(csv,'typed.csv'),re=createEvaluator(loaded,loaded.activeSheetId);assert.equal(re.result('A1').error,null);assert.equal(re.result('A2').error,'#N/A');assert.equal(re.result('A3').error,null);assert.equal(re.result('A4').error,'#DIV/0!');
});
test('chart data can read array results omitted from cell XML and never replaces literal #NAME text with numeric cache',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'0'},A2:{raw:'4'},A3:{raw:'6'},B1:{raw:'=A1:A3*2',xlsxFormula:{raw:'=A1:A3*2',formula:'A1:A3*2',shareType:'array',ref:'B1:B3'}},C1:{raw:'="#NAME?"',cachedValue:999,cachedFormula:'="#NAME?"'}};
  assert.deepEqual(refValues({sheetId:s.id,range:'B1:B3'},b),[0,8,12]);assert.deepEqual(chartData({type:'bar',range:'B1:B3',firstRowAsHeader:false},s,b)[0].values,[0,8,12]);assert.deepEqual(refValues({sheetId:s.id,range:'C1:C1'},b),['#NAME?']);
});
test('modern Excel error constants and same-spelling strings retain distinct types in literals, cached results, value paste and XLSX',async()=>{
 const {EXCEL_ERROR_VALUES}=await import('../src/cell-values.mjs');const source=new ExcelJS.Workbook(),s=source.addWorksheet('Modern');
 EXCEL_ERROR_VALUES.forEach((error,i)=>{s.getCell('A'+(i+1)).value={error};s.getCell('B'+(i+1)).value=error;s.getCell('C'+(i+1)).value={formula:'_xlfn.FUTUREERROR()',result:{error}};s.getCell('D'+(i+1)).value={formula:'_xlfn.FUTURETEXT()',result:error};});let bytes=await source.xlsx.writeBuffer();
 for(let round=0;round<3;round++){
  const b=await readSpreadsheet(bytes,'modern.xlsx'),sheet=b.sheets[0],e=createEvaluator(b,sheet.id);EXCEL_ERROR_VALUES.forEach((error,i)=>{assert.equal(e.result('A'+(i+1)).error,error);assert.equal(e.result('B'+(i+1)).error,null);sheet.cells['E'+(i+1)]={raw:'=IFERROR(A'+(i+1)+',"caught")'};sheet.cells['F'+(i+1)]={raw:'=IFERROR(B'+(i+1)+',"caught")'};sheet.cells['G'+(i+1)]={raw:'=IFERROR('+error+',"caught")'};});const next=createEvaluator(b,sheet.id);EXCEL_ERROR_VALUES.forEach((error,i)=>{assert.equal(next('E'+(i+1)),'caught');assert.equal(next('F'+(i+1)),error);assert.equal(next('G'+(i+1)),'caught');});
  const clip={r:0,c:0,data:[[{raw:'#CALC!'}]],values:[['#CALC!']],errors:[['#CALC!']]};assert.equal(pastePlan(sheet,clip,{r:0,c:8},{mode:'values'}).updates[0].cell.raw,'#CALC!');bytes=(await writeSpreadsheet(b)).data;const check=new ExcelJS.Workbook();await check.xlsx.load(bytes);EXCEL_ERROR_VALUES.forEach((error,i)=>{const row=i+1;assert.deepEqual(check.worksheets[0].getCell('A'+row).value,{error});assert.equal(check.worksheets[0].getCell('B'+row).value,error);assert.deepEqual(check.worksheets[0].getCell('C'+row).result,{error});assert.equal(check.worksheets[0].getCell('D'+row).result,error);});
 }
});
