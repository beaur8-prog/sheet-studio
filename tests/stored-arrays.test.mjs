import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {createWorkbook,createEvaluator,evaluateFormulaArray,validateWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {storedArrayIndex,clearReplacedArrayDefinitions,normalizeArrayCopy} from '../src/stored-arrays.mjs';
import {assertCellUpdates} from '../src/cell-constraints.mjs';
import {pastePlan,scalarRaw} from '../src/cell-paste.mjs';
function array(sheet,key,formula,ref){sheet.cells[key]={raw:formula,xlsxFormula:{raw:formula,formula:formula.slice(1),shareType:'array',ref}};}
function fixture(){const book=createWorkbook(),s=book.sheets[0];s.cells={A1:{raw:'1'},A2:{raw:'2'},A3:{raw:'3'}};return {book,s};}
test('stored arrays recalculate every result from its source, replace stale cache cells and persist fresh follower values',async()=>{
  const {book,s}=fixture();array(s,'B1','=A1:A3*2','B1:B3');s.cells.B2={raw:'999'};s.cells.B3={raw:'999'};assert.deepEqual(['B1','B2','B3'].map(createEvaluator(book,s.id)),[2,4,6]);s.cells.A2.raw='4';assert.deepEqual(['B1','B2','B3'].map(createEvaluator(book,s.id)),[2,8,6]);
  const bytes=(await writeSpreadsheet(book)).data,out=new ExcelJS.Workbook();await out.xlsx.load(bytes);assert.equal(out.worksheets[0].getCell('B1').value.shareType,'array');assert.equal(out.worksheets[0].getCell('B2').value,8);const loaded=await readSpreadsheet(bytes,'arrays.xlsx');assert.equal(createEvaluator(loaded,loaded.activeSheetId)('B2'),8);
});
test('array IF, IFERROR and IFNA select per element without propagating unused branch errors',()=>{
  const {book,s}=fixture();s.cells.B1={raw:'10'};s.cells.B2={raw:'=1/0'};s.cells.C1={raw:'=1/0'};s.cells.C2={raw:'20'};
  assert.deepEqual(evaluateFormulaArray(book,s.id,'IF({TRUE;FALSE},B1:B2,C1:C2)').values,[10,20]);assert.deepEqual(evaluateFormulaArray(book,s.id,'IF({TRUE;TRUE},10,Missing())').values,[10,10]);assert.deepEqual(evaluateFormulaArray(book,s.id,'IFERROR({#N/A,5},{1,2})').values,[1,5]);assert.deepEqual(evaluateFormulaArray(book,s.id,'IFNA({#DIV/0!,#N/A},{1,2})').values,['#DIV/0!',2]);
  array(s,'D1','=IF(A1:A3>1,A1:A3*2,0)','D1:D3');assert.deepEqual(['D1','D2','D3'].map(createEvaluator(book,s.id)),[0,4,6]);
});
test('array scalar and single-axis broadcasts, padding, ROW/COLUMN and missing followers behave as one result block',async()=>{
  const {book,s}=fixture();array(s,'B1','=SUM(A1:A3)','B1:B3');array(s,'C1','=A1:A3*2','C1:C4');array(s,'D1','=COLUMN(A1:C3)','D1:F3');array(s,'G1','=ROW(A5:A7)','G1:G3');const e=createEvaluator(book,s.id);assert.deepEqual(['B1','B2','B3'].map(e),[6,6,6]);assert.deepEqual(['C1','C2','C3','C4'].map(e),[2,4,6,'#N/A']);assert.deepEqual(['D1','E1','F1','D3','E3','F3'].map(e),[1,2,3,1,2,3]);assert.deepEqual(['G1','G2','G3'].map(e),[5,6,7]);s.cells.H1={raw:'=SUM(B:B)'};assert.equal(createEvaluator(book,s.id)('H1'),18);
  const csv=new TextDecoder().decode((await writeSpreadsheet(book,'csv')).data);assert.ok(csv.includes('6,4,1,2,3,6'));const loaded=await readSpreadsheet((await writeSpreadsheet(book)).data,'filled.xlsx');assert.equal(loaded.sheets[0].cells.F3.raw,'3');assert.equal(createEvaluator(loaded,loaded.activeSheetId)('B3'),6);
});
test('complete array replacement cannot leave a dormant definition and partial literal replacement is refused',()=>{
  const {book,s}=fixture();array(s,'B1','=A1:A3*2','B1:B3');assert.throws(()=>assertCellUpdates(book,s,[{key:'B1',cell:{...s.cells.B1,raw:'value'}}],{editAnchor:true}),/배열/);
  const updates=['B1','B2','B3'].map(key=>({key,cell:{...s.cells[key],raw:''}}));assert.doesNotThrow(()=>assertCellUpdates(book,s,updates));clearReplacedArrayDefinitions(s,updates);for(const u of updates)s.cells[u.key]=u.cell;assert.equal(s.cells.B1.xlsxFormula,undefined);s.cells.B1.raw='=SUM(A1:A3)';assert.equal(storedArrayIndex(s).regions.length,0);assert.equal(createEvaluator(book,s.id)('B2'),'');
  array(s,'B1','=A1:A3*2','B1:B3');book.nativeSource={};assert.throws(()=>assertCellUpdates(book,s,[{key:'B1',cell:{raw:'=A1:A3'}}],{editAnchor:true}),/XLSM/);
  assert.doesNotThrow(()=>assertCellUpdates(book,s,[{key:'B2',cell:{raw:'',style:{bold:true}}}],{formatOnly:true}));
});
test('full array copy keeps a relocated definition and current values; partial and value copies cannot resurrect the original block',async()=>{
  const {book,s}=fixture();array(s,'B1','=$A$1:$A$3*2','B1:B3');s.cells.B2={raw:'999'};const rows=[0,1,2],cols=[1],e=createEvaluator(book,s.id),values=rows.map(r=>[e('B'+(r+1))]),data=rows.map(r=>[structuredClone(s.cells['B'+(r+1)]||{raw:''})]);normalizeArrayCopy(s,data,values,rows,cols,scalarRaw);assert.equal(data[1][0].raw,'4');const clip={data,values,r:0,c:1,sourceRows:rows,sourceCols:cols},plan=pastePlan(s,clip,{r:0,c:4});assert.equal(plan.updates[0].cell.xlsxFormula.ref,'E1:E3');for(const u of plan.updates)s.cells[u.key]=u.cell;assert.equal(createEvaluator(book,s.id)('E2'),4);validateWorkbook(book);
  const single=[[structuredClone(s.cells.B1)]];normalizeArrayCopy(s,single,[[2]],[0],[1],scalarRaw);assert.equal(single[0][0].raw,'2');assert.equal(single[0][0].xlsxFormula,undefined);assert.throws(()=>pastePlan(s,clip,{r:4,c:0},{transpose:true}),/값으로 전치/);const transposed=pastePlan(s,clip,{r:4,c:0},{transpose:true,mode:'values'});assert.deepEqual(transposed.updates.map(u=>u.cell.raw),['2','4','6']);assert.ok(transposed.updates.every(u=>!u.cell.xlsxFormula));
  const replacement=pastePlan(s,clip,{r:0,c:4});clearReplacedArrayDefinitions(s,replacement.updates);assert.ok(replacement.updates[0].cell.xlsxFormula);
  s.cells.B1.raw='=$A$1:$A$3*3';const editedData=rows.map(r=>[structuredClone(s.cells['B'+(r+1)]||{raw:''})]),editedValues=[[3],[6],[9]];normalizeArrayCopy(s,editedData,editedValues,rows,cols,scalarRaw);const edited=pastePlan(s,{...clip,data:editedData,values:editedValues},{r:0,c:7});for(const u of edited.updates)s.cells[u.key]=u.cell;const loaded=await readSpreadsheet((await writeSpreadsheet(book)).data,'edited-copy.xlsx');assert.equal(createEvaluator(loaded,loaded.activeSheetId)('H3'),9);
});
