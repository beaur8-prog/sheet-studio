import test from 'node:test';
import assert from 'node:assert/strict';
import {performance} from 'node:perf_hooks';
import {createWorkbook,createEvaluator,transformFormulaReferences,address} from '../src/engine.mjs';
import {structuralWorkbookPlan} from '../src/workbook-structure.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {tableFixture} from './tables-fixture.mjs';
import {attachNativeSource,buildNativePlan} from '../src/native-plan.mjs';
import {chartData} from '../src/chart-model.mjs';
import {xml,all,attr} from '../src/chart-xml.mjs';
const change=(book,options)=>structuralWorkbookPlan(book,book.sheets[0].id,options);
test('contiguous deletions contract ranges at either boundary and invalidate only fully deleted references',()=>{
  const move=(formula,axis,index,delta)=>transformFormulaReferences(formula,{axis,index,delta,targetSheet:'S',ownerSheet:'S'});
  for(const [formula,want] of [['=SUM(A1:A10)','=SUM(A1:A7)'],['=SUM(A3:A8)','=SUM(A2:A5)'],['=SUM(A2:A4)','=SUM(#REF!)'],['=$A$3+A5','=#REF!+A2'],['=A10:A1','=A7:A1']])assert.equal(move(formula,'r',1,-3),want);
  assert.equal(move('=SUM(A:E)+ROWS(2:10)+$B$2+"A2"','c',1,-3),'=SUM(A:B)+ROWS(2:10)+#REF!+"A2"');
  assert.equal(move("=SUM('S'!$2:$10)+Other!A3",'r',1,-3),"=SUM('S'!$2:$7)+Other!A3");
  assert.equal(move('=SUM(A1:A1048576)+A20','r',10,10000),'=SUM(A1:A1048576)+A10020');
});
test('bulk row changes keep calculated templates, totals, cross-sheet references and original books intact',async()=>{
  const book=await readSpreadsheet(await tableFixture(),'bulk.xlsx'),s=book.sheets[0],t=s.tables[0];t.columns[2].formula='A2*B2';s.cells.C2.raw='=A2*B2';s.cells.C3.raw='=A3*B3';
  const before=JSON.stringify(book);let result=change(book,{axis:'r',index:1,count:3,delta:1}),target=result.sheets[0];assert.equal(JSON.stringify(book),before);assert.equal(target.tables[0].range,'A1:C7');assert.equal(target.tables[0].columns[2].formula,'A2*B2');
  for(let r=2;r<=6;r++)assert.equal(target.cells['C'+r].raw,'=A'+r+'*B'+r);assert.equal(createEvaluator(result,target.id)('C7'),26);
  result=change(result,{axis:'r',index:1,count:3,delta:-1});target=result.sheets[0];assert.equal(target.tables[0].range,'A1:C4');assert.equal(target.cells.C2.raw,'=A2*B2');assert.equal(createEvaluator(result,target.id)('C4'),26);
  const reloaded=await readSpreadsheet((await writeSpreadsheet(result)).data,'bulk.xlsx');assert.equal(createEvaluator(reloaded,reloaded.sheets[0].id)('C4'),26);
});
test('bulk column insertion and partial leading deletion retain unique headers, identities and deleted structured references',async()=>{
  let book=await readSpreadsheet(await tableFixture(),'bulkcols.xlsx');book=change(book,{axis:'c',index:1,count:3,delta:1});let s=book.sheets[0];assert.equal(s.tables[0].range,'A1:F4');assert.deepEqual(s.tables[0].columns.map(c=>c.name),['Qty','열2','열3','열4','Price','Amount']);assert.equal(s.cells.B1.raw,'열2');assert.equal(s.cells.D1.raw,'열4');assert.equal(createEvaluator(book,s.id)('H1'),26);
  book=change(book,{axis:'c',index:0,count:2,delta:-1});s=book.sheets[0];assert.equal(s.tables[0].range,'A1:D4');assert.deepEqual(s.tables[0].columns.map(c=>c.name),['열3','열4','Price','Amount']);assert.equal(new Set(s.tables[0].columns.map(c=>c.id)).size,4);assert.equal(createEvaluator(book,s.id)('F1'),'#REF!');
  book=await readSpreadsheet((await writeSpreadsheet(book)).data,'columns.xlsx');assert.equal(book.sheets[0].tables[0].filter.range.c2,3);
});
test('bulk changes rebase validation and conditional formulas, hidden sizes, protection, filters, freeze and chart OFFSET references',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.name='S';s.cells={A1:{raw:'header'},A2:{raw:'1'},A3:{raw:'2'},A4:{raw:'3'},A5:{raw:'4'},A6:{raw:'5'}};s.freeze={rows:4,cols:0};s.hiddenSizes={rows:{1:25,2:30,5:40},cols:{}};s.xlsxProtectionDefaults={rows:{1:{locked:false},5:{locked:false}},columns:{}};s.rowHeights={5:0};s.autoFilter={range:{r1:0,r2:5,c1:0,c2:0},columns:{},hiddenRows:[2,5]};
  s.conditionalRules=[{id:'rule',range:'A2:A6',type:'expression',formulae:['A2>0'],style:{color:'#ff0000'}}];s.xlsxFeatures={dataValidations:'<dataValidations count="1"><dataValidation type="custom" sqref="A2:A6"><formula1>A2&gt;0</formula1></dataValidation></dataValidations>'};s.charts=[{id:'c',type:'bar',range:'A2:A6',series:[{name:'offset',valRef:{formula:"OFFSET('S'!$A$5,0,0,2,1)",sheetId:s.id,ownerSheetId:s.id,range:'A5:A6'}}]}];
  const out=change(b,{axis:'r',index:1,count:3,delta:-1}),target=out.sheets[0];assert.deepEqual(target.freeze,{rows:1,cols:0});assert.deepEqual(target.hiddenSizes.rows,{'2':40});assert.deepEqual(target.xlsxProtectionDefaults.rows,{'2':{locked:false}});assert.deepEqual(target.autoFilter.hiddenRows,[2]);assert.equal(target.conditionalRules[0].range,'A2:A3');assert.equal(target.conditionalRules[0].formulae[0],'A2>0');
  const validation=xml(target.xlsxFeatures.dataValidations);assert.equal(attr(all(validation,'dataValidation')[0],'sqref'),'A2:A3');assert.equal(all(validation,'formula1')[0].textContent,'A2>0');assert.match(target.charts[0].series[0].valRef.formula,/\$A\$2/);assert.deepEqual(chartData(target.charts[0],target,out)[0].values,[4]);assert.deepEqual(chartData({...target.charts[0],plotVisibleOnly:false},target,out)[0].values,[4,5]);
});
test('sheet-edge inserts reject loss before mutation, and native baselines mirror bulk operations once',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells={A1048576:{raw:'last'}};const before=JSON.stringify(b);assert.throws(()=>change(b,{axis:'r',index:0,count:2,delta:1}),/사라질/);assert.equal(JSON.stringify(b),before);
  s.cells={A1:{raw:'1'},A5:{raw:'=A1+1'}};attachNativeSource(b,new Uint8Array([80,75,3,4]),'sample.xlsm');const out=change(b,{axis:'r',index:1,count:3,delta:1}),plan=buildNativePlan(out);assert.equal(out.sheets[0].cells.A8.raw,'=A1+1');assert.equal(out.nativeSource.baseline[0].cells.A8.raw,'=A1+1');assert.equal(plan.operations.length,3);assert.ok(plan.operations.every(o=>o.delta===1));assert.equal(plan.sheets[0].cells.length,0);
});
test('10,000-row edits visit a large sparse document in bounded time without repeated whole-book transforms',()=>{
  const b=createWorkbook(),s=b.sheets[0];for(let r=0;r<100000;r++)s.cells[address(r,0)]={raw:'1'};s.cells.B1={raw:'=SUM(A1:A100000)'};
  const start=performance.now(),out=change(b,{axis:'r',index:0,count:10000,delta:1}),elapsed=performance.now()-start;assert.equal(Object.keys(out.sheets[0].cells).length,100001);assert.equal(out.sheets[0].cells.B10001.raw,'=SUM(A10001:A110000)');assert.equal(createEvaluator(out,out.activeSheetId)('B10001'),100000);assert.ok(elapsed<8000,'bulk structure took '+elapsed+'ms');assert.equal(s.cells.B1.raw,'=SUM(A1:A100000)');
});
