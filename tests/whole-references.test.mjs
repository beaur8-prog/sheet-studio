import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,evaluateFormula,createEvaluator,formulaReference,shiftFormula,transposeFormula,transformFormulaReferences,absoluteFormula} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {formulaReferences,cycleReference} from '../src/formula-tools.mjs';
const fixture=cells=>{const book=createWorkbook(),s=book.sheets[0];s.cells=Object.fromEntries(Object.entries(cells).map(([key,raw])=>[key,{raw:String(raw)}]));return {book,s,calc:f=>evaluateFormula(book,s.id,f)};};
test('entire row/column references retain exact dimensions, sheet qualification and copy/insert/delete/transpose locks',()=>{
  const {book,s,calc}=fixture({});assert.equal(calc('ROWS(A:A)'),1048576);assert.equal(calc('COLUMNS(1:1)'),16384);assert.equal(calc("ROWS('Sheet1'!$B:$D)"),1048576);assert.equal(calc("COLUMNS('Sheet1'!$2:$4)"),16384);assert.equal(calc('ISREF(A:A)'),true);assert.equal(calc('SHEET(A:A)'),1);assert.equal(calc('SUM(XFE:XFE)'),'#REF!');
  assert.deepEqual(formulaReference(book,s.id,'B:D').range,{r1:0,r2:1048575,c1:1,c2:3});assert.equal(shiftFormula('=SUM(A:A,$B:C,1:2,$3:4)+A1',2,1),'=SUM(B:B,$B:D,3:4,$3:6)+B3');assert.equal(shiftFormula('=SUM(A:A)+"A:A"',0,-1),'=SUM(#REF!)+"A:A"');assert.equal(transposeFormula('=SUM(A:B,$2:$3)',{r:0,c:0},{r:2,c:4}),'=SUM(3:4,$B:$C)');assert.equal(absoluteFormula('SUM(A:A,1:2)+A1',s.name),"SUM('Sheet1'!$A:$A,'Sheet1'!$1:$2)+'Sheet1'!$A$1");
  const shift=(f,axis,index,delta)=>transformFormulaReferences(f,{axis,index,delta,targetSheet:s.name,ownerSheet:s.name});assert.equal(shift('=SUM(A:C,$2:$4)+A1','c',1,-1),'=SUM(A:B,$2:$4)+A1');assert.equal(shift('=SUM(B:B)','c',1,-1),'=SUM(#REF!)');assert.equal(shift('=SUM(XFD:XFD)','c',0,1),'=SUM(#REF!)');assert.equal(shift('=SUM(A:XFD)','c',0,1),'=SUM(B:XFD)');
  assert.deepEqual(formulaReferences('=SUM(A:A,\'Sheet1\'!$2:$4)+Sales[Some A:A]').map(r=>r.text),['A:A',"'Sheet1'!$2:$4"]);assert.equal(cycleReference('=SUM(A:C)',6).formula,'=SUM($A:$C)');assert.equal(cycleReference('=SUM($A:$C)',7).formula,'=SUM(A:C)');
});
test('aggregations count logical blanks without allocating trailing rows and preserve scalar/reference coercion',()=>{
  const {calc}=fixture({A1:1,A3:2,A5:'TRUE',A7:"'3",A9:'=""',A1048576:5});
  for(const [f,wanted] of [['SUM(A:A)',8],['AVERAGE(A:A)',8/3],['MIN(A:A)',1],['MAX(A:A)',5],['COUNT(A:A)',3],['COUNTA(A:A)',6],['COUNTBLANK(A:A)',1048571],['SUM(A:A,TRUE,"2")',11],['COUNT(A:A,TRUE,"2")',5],['COUNTA(A:A,NA())',7],['SUM(A1:B1048576)',8]])assert.equal(calc(f),wanted,f);
  const errors=fixture({A1:'=1/0',A2:2});assert.equal(errors.calc('COUNT(A:A)'),1);assert.equal(errors.calc('COUNTA(A:A)'),2);assert.equal(errors.calc('SUM(A:A)'),'#DIV/0!');
});
test('whole-column criteria align sparse cells across ranges and include trailing blanks only when matched',()=>{
  const {calc}=fixture({A1:1,A3:2,A1048576:5,B1:'x',B3:'y',B1048576:'x',C2:7,C3:8});
  for(const [f,wanted] of [['COUNTIF(B:B,"x")',2],['SUMIF(B:B,"x",A:A)',6],['AVERAGEIF(B:B,"x",A:A)',3],['COUNTIF(B:B,"")',1048573],['COUNTIFS(B:B,"",A:A,"")',1048573],['SUMIFS(A:A,B:B,"x")',6],['SUMIF(B:B,"",C:C)',7],['AVERAGEIFS(C:C,B:B,"")',7],['COUNTIFS(A:A,">0",B:B,"x")',2],['MAXIFS(A:A,B:B,"x")',5],['MINIFS(C:C,B:B,"x")',0]])assert.equal(calc(f),wanted,f);
  assert.equal(calc('SUMIFS(A:A,B1:B3,"x")'),'#VALUE!');
});
test('INDEX, dynamic reference metadata and lookup functions reach the final row without reading a huge matrix or unused fallback',()=>{
  const {book,s,calc}=fixture({A1:1,A3:2,A1048576:5,B1:'x',B3:'y',B1048576:'last'});
  for(const [f,wanted] of [['INDEX(A:A,1048576)',5],['SUM(INDEX(A:A,0))',8],['SUM(INDIRECT("A:A"))',8],['SUM(OFFSET(A:A,0,0))',8],['ROWS(INDEX(A:A,0))',1048576],['VLOOKUP(5,A:B,2,FALSE)','last'],['XLOOKUP(5,A:A,B:B)','last'],['MATCH(5,A:A,0)',1048576],['XMATCH(5,A:A,0,-1)',1048576],['INDEX((A1:A3,B1:B3),3,1,2)','y'],['INDEX({1,2;3,4},2,2)',4]])assert.equal(calc(f),wanted,f);
  s.cells.D1={raw:'=XLOOKUP(5,A:A,B:B,RAND())'};let calls=0;const random=Math.random;try{Math.random=()=>{calls++;return .5;};assert.equal(createEvaluator(book,s.id)('D1'),'last');assert.equal(calls,0);}finally{Math.random=random;}
  s.cells.A1={raw:'=INDEX(A:A,3)'};assert.equal(createEvaluator(book,s.id)('A1'),2);assert.equal(calc('INDIRECT("Missing")'),'#REF!');
  assert.equal(calc('ROWS(INDEX({1,2;3,4},0,1))'),2);assert.equal(calc('COLUMNS(INDEX({1,2;3,4},1,0))'),2);
});
test('SUBTOTAL excludes filtered/manual rows and nested subtotals before reading cells, with no poisoned range cache',()=>{
  const {book,s}=fixture({A1:1,A2:2,A3:3,A5:'=SUBTOTAL(9,A:A)',B1:'=SUM(A:A)',B2:'=SUBTOTAL(109,A:A)'});s.rowHeights[1]=0;let e=createEvaluator(book,s.id);assert.equal(e('A5'),6);assert.equal(e('B1'),12);assert.equal(e('B2'),4);
  s.autoFilter={range:{r1:0,r2:2,c1:0,c2:0},columns:{},hiddenRows:[2]};e=createEvaluator(book,s.id);assert.equal(e('B2'),1);assert.equal(e('A5'),3);assert.equal(e('B1'),9);
});
test('whole-column reads resume long formula dependencies once and remain current after XLSX save and reopen',async()=>{
  const {book,s}=fixture({A1:1,B1:'=SUM(A:A)',B2:'=COUNTBLANK(A:A)'});for(let r=2;r<=5000;r++)s.cells['A'+r]={raw:'=A'+(r-1)+'+1'};const e=createEvaluator(book,s.id);assert.equal(e('B1'),12502500);assert.equal(e('B2'),1043576);
  const loaded=await readSpreadsheet((await writeSpreadsheet(book)).data,'whole.xlsx');assert.equal(createEvaluator(loaded,loaded.activeSheetId)('B1'),12502500);assert.equal(loaded.sheets[0].cells.B1.raw,'=SUM(A:A)');
});
test('a million populated cells share compact range summaries across many whole-column totals',()=>{
  const {book,s}=fixture({});for(let r=1;r<=999900;r++)s.cells['A'+r]={raw:'1'};for(let r=1;r<=100;r++)s.cells['B'+r]={raw:'=SUM(A:A)'};const e=createEvaluator(book,s.id),start=performance.now();for(let r=1;r<=100;r++)assert.equal(e('B'+r),999900);const elapsed=performance.now()-start;assert.ok(elapsed<10000,`whole-column summaries took ${elapsed}ms`);assert.equal(evaluateFormula(book,s.id,'COUNTBLANK(A:A)'),48676);
});
