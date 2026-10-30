import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator} from '../src/engine.mjs';
import {formulaReferences,referenceSlot,cycleReference,rangeReference,autoSumPlan,completeFormula} from '../src/formula-tools.mjs';

test('formula references recognize incomplete formulas, quoted sheets, absolute ranges and skip strings/functions/names',()=>{
  const f='=SUM($A$1:B3)+LOG10(C4)+"A8"+\'매출 \'\'원본\'!D2:D6+Named_A1';
  assert.deepEqual(formulaReferences(f).map(r=>r.text),['$A$1:B3','C4',"'매출 ''원본'!D2:D6"]);
  assert.equal(formulaReferences(f)[2].sheet,"매출 '원본");
  assert.deepEqual(formulaReferences('=SUM(B7:B2').map(r=>[r.r1,r.r2]),[[1,6]]);
  assert.equal(formulaReferences('="unfinished A1').length,0);
  assert.equal(formulaReferences('=XFE1+A1048577+Name.B2').length,0);
});

test('reference slots replace only the current reference and F4 cycles mixed addresses without changing sheet names',()=>{
  const f='=SUM(A1:B3)+C5';assert.deepEqual(referenceSlot(f,6,6),{start:5,end:10});
  assert.equal(referenceSlot('="text A1',9),null);assert.deepEqual(referenceSlot('=SUM(',5),{start:5,end:5});
  let value='=SUM(A1:B3)';for(const wanted of ['=SUM($A$1:$B$3)','=SUM(A$1:B$3)','=SUM($A1:$B3)','=SUM(A1:B3)']){value=cycleReference(value,6).formula;assert.equal(value,wanted);}
  assert.equal(cycleReference('=A1!$B$2',6).formula,'=A1!B$2');
  assert.equal(rangeReference({r1:4,c1:2,r2:1,c2:0},'Sheet1',"매출 '원본"),"'매출 ''원본'!A2:C5");
  assert.equal(completeFormula('=SUM(A1:B2'),'=SUM(A1:B2)');assert.equal(completeFormula('=IF(A1>0,SUM(B1:B2'),'=IF(A1>0,SUM(B1:B2))');assert.equal(completeFormula('="('),'="(');
});

function fixture(cells){const b=createWorkbook();b.sheets[0].cells=Object.fromEntries(Object.entries(cells).map(([k,v])=>[k,{raw:String(v)}]));return [b.sheets[0],createEvaluator(b,b.activeSheetId)];}
const at=(r,c,r2=r,c2=c)=>({r,c,r2,c2});
test('AutoSum proposes numeric blocks above then left, excludes headers, and never creates a self-reference',()=>{
  const [s,e]=fixture({B1:'매출',B2:10,B3:20,B4:'=B2+B3',A7:4,B7:8});
  assert.deepEqual(autoSumPlan(s,at(4,1),e),[{r:4,c:1,raw:'=SUM(B2:B4)'}]);
  assert.deepEqual(autoSumPlan(s,at(6,2),e),[{r:6,c:2,raw:'=SUM(A7:B7)'}]);
  assert.equal(autoSumPlan(s,at(0,0),e)[0].raw,'=SUM()');
});
test('AutoSum selected rows/columns reserve blank totals and never overwrite occupied destinations',()=>{
  const [s,e]=fixture({A1:1,A2:2,B1:3,B2:4});
  assert.deepEqual(autoSumPlan(s,at(0,0,1,1),e),[{r:2,c:0,raw:'=SUM(A1:A2)'},{r:2,c:1,raw:'=SUM(B1:B2)'}]);
  assert.deepEqual(autoSumPlan(s,at(0,0,2,1),e),autoSumPlan(s,at(0,0,1,1),e));
  assert.deepEqual(autoSumPlan(s,at(0,0,0,2),e),[{r:0,c:2,raw:'=SUM(A1:B1)'}]);
  const both=autoSumPlan(s,at(0,0,2,2),e);assert.equal(both.length,5);assert.deepEqual(both.at(-1),{r:2,c:2,raw:'=SUM(A3:B3)'});
  s.cells.A3={raw:'keep'};assert.throws(()=>autoSumPlan(s,at(0,0,1,1),e),/데이터/);assert.equal(s.cells.A3.raw,'keep');
  assert.throws(()=>autoSumPlan(s,at(0,0,100000,0),e),/100,000/);
});

test('spill references keep their operator during reference picking and F4 lock cycling',()=>{
 const formula='=SUM(A1#)';assert.deepEqual(formulaReferences(formula).map(r=>({text:r.text,spill:r.spill})),[{text:'A1#',spill:true}]);assert.deepEqual(referenceSlot(formula,8),{start:5,end:8});let f=formula;for(const expected of ['=SUM($A$1#)','=SUM(A$1#)','=SUM($A1#)','=SUM(A1#)']){f=cycleReference(f,6).formula;assert.equal(f,expected);}
});
test('AutoSum includes derived spill followers and reserves the next genuinely empty row',()=>{
 const [s,e]=fixture({A1:'=SEQUENCE(1000)'});assert.deepEqual(autoSumPlan(s,at(1000,0),e),[{r:1000,c:0,raw:'=SUM(A1:A1000)'}]);assert.deepEqual(autoSumPlan(s,at(0,0,999,0),e),[{r:1000,c:0,raw:'=SUM(A1:A1000)'}]);
});
