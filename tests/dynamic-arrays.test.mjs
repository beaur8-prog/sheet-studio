import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator,evaluateFormulaArray,address} from '../src/engine.mjs';
import {calculatedArrayRegions} from '../src/calculated-arrays.mjs';

function make(cells){const b=createWorkbook();b.sheets[0].cells=Object.fromEntries(Object.entries(cells).map(([ref,raw])=>[ref,{raw}]));return b;}

test('spill followers are calculated before the anchor is requested without materializing editable cells',()=>{
  const b=make({A1:'=SEQUENCE(3,2)',D1:'=SUM(A1#)',E1:'=TEXT(A1#,"0.0")'}),s=b.sheets[0],before=JSON.stringify(b),e=createEvaluator(b,s.id);
  assert.equal(e('B3'),6);assert.equal(e('A2'),3);assert.equal(e('D1'),21);assert.equal(e('F3'),'6.0');assert.equal(e('A1'),1);
  assert.equal(e.isComputed('B3'),true);assert.equal(e.isComputed('C1'),false);assert.equal(e.spillAt('B3').ref,'A1');assert.equal(e.spills().length,2);assert.equal(calculatedArrayRegions(s).length,2);assert.equal(JSON.stringify(b),before);assert.equal(Object.keys(s.cells).length,3);
});
test('new calculation sessions resize, shrink and delete the old spill after source edits',()=>{
  const b=make({A1:'=SEQUENCE(D1)',D1:'4'}),s=b.sheets[0];let e=createEvaluator(b,s.id);assert.equal(e('A4'),4);
  s.cells.D1.raw='2';e=createEvaluator(b,s.id);assert.equal(e('A2'),2);assert.equal(e('A3'),'');assert.equal(e.spills()[0].r2,1);
  s.cells.A1.raw='plain';e=createEvaluator(b,s.id);assert.equal(e('A1'),'plain');assert.equal(e('A2'),'');assert.deepEqual(e.spills(),[]);
});
test('occupied values, empty-string formulas and explicit empty text block spills while style-only cells do not',()=>{
  for(const raw of ['x','0','FALSE','=""',"'"]){const b=make({A1:'=SEQUENCE(3)',A2:raw}),e=createEvaluator(b,'sheet-1');assert.equal(e('A1'),'#SPILL!',raw);assert.equal(e.spills().length,0);assert.equal(e.spills({includeBlocked:true})[0].obstruction,'A2');}
  const b=make({A1:'=SEQUENCE(3)',A2:''});b.sheets[0].cells.A2.style={background:'#00ff00'};const e=createEvaluator(b,'sheet-1');assert.equal(e('A3'),3);assert.equal(b.sheets[0].cells.A2.raw,'');
});
test('merges, Excel tables, fixed arrays, collisions and sheet edges reject the complete spill',()=>{
  let b=make({A1:'=SEQUENCE(3)'});b.sheets[0].merges=['A2:B2'];assert.equal(createEvaluator(b,'sheet-1')('A1'),'#SPILL!');
  b=make({A1:'=SEQUENCE(3,2)'});b.sheets[0].tables=[{range:'B2:C5'}];assert.equal(createEvaluator(b,'sheet-1')('A1'),'#SPILL!');
  b=make({A2:'=SEQUENCE(1,2)',B1:'={1;2;3}'});b.sheets[0].cells.B1.xlsxFormula={shareType:'array',raw:'={1;2;3}',formula:'{1;2;3}',ref:'B1:B3'};assert.equal(createEvaluator(b,'sheet-1')('A2'),'#SPILL!');
  for(const cells of [{A2:'=SEQUENCE(1,2)',B1:'=SEQUENCE(2)'},{B1:'=SEQUENCE(2)',A2:'=SEQUENCE(1,2)'}]){b=make(cells);const e=createEvaluator(b,'sheet-1');assert.equal(e('A2'),'#SPILL!');assert.equal(e('B2'),2);}
  b=make({A1048576:'=SEQUENCE(2)',XFD1:'=SEQUENCE(1,2)'});const e=createEvaluator(b,'sheet-1');assert.equal(e('A1048576'),'#SPILL!');assert.equal(e('XFD1'),'#SPILL!');
});
test('spilled references, names, dependent spill inputs, element errors and implicit intersection compose',()=>{
  const b=make({A1:'=SEQUENCE(B2,2)',B1:'=SEQUENCE(2)',D1:'=SORT(UNIQUE({3;1;2;1}))',F1:'={1;#N/A;3}',G1:'=FILTER(F1#,{TRUE;FALSE;TRUE})',H2:'=@D1:D3',I1:'=@SEQUENCE(3)',J1:'=_xlfn.ANCHORARRAY(D1)',K1:'=_xlfn.SINGLE(D1:D3)'}),s=b.sheets[0];
  // A1's eventual two-column result would overlap B1, so the existing formula
  // blocks the whole spill; B1 still produces a valid value used as input.
  const e=createEvaluator(b,s.id);assert.equal(e('A1'),'#SPILL!');assert.equal(e('B2'),2);assert.equal(e('D3'),3);assert.equal(e('F2'),'#N/A');assert.equal(e('G2'),3);assert.equal(e('H2'),2);assert.equal(e('I1'),1);assert.equal(e('J3'),3);assert.equal(e('K1'),1);
  b.chartNames=[{name:'Numbers',formula:'Sheet1!$D$1#'}];s.cells.M1={raw:'=Numbers+10'};const named=createEvaluator(b,s.id);assert.equal(named('M3'),13);
  assert.deepEqual(evaluateFormulaArray(b,s.id,'=FILTER(D1#,LEN(D1#)>0)'),{rows:3,cols:1,values:[1,2,3]});
});
test('whole-column aggregation and shared cross-sheet views include derived values exactly once',()=>{
  const b=make({A1:'=SEQUENCE(4)',C1:'=SUM(A:A)',C2:'=COUNT(A:A)',C3:'=AVERAGE(A:A)'}),s=b.sheets[0];b.sheets.push({id:'other',name:'Other',cells:{A1:{raw:'=Sheet1!A1#*2'},C1:{raw:'=SUM(Sheet1!A1#)'}},charts:[],merges:[],columnWidths:{},rowHeights:{}});
  const e=createEvaluator(b,s.id),other=e.forSheet('other');assert.equal(e('C1'),10);assert.equal(e('C2'),4);assert.equal(e('C3'),2.5);assert.equal(other('A4'),8);assert.equal(other('C1'),10);assert.equal(other.spills()[0].r2,3);
});
test('direct and indirect references into a formula own output return a circular-reference error',()=>{
  for(const cells of [{A1:'=A2:A3'},{A1:'=B1:B2',B1:'=A2'}]){const b=make(cells);assert.equal(createEvaluator(b,'sheet-1')('A1'),'#CYCLE!');}
});
test('a large spill stays sparse in the document and uses one shared calculation session',()=>{
  const b=make({A1:'=SEQUENCE(100000)',C1:'=SUM(A1#)'}),before=JSON.stringify(b),e=createEvaluator(b,'sheet-1');assert.equal(e('A100000'),100000);assert.equal(e('C1'),5000050000);assert.equal(e.spills()[0].r2,99999);assert.equal(JSON.stringify(b),before);assert.equal(Object.keys(b.sheets[0].cells).length,2);
  // Ordinary one-cell ROW/INDEX formulas do not become 10,000 spill anchors.
  const ordinary=make({});for(let r=0;r<12000;r++)ordinary.sheets[0].cells[address(r,0)]={raw:'=ROW()'};const n=createEvaluator(ordinary,'sheet-1');assert.equal(n('A12000'),12000);assert.equal(n.spills().length,0);
});

test('dynamic followers participate in constraints, copy, current region and data navigation',async()=>{
 const {assertCellUpdates,cellEditError}=await import('../src/cell-constraints.mjs'),{normalizeArrayCopy}=await import('../src/stored-arrays.mjs'),{scalarRaw}=await import('../src/cell-paste.mjs'),{currentRegion}=await import('../src/data-region.mjs'),{dataEdge}=await import('../src/cell-navigation.mjs');
 const b=make({A1:'=SEQUENCE(3,2)'}),s=b.sheets[0],e=createEvaluator(b,s.id);assert.equal(e('B3'),6);
 assert.match(cellEditError(b,s,'B2'),/배열/);assert.equal(cellEditError(b,s,'A1'),null);assert.throws(()=>assertCellUpdates(b,s,[{key:'B2',cell:{raw:'x'}}]),/배열/);assert.doesNotThrow(()=>assertCellUpdates(b,s,[{key:'A1',cell:{raw:'x'}}],{editAnchor:true}));
 assert.deepEqual(currentRegion(s,{r:1,c:1}),{r1:0,r2:2,c1:0,c2:1});assert.deepEqual(dataEdge(s,{r:0,c:1},'r',1),{r:2,c:1});
 const full=[[{raw:s.cells.A1.raw},{raw:''}],[{raw:''},{raw:''}],[{raw:''},{raw:''}]];normalizeArrayCopy(s,full,[[1,2],[3,4],[5,6]],[0,1,2],[0,1],scalarRaw);assert.equal(full[0][0].raw,s.cells.A1.raw);assert.equal(full[1][0].raw,'');
 const single=[[{raw:''}]];normalizeArrayCopy(s,single,[[4]],[1],[1],scalarRaw);assert.equal(single[0][0].raw,'4');
});

test('invalidated spill anchors retain REF errors through copy, transpose and structural rewrites',async()=>{
 const {shiftFormula,transformFormulaReferences,transposeFormula,rewriteSheetReferences}=await import('../src/engine.mjs');
 assert.equal(shiftFormula('=SUM(A1#)',-1,0),'=SUM(#REF!)');assert.equal(shiftFormula('=FILTER(A1#,A1#>2)',0,-1),'=FILTER(#REF!,#REF!>2)');assert.equal(shiftFormula('="A1#"',-1,-1),'="A1#"');
 assert.equal(transformFormulaReferences('=SUM(A1#)',{axis:'r',index:0,delta:-1,targetSheet:'Sheet1',ownerSheet:'Sheet1'}),'=SUM(#REF!)');
 assert.equal(transposeFormula('=SUM(A1#)',{r:1,c:1},{r:0,c:0}),'=SUM(#REF!)');assert.equal(rewriteSheetReferences("=SUM('Removed'!A1#)",'Removed',null),'=SUM(#REF!)');
});
