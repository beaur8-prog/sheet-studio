import test from 'node:test';
import assert from 'node:assert/strict';
import {pastePlan} from '../src/cell-paste.mjs';
import {autoFillPlan} from '../src/autofill.mjs';
const limit=100_000_000;
test('paste and autofill reject excess document text before building repeated long-cell copies',()=>{
 const text='x'.repeat(32768),sheet={cells:{A1:{raw:text}},merges:[]},clip={data:[[{raw:text}]],r:0,c:0};const before=JSON.stringify(sheet);
 assert.throws(()=>pastePlan(sheet,clip,{r:0,c:2},{selection:{r1:0,r2:4999,c1:2,c2:2},textChars:limit-1000}),/100,000,000/);
 assert.throws(()=>autoFillPlan({sourceCells:sheet.cells,sourceRange:{r1:0,r2:0,c1:0,c2:0},targetRange:{r1:0,r2:4999,c1:0,c2:0},textChars:limit-1000}),/100,000,000/);assert.equal(JSON.stringify(sheet),before);
});

test('overwritten cells are released together so valid replacement order never causes a false overflow',()=>{
 const long='x'.repeat(1000),sheet={cells:{C1:{raw:'a'},C2:{raw:long}},merges:[]},clip={data:[[{raw:long}],[{raw:'b'}]],r:0,c:0};assert.deepEqual(pastePlan(sheet,clip,{r:0,c:2},{textChars:limit}).updates.map(u=>u.cell.raw),[long,'b']);assert.equal(sheet.cells.C1.raw,'a');
});
test('cut moves can reuse released text including overlapping source and destination cells',()=>{
 const sheet={cells:{A1:{raw:'abc'},A2:{raw:'def'}},merges:[]},clip={cut:true,data:[[{raw:'abc'}],[{raw:'def'}]],r:0,c:0};const plan=pastePlan(sheet,clip,{r:1,c:0},{textChars:limit-6,removedKeys:new Set(['A1','A2'])});assert.deepEqual(plan.updates.map(u=>u.cell.raw),['abc','def']);
});
test('skipped ragged/blank cells do not create fictitious free text capacity',()=>{
 const sheet={cells:{C2:{raw:'x'.repeat(1000)}},merges:[]},clip={data:[[{raw:'abc'}],[]],r:0,c:0};assert.throws(()=>pastePlan(sheet,clip,{r:0,c:2},{textChars:limit-2,skipBlanks:true}),/100,000,000/);
});
test('values and formula modes use resulting text while formatting-only copies leave the budget unchanged',()=>{
 const sheet={cells:{},merges:[]},clip={data:[[{raw:'=A1',style:{bold:true}}]],values:[['x'.repeat(1000)]],r:0,c:0};assert.throws(()=>pastePlan(sheet,clip,{r:0,c:2},{textChars:limit-999,mode:'values'}),/100,000,000/);assert.equal(pastePlan(sheet,clip,{r:0,c:2},{textChars:limit-3,mode:'formulas'}).updates[0].cell.raw,'=C1');assert.equal(pastePlan(sheet,clip,{r:0,c:2},{textChars:limit,mode:'formats'}).updates[0].cell.style.bold,true);
 assert.throws(()=>pastePlan(sheet,{data:[[{raw:'😀'}]],r:0,c:0},{r:0,c:0},{textChars:limit-1}),/100,000,000/);
});
test('autofill frees replaced cells and accounts only for visible editable destinations',()=>{
 const long='x'.repeat(1000),cells={A1:{raw:long},A2:{raw:long}},sourceRange={r1:0,r2:0,c1:0,c2:0};assert.equal(autoFillPlan({sourceCells:cells,currentCells:cells,sourceRange,targetRange:{r1:0,r2:1,c1:0,c2:0},textChars:limit}).length,1);assert.throws(()=>autoFillPlan({sourceCells:cells,currentCells:cells,sourceRange,targetRange:{r1:0,r2:2,c1:0,c2:0},hiddenRow:r=>r===1,textChars:limit}),/100,000,000/);assert.deepEqual(autoFillPlan({sourceCells:cells,sourceRange,targetRange:{r1:0,r2:2,c1:0,c2:0},isEditable:()=>false,textChars:limit}),[]);
});
