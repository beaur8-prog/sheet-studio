import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator,address} from '../src/engine.mjs';
import {clearContentsPlan} from '../src/clear-contents.mjs';
import {assertCellUpdates} from '../src/cell-constraints.mjs';
import {clearReplacedArrayDefinitions} from '../src/stored-arrays.mjs';

test('clearing a million-result spill changes only its anchor and retains blank formatting and notes',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'=SEQUENCE(1000000)'},A3:{raw:'',style:{bold:true},note:'Keep'},B1:{raw:'outside'}};const e=createEvaluator(b,s.id);assert.equal(e('A1000000'),1000000);const before=JSON.stringify(b),updates=clearContentsPlan(s,{r1:0,r2:999999,c1:0,c2:0});assert.equal(updates.length,1);assert.equal(JSON.stringify(b),before);assert.doesNotThrow(()=>assertCellUpdates(b,s,updates));for(const u of updates)s.cells[u.key]=u.cell;assert.equal(createEvaluator(b,s.id)('A999999'),'');assert.equal(s.cells.A3.note,'Keep');assert.equal(s.cells.A3.style.bold,true);assert.equal(s.cells.B1.raw,'outside');
});
test('partial or hidden spill-anchor clearing is refused even when the selection has no stored follower records',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'=SEQUENCE(1000000)'};createEvaluator(b,s.id)('A2');assert.throws(()=>clearContentsPlan(s,{r1:1,r2:999999,c1:0,c2:0}),/배열/);assert.throws(()=>clearContentsPlan(s,{r1:0,r2:999999,c1:0,c2:0},{editable:r=>r!==0}),/배열/);
});
test('legacy CSE replacement remains complete and ordinary sparse clearing respects hidden cells and edit limits',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'={1;2;3}',xlsxFormula:{raw:'={1;2;3}',formula:'{1;2;3}',shareType:'array',ref:'A1:A3'}},B1:{raw:'keep'}};assert.throws(()=>clearContentsPlan(s,{r1:0,r2:1,c1:0,c2:0}),/배열/);const updates=clearContentsPlan(s,{r1:0,r2:2,c1:0,c2:0});assert.equal(updates.length,3);assertCellUpdates(b,s,updates);clearReplacedArrayDefinitions(s,updates);assert.equal(updates[0].cell.xlsxFormula,undefined);
 s.cells={A1:{raw:'1'},A1000000:{raw:'2'},B1:{raw:'outside'}};const sparse=clearContentsPlan(s,{r1:0,r2:999999,c1:0,c2:0},{editable:r=>r!==0});assert.equal(sparse.length,1);assert.equal(sparse[0].key,'A1000000');
 s.cells={};for(let r=0;r<=100000;r++)s.cells[address(r,0)]={raw:'1'};assert.throws(()=>clearContentsPlan(s,{r1:0,r2:100000,c1:0,c2:0}),/100,000/);assert.equal(s.cells.A1.raw,'1');
});
