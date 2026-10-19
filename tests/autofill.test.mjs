import test from 'node:test';
import assert from 'node:assert/strict';
import {autoFillPlan} from '../src/autofill.mjs';
const cell=raw=>({raw:String(raw)}),range=(r1,c1,r2,c2)=>({r1,c1,r2,c2});
const plan=(sourceCells,sourceRange,targetRange,extra={})=>autoFillPlan({sourceCells,sourceRange,targetRange,...extra});
const values=updates=>updates.map(u=>u.cell.raw);

test('two numeric seeds extend vertically, horizontally and in reverse directions',()=>{
 assert.deepEqual(values(plan({A1:cell(1),A2:cell(2)},range(0,0,1,0),range(0,0,5,0))),['3','4','5','6']);
 assert.deepEqual(values(plan({A1:cell(1),B1:cell(2)},range(0,0,0,1),range(0,0,0,5))),['3','4','5','6']);
 assert.deepEqual(values(plan({A5:cell(5),A6:cell(10)},range(4,0,5,0),range(0,0,5,0))),['-15','-10','-5','0']);
 assert.deepEqual(values(plan({E2:cell(2),F2:cell(4)},range(1,4,1,5),range(1,0,1,5))),['-6','-4','-2','0']);
});
test('arbitrary steps, descending values, decimals and percentages retain numeric precision',()=>{
 for(const [a,b,expected] of [[2,4,['6','8','10','12']],[100,95,['90','85','80','75']],[-2,-1,['0','1','2','3']],[.1,.2,['0.3','0.4','0.5','0.6']],[4,4,['4','4','4','4']],['10%','20%',['30%','40%','50%','60%']]])assert.deepEqual(values(plan({A1:cell(a),A2:cell(b)},range(0,0,1,0),range(0,0,5,0))),expected);
});
test('three or more numeric seeds use a linear best-fit trend without changing seeds',()=>{
 const source={A1:cell(3),A2:cell(5),A3:cell(8)},before=structuredClone(source),out=plan(source,range(0,0,2,0),range(0,0,4,0));
 assert.ok(Math.abs(Number(out[0].cell.raw)-31/3)<1e-12);assert.ok(Math.abs(Number(out[1].cell.raw)-77/6)<1e-12);assert.deepEqual(source,before);
 assert.deepEqual(values(plan({A1:cell(1),A2:cell(2),A3:cell(3)},range(0,0,2,0),range(0,0,5,0))),['4','5','6']);
});
test('rectangular seeds calculate an independent series for every row or column',()=>{
 const source={A1:cell(1),A2:cell(2),B1:cell(10),B2:cell(20)};
 assert.deepEqual(values(plan(source,range(0,0,1,1),range(0,0,3,1))),['3','30','4','40']);
 assert.deepEqual(values(plan(source,range(0,0,1,1),range(0,0,1,3))),['19','28','38','56']);
});
test('one seed, text, explicit numeric text and mixed types continue to copy',()=>{
 assert.deepEqual(values(plan({A1:cell(7)},range(0,0,0,0),range(0,0,3,0))),['7','7','7']);
 for(const [a,b] of [['월','화'],["'1","'2"],['1','text'],['1','']])assert.deepEqual(values(plan({A1:cell(a),A2:cell(b)},range(0,0,1,0),range(0,0,5,0))),[a,b,a,b]);
 assert.deepEqual(values(plan({A1:{raw:'1',style:{numberFormat:'@'}},A2:{raw:'2',style:{numberFormat:'@'}}},range(0,0,1,0),range(0,0,3,0))),['1','2']);
});
test('formula copies adjust only relative references and discard stale caches',()=>{
 const source={A1:{raw:'=B1+$C$1',cachedValue:42},A2:cell('=B2+$C$1')};
 const out=plan(source,range(0,0,1,0),range(0,0,5,0));assert.deepEqual(values(out),['=B3+$C$1','=B4+$C$1','=B5+$C$1','=B6+$C$1']);assert.equal(out[0].cell.cachedValue,undefined);assert.equal(source.A1.cachedValue,42);
});
test('source styles alternate independently of the generated numeric values',()=>{
 const source={A1:{raw:'1',style:{bold:true,borders:{top:{style:'thin',color:'#ff0000'}}}},A2:{raw:'2',style:{italic:true}}},out=plan(source,range(0,0,1,0),range(0,0,5,0));
 assert.equal(out[0].cell.raw,'3');assert.deepEqual(out[0].cell.style,source.A1.style);assert.deepEqual(out[1].cell.style,source.A2.style);out[0].cell.style.borders.top.color='#000000';assert.equal(source.A1.style.borders.top.color,'#ff0000');
});
test('hidden rows and columns neither change nor consume visible series positions',()=>{
 const source={A1:cell(1),A2:cell(2),A3:cell('keep')};
 assert.deepEqual(plan(source,range(0,0,1,0),range(0,0,5,0),{hiddenRow:r=>r===2}).map(u=>[u.key,u.cell.raw]),[['A4','3'],['A5','4'],['A6','5']]);
 assert.deepEqual(values(plan({A1:cell(1),A2:cell(999),A3:cell(2)},range(0,0,2,0),range(0,0,5,0),{hiddenRow:r=>r===1})),['3','4','5']);
 assert.deepEqual(plan({A1:cell(1),B1:cell(2)},range(0,0,0,1),range(0,0,0,5),{hiddenCol:c=>c===2}).map(u=>u.key),['D1','E1','F1']);
 assert.equal(plan(source,range(0,0,1,0),range(0,0,5,0),{hiddenCol:()=>true}).length,0);
});
test('protected destinations and original cells are not overwritten; a no-op has no updates',()=>{
 const source={A1:cell(1),A2:cell(2)},out=plan(source,range(0,0,1,0),range(0,0,5,0),{isEditable:r=>r!==3});
 assert.deepEqual(out.map(u=>u.key),['A3','A5','A6']);assert.deepEqual(values(out),['3','5','6']);assert.deepEqual(plan(source,range(0,0,1,0),range(0,0,1,0)),[]);
});
test('diagonal expansion keeps the existing two-dimensional copy pattern',()=>{
 const out=plan({A1:cell(1),A2:cell(2)},range(0,0,1,0),range(0,0,3,1));assert.deepEqual(out.map(u=>[u.key,u.cell.raw]),[['B1','1'],['B2','2'],['A3','1'],['B3','1'],['A4','2'],['B4','2']]);
});
test('range and arithmetic overflow errors are atomic',()=>{
 const source={A1:cell(1),A2:cell(2)},before=structuredClone(source);
 assert.throws(()=>plan(source,range(0,0,1,0),range(0,0,100000,0)),/100,000/);assert.throws(()=>plan(source,range(0,0,1,0),range(0,0,1048576,0)),/범위/);assert.throws(()=>plan(source,range(0,0,1,0),range(2,0,5,0)),/원본/);assert.deepEqual(source,before);
 assert.throws(()=>plan({A1:cell('1e308'),A2:cell('1.7e308')},range(0,0,1,0),range(0,0,3,0)),/계산 범위/);
});
