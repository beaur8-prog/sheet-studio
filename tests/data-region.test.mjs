import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,address,createEvaluator} from '../src/engine.mjs';
import {currentRegion,inferSortPlan} from '../src/data-region.mjs';
import {sortRange,inferFilterRange} from '../src/filters.mjs';

function fixture(){const book=createWorkbook(),sheet=book.sheets[0];for(const [ref,raw] of Object.entries({A1:'이름',B1:'매출',A2:'서울',B2:'20',A3:'부산',B3:'10',D10:'상품',E10:'수량',D11:'사과',E11:'2',D12:'배',E12:'1'}))sheet.cells[ref]={raw};return {book,sheet};}
test('single-cell sorting uses its current region and leaves a separate table intact',()=>{
  const {book,sheet}=fixture(),before=structuredClone(sheet.cells),plan=inferSortPlan(book,sheet,{r1:10,r2:10,c1:4,c2:4},4);
  assert.deepEqual(plan.range,{r1:9,r2:11,c1:3,c2:4});assert.equal(plan.headers,true);assert.equal(plan.review,false);
  sortRange(book,sheet,{...plan.range,r1:10},plan.column,false);
  for(const ref of ['A1','B1','A2','B2','A3','B3','D10','E10'])assert.deepEqual(sheet.cells[ref],before[ref]);
  assert.equal(sheet.cells.D11.raw,'배');assert.equal(sheet.cells.E11.raw,'1');assert.equal(sheet.cells.D12.raw,'사과');
  assert.equal(sheet.cells.D2,undefined);
});
test('sort plans retain numeric first rows and review ambiguous headers/partial columns',()=>{
  let {book,sheet}=fixture();sheet.cells={A1:{raw:'3'},A2:{raw:'2'},A3:{raw:'1'}};
  let plan=inferSortPlan(book,sheet,{r1:1,r2:1,c1:0,c2:0},0);assert.equal(plan.headers,false);assert.equal(plan.review,false);sortRange(book,sheet,plan.range,0,false);assert.deepEqual(['A1','A2','A3'].map(ref=>sheet.cells[ref].raw),['1','2','3']);
  ({book,sheet}=fixture());plan=inferSortPlan(book,sheet,{r1:1,r2:2,c1:1,c2:1},1);
  assert.deepEqual(plan.expandedRange,{r1:1,r2:2,c1:0,c2:1});assert.equal(plan.review,true);assert.equal(plan.headers,false);assert.equal(plan.expandedGuess.headers,false);
  sheet.cells={A1:{raw:'pear'},A2:{raw:'apple'},A3:{raw:'plum'}};plan=inferSortPlan(book,sheet,{r1:1,r2:1,c1:0,c2:0},0);assert.equal(plan.ambiguous,true);assert.equal(plan.headers,false);
  assert.throws(()=>inferSortPlan(book,sheet,{r1:50,r2:50,c1:2,c2:2},2),/데이터 안/);
});
test('current region handles gaps, reverse insertion, styled blanks and whole-column selections',()=>{
  const {book,sheet}=fixture();sheet.cells.C5={raw:'',style:{bold:true}};assert.equal(currentRegion(sheet,{r:4,c:2}),null);
  assert.deepEqual(inferFilterRange(sheet,{r1:10,r2:10,c1:4,c2:4}),{r1:9,r2:11,c1:3,c2:4});
  sheet.cells={};for(let r=9999;r>=0;r--)sheet.cells[address(r,0)]={raw:String(r+1)};
  assert.deepEqual(currentRegion(sheet,{r:9999,c:0}),{r1:0,r2:9999,c1:0,c2:0});
  const plan=inferSortPlan(book,sheet,{r1:0,r2:1048575,c1:0,c2:0},0);assert.equal(plan.range.r2,9999);assert.equal(plan.headers,false);
  sheet.cells={A1:{raw:'x'},B2:{raw:'y'},A4:{raw:'isolated'}};assert.deepEqual(currentRegion(sheet,{r:0,c:0}),{r1:0,r2:1,c1:0,c2:1});
});
test('sorting a chosen subset keeps row values together and shifts copied formulas',()=>{
  const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'id'},B1:{raw:'value'},C1:{raw:'double'},A2:{raw:'x'},B2:{raw:'20'},C2:{raw:'=B2*2'},A3:{raw:'y'},B3:{raw:'10'},C3:{raw:'=B3*2'}};
  const plan=inferSortPlan(book,sheet,{r1:1,r2:2,c1:1,c2:1},1);sortRange(book,sheet,plan.expandedRange,1,false);
  assert.equal(sheet.cells.A2.raw,'y');assert.equal(sheet.cells.C2.raw,'=B2*2');assert.equal(createEvaluator(book,sheet.id)('C2'),20);assert.equal(sheet.cells.C1.raw,'double');
});
