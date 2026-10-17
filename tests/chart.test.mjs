import test from 'node:test';
import assert from 'node:assert/strict';
import {chartRect,resizeChart} from '../src/chart-geometry.mjs';
import {createWorkbook,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';

test('chart resizing preserves opposite anchors and stays within usable dimensions',()=>{
  const r={x:40,y:25,width:500,height:300};
  assert.deepEqual(resizeChart(r,'se',150,90),{x:40,y:25,width:650,height:390});
  assert.deepEqual(resizeChart(r,'nw',100,50),{x:140,y:75,width:400,height:250});
  assert.deepEqual(resizeChart(r,'nw',-500,-500),{x:0,y:0,width:540,height:325});
  assert.deepEqual(resizeChart(r,'se',-1000,-1000),{x:40,y:25,width:280,height:180});
  assert.deepEqual(resizeChart(r,'se',10000,10000),{x:40,y:25,width:2400,height:1800});
  assert.deepEqual(resizeChart(r,'e',20,200),{...r,width:520});
  assert.deepEqual(resizeChart(r,'n',200,20),{...r,y:45,height:280});
});

test('chart geometry survives document round trips and old documents get default dimensions',()=>{
  const book=createWorkbook();
  book.sheets[0].charts=[{id:'chart',type:'bar',range:'A1:B4',title:'매출',x:30,y:20,width:640,height:400}];
  assert.deepEqual(parseWorkbook(serializeWorkbook(book)).sheets[0].charts,book.sheets[0].charts);
  delete book.sheets[0].charts[0].width;delete book.sheets[0].charts[0].height;
  assert.deepEqual(chartRect(parseWorkbook(serializeWorkbook(book)).sheets[0].charts[0]),{x:30,y:20,width:500,height:300});
  book.sheets[0].charts[0].width='900';
  assert.throws(()=>parseWorkbook(JSON.stringify(book)));
});
