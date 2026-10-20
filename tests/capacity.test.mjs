import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook, validateWorkbook, serializeWorkbook, parseWorkbook, evaluateFormula} from '../src/engine.mjs';
import {sortRange} from '../src/filters.mjs';

test('one million recorded cells survive JSON and the document limit is shared across sheets', () => {
  let book=createWorkbook();
  for(let i=1;i<=1_000_000;i++)book.sheets[0].cells['A'+i]={raw:'1'};
  book.sheets[0].cells.A1={raw:'한글',style:{bold:true,color:'#ff0000'}};
  assert.equal(validateWorkbook(book),book);
  let text=serializeWorkbook(book);book=null;
  let restored=parseWorkbook(text);text=null;
  assert.equal(Object.keys(restored.sheets[0].cells).length,1_000_000);
  assert.equal(restored.sheets[0].cells.A1000000.raw,'1');
  assert.deepEqual(restored.sheets[0].cells.A1,{raw:'한글',style:{bold:true,color:'#ff0000'}});
  restored.sheets.push({id:'extra',name:'Extra',cells:{B1:{raw:'',style:{bold:true}}},charts:[]});
  assert.throws(()=>validateWorkbook(restored),/셀 수가 제한/);
  restored=null;
});

test('million-cell aggregates calculate completely and a larger range still fails explicitly', () => {
  const book=createWorkbook();
  for(let i=1;i<=1_000_000;i++)book.sheets[0].cells['A'+i]={raw:'1'};
  assert.equal(evaluateFormula(book,book.sheets[0].id,'=SUM(A1:A1000000)'),1_000_000);
  assert.equal(evaluateFormula(book,book.sheets[0].id,'=SUM(A1:A1000001)'),'#NUM!');
});

test('background sort logic preserves sparse cells, stable ties and relative formula references',()=>{
  const book=createWorkbook(),sheet=book.sheets[0];
  sheet.cells={A1:{raw:'Heading'},A2:{raw:'2'},B2:{raw:'=A2*2',style:{bold:true}},A3:{raw:'1'},A4:{raw:'1'},B4:{raw:'=A4*3'}};
  sortRange(book,sheet,{r1:1,r2:3,c1:0,c2:1},0,false);
  assert.equal(sheet.cells.A1.raw,'Heading');assert.equal(sheet.cells.A2.raw,'1');assert.equal(sheet.cells.B2,undefined);
  assert.equal(sheet.cells.B3.raw,'=A3*3');assert.equal(sheet.cells.B4.raw,'=A4*2');assert.equal(sheet.cells.B4.style.bold,true);
});
