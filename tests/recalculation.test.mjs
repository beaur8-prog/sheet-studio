import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {createWorkbook,createEvaluator,evaluateCell,evaluateFormula} from '../src/engine.mjs';

const fixture=cells=>{
  const book=createWorkbook();book.sheets[0].cells=Object.fromEntries(Object.entries(cells).map(([ref,raw])=>[ref,{raw:String(raw)}]));return book;
};

test('date conditions compare serial dates and AVERAGEIFS treats empty criteria cells as zero',()=>{
  const book=fixture({A1:'=DATE(2011,3,11)',A2:'=DATE(2010,1,1)',A3:'=DATE(2010,12,31)',A4:'=DATE(2010,6,30)',B1:10,B2:20,B3:30,B4:40,D1:0,D3:2});
  const calc=formula=>evaluateFormula(book,book.activeSheetId,formula);
  assert.equal(calc('=COUNTIF(A1:A4,">3/1/2010")'),3);
  assert.equal(calc('=SUMIF(A1:A4,">2010-03-01",B1:B4)'),80);
  assert.equal(calc('=COUNTIFS(A1:A4,">=1/1/2010",A1:A4,"<=12/31/2010")'),3);
  assert.equal(calc('=AVERAGEIFS(B1:B3,D1:D3,0)'),15);
  assert.equal(calc('=AVERAGEIF(B1:B3,">0",B1)'),20);
  assert.equal(calc('=AVERAGEIF(B1:B3,">0",B1:B2)'),20);
});

test('manual worker recalculation refreshes volatile dependents and export uses the displayed session',async()=>{
  // Load browser-bundled dependencies before installing the small worker mock;
  // their environment detection should continue to use Node's task scheduler.
  await import('../src/io.mjs');
  const originalSelf=globalThis.self,random=Math.random,pending=new Map();let sequence=0;
  globalThis.self={postMessage:message=>{
    const item=pending.get(message.id);if(!item)return;pending.delete(message.id);
    if(message.error)item.reject(Error(message.error));else item.resolve(message.value);
  }};
  try{
    await import('../src/workbook-worker.mjs?recalculation-test');
    const onmessage=globalThis.self.onmessage;
    const send=(type,payload)=>new Promise((resolve,reject)=>{const id=++sequence;pending.set(id,{resolve,reject});onmessage({data:{id,type,payload}});});
    const book=fixture({A1:'=RAND()',B1:'=A1*2'}),refs=[[book.activeSheetId,'A1'],[book.activeSheetId,'B1']];
    await send('sync',{metadata:{...book,sheets:book.sheets.map(({cells,...sheet})=>sheet)},sheets:book.sheets.map(({id,cells})=>({id,cells})),revision:1});
    Math.random=()=>.25;
    let result=await send('calculate',{refs,revision:1});
    assert.deepEqual(result.values.map(v=>v[2]),[.25,.5]);
    Math.random=()=>.75;
    result=await send('calculate',{refs,revision:1});assert.deepEqual(result.values.map(v=>v[2]),[.25,.5]);
    let output=await send('export',{format:'xlsx',revision:1}),check=new ExcelJS.Workbook();await check.xlsx.load(output.data);
    assert.equal(check.worksheets[0].getCell('A1').result,.25);
    assert.equal(check.worksheets[0].getCell('B1').result,.5);
    result=await send('recalculate',{refs,revision:1});assert.deepEqual(result.values.map(v=>v[2]),[.75,1.5]);
    output=await send('export',{format:'xlsx',revision:1});check=new ExcelJS.Workbook();await check.xlsx.load(output.data);
    assert.equal(check.worksheets[0].getCell('A1').result,.75);
    assert.equal(check.worksheets[0].getCell('B1').result,1.5);
  }finally{globalThis.self=originalSelf;Math.random=random;}
});

test('long dependency chains calculate identically in either viewport order',()=>{
  const book=fixture({A1:1}),sheet=book.sheets[0];
  for(let i=2;i<=10000;i++)sheet.cells['A'+i]={raw:`=A${i-1}+1`};
  const lastFirst=createEvaluator(book,sheet.id);
  assert.equal(lastFirst('A10000'),10000);
  assert.equal(lastFirst('A5000'),5000);
  const firstFirst=createEvaluator(book,sheet.id);
  for(let i=1;i<=10000;i++)assert.equal(firstFirst('A'+i),i);
  assert.equal(evaluateFormula(book,sheet.id,'=IFERROR(A10000,0)+1'),10001);
  assert.equal(evaluateFormula(book,sheet.id,'=ISREF(A10000)'),true);
  assert.equal(evaluateFormula(book,sheet.id,'=TYPE(A10000)'),1);
});

test('deferred dependencies preserve lazy branches and evaluate each volatile expression once',()=>{
  const book=fixture({A1:1,B1:'=IF(FALSE,B1,A500)',C1:'=IFERROR(1/0,A500)',D1:'=RAND()+A500'}),sheet=book.sheets[0];
  for(let i=2;i<=500;i++)sheet.cells['A'+i]={raw:`=A${i-1}+1`};
  assert.equal(evaluateCell(book,sheet.id,'B1'),500);
  assert.equal(evaluateCell(book,sheet.id,'C1'),500);
  const original=Math.random;let calls=0;
  try{Math.random=()=>{calls++;return .25;};assert.equal(evaluateCell(book,sheet.id,'D1'),500.25);assert.equal(calls,1);}
  finally{Math.random=original;}
  sheet.cells.E1={raw:'=F1+1'};sheet.cells.F1={raw:'=E1+1'};
  assert.equal(evaluateCell(book,sheet.id,'E1'),'#CYCLE!');
  assert.equal(evaluateCell(book,sheet.id,'F1'),'#CYCLE!');
});

test('formula-heavy ranges and filtered SUBTOTAL resume without a quadratic restart',()=>{
  const book=fixture({}),sheet=book.sheets[0],count=5000;
  for(let i=1;i<=count;i++)sheet.cells['A'+i]={raw:`=${i}*2`};
  sheet.cells.B1={raw:`=SUM(A1:A${count})`};sheet.cells.B2={raw:`=SUBTOTAL(109,A1:A${count})`};
  sheet.rowHeights[1]=0;
  assert.equal(evaluateCell(book,sheet.id,'B1'),count*(count+1));
  assert.equal(evaluateCell(book,sheet.id,'B2'),count*(count+1)-4);
});
