import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook} from '../src/engine.mjs';
import {findWorkbook,replaceWorkbookPlan,replacementPatch,applyReplacementPatch} from '../src/find-replace.mjs';
import {createHistory} from '../src/history.mjs';

function fixture(){const b=createWorkbook(),s=b.sheets[0];s.name='매출';s.cells={A3:{raw:'alpha',style:{}},B1:{raw:"'001",style:{}},A1:{raw:'ALPHA',style:{bold:true}},A2:{raw:'Alpha Beta',style:{}},C1:{raw:'=SUM(1,2)',style:{}},D1:{raw:'0.5',style:{format:'percent',numberFormat:'0%'}}};b.sheets.push({id:'second',name:'재고',cells:{A1:{raw:'alpha',style:{}}},charts:[],merges:[],columnWidths:{},rowHeights:{}});return b;}

test('find and replace share case/whole-cell rules and preserve literal replacement characters',()=>{
  const b=fixture(),o={sheetId:b.sheets[0].id,query:'alpha'};
  assert.equal(findWorkbook(b,o).count,3);
  assert.equal(findWorkbook(b,{...o,matchCase:true}).count,1);
  assert.equal(findWorkbook(b,{...o,wholeCell:true}).count,2);
  const plan=replaceWorkbookPlan(b,{...o,replacement:'$&'});assert.equal(plan.count,3);
  assert.equal(plan.book.sheets[0].cells.A1.raw,'$&');assert.equal(plan.book.sheets[0].cells.A2.raw,'$& Beta');assert.equal(plan.book.sheets[0].cells.A3.raw,'$&');
  assert.equal(b.sheets[0].cells.A1.raw,'ALPHA');assert.equal(plan.book.sheets[0].cells.A1.style.bold,true);
});

test('selection/workbook scope, formulas/display values and bounded row-order results',()=>{
  const b=fixture(),id=b.sheets[0].id;
  assert.deepEqual(findWorkbook(b,{sheetId:id,query:'alpha',limit:2}).results.map(r=>r.ref),['A1','A2']);
  assert.equal(findWorkbook(b,{sheetId:id,query:'alpha',limit:2}).truncated,true);
  assert.equal(findWorkbook(b,{sheetId:id,query:'alpha',scope:'workbook'}).count,4);
  assert.equal(findWorkbook(b,{sheetId:id,query:'alpha',scope:'selection',range:{r1:1,r2:1,c1:0,c2:0}}).count,1);
  assert.equal(findWorkbook(b,{sheetId:id,query:'001',wholeCell:true}).count,1);
  assert.equal(findWorkbook(b,{sheetId:id,query:'SUM',lookIn:'values'}).count,0);
  assert.equal(findWorkbook(b,{sheetId:id,query:'3',lookIn:'values',wholeCell:true}).results[0].ref,'C1');
  assert.equal(findWorkbook(b,{sheetId:id,query:'50%',lookIn:'values',wholeCell:true}).results[0].ref,'D1');
  assert.throws(()=>replaceWorkbookPlan(b,{sheetId:id,query:'3',lookIn:'values',replacement:'4'}),/검색 전용/);
});

test('find reads dynamic array followers without creating source cells',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'=SEQUENCE(3)',style:{}};
  const result=findWorkbook(b,{sheetId:s.id,query:'3',lookIn:'values',wholeCell:true});assert.equal(result.results[0].ref,'A3');assert.equal(s.cells.A3,undefined);
});

test('workbook replacement validates all sheets first and is one undoable transaction',()=>{
  const b=fixture(),history=createHistory(b),id=b.sheets[0].id;
  const before=JSON.stringify(history.book),plan=replaceWorkbookPlan(history.book,{sheetId:id,query:'alpha',replacement:'한글',scope:'workbook'});
  assert.equal(plan.count,4);history.begin();history.replace(plan.book);history.commit();assert.equal(history.book.sheets[1].cells.A1.raw,'한글');
  history.undo();assert.equal(JSON.stringify(history.book),before);history.undo(true);assert.equal(history.book.sheets[0].cells.A2.raw,'한글 Beta');
  b.sheets[1].xlsxFeatures={sheetProtection:'<sheetProtection sheet="1"/>'};
  const original=JSON.stringify(b);assert.throws(()=>replaceWorkbookPlan(b,{sheetId:id,query:'alpha',replacement:'bad',scope:'workbook'}),/보호/);assert.equal(JSON.stringify(b),original);
});

test('replacement validates dependent rules after all replacement values are staged',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'1',style:{}},B1:{raw:'1',style:{}}};
  s.xlsxFeatures={dataValidations:'<dataValidations count="1"><dataValidation type="custom" sqref="B1" showErrorMessage="1"><formula1>B1=A1</formula1></dataValidation></dataValidations>'};
  const plan=replaceWorkbookPlan(b,{sheetId:s.id,query:'1',wholeCell:true,replacement:'2'});assert.equal(plan.count,2);assert.equal(plan.book.sheets[0].cells.B1.raw,'2');
  assert.throws(()=>replaceWorkbookPlan(b,{sheetId:s.id,query:'1',wholeCell:true,replacement:'2',scope:'selection',range:{r1:0,r2:0,c1:1,c2:1}}),/입력한 값/);
});

test('replacement rejects oversized cells, partial stored arrays, pivots and invalid ranges without mutation',()=>{
  const b=fixture(),s=b.sheets[0],before=JSON.stringify(b),o={sheetId:s.id,query:'ALPHA',wholeCell:true};
  assert.throws(()=>replaceWorkbookPlan(b,{...o,replacement:'x'.repeat(32769)}),/셀|텍스트/);assert.equal(JSON.stringify(b),before);
  assert.throws(()=>findWorkbook(b,{...o,scope:'selection',range:{r1:2,r2:1,c1:0,c2:0}}),/선택/);
  s.cells.C1.xlsxFormula={raw:'=SUM(1,2)',formula:'SUM(1,2)',shareType:'array',ref:'C1:C3'};
  s.cells.C2={raw:'needle',style:{}};assert.throws(()=>replaceWorkbookPlan(b,{sheetId:s.id,query:'needle',replacement:'new'}),/배열/);
});

test('literal special-character queries are safe and nonmatching replacement has no transaction',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'[a].* $1',style:{}},B1:{raw:"'001",style:{}}};
  const plan=replaceWorkbookPlan(b,{sheetId:s.id,query:'[a].*',replacement:'$1'});assert.equal(plan.book.sheets[0].cells.A1.raw,'$1 $1');
  assert.equal(replaceWorkbookPlan(b,{sheetId:s.id,query:'001',replacement:'002'}).book.sheets[0].cells.B1.raw,"'002");
  assert.deepEqual(replaceWorkbookPlan(b,{sheetId:s.id,query:'missing',replacement:'new'}),{book:null,count:0});
});

test('small replacements keep delta undo rather than retaining a large replacement document',()=>{
  const b=createWorkbook(),s=b.sheets[0];for(let i=1;i<=1000;i++)s.cells['A'+i]={raw:i===1?'needle':'same',style:{}};
  const h=createHistory(b,{maxBytes:4096}),id=h.book.sheets[0].id,plan=replaceWorkbookPlan(h.raw,{sheetId:id,query:'needle',replacement:'replaced'}),patch=replacementPatch(h.raw,plan.book);
  assert.equal(patch.updates.length,1);h.begin();applyReplacementPatch(h.book,patch);assert.equal(h.commit().discarded,false);assert.ok(h.stats.bytes<4096);assert.equal(h.book.sheets[0].cells.A1.raw,'replaced');h.undo();assert.equal(h.book.sheets[0].cells.A1.raw,'needle');
});

test('whole-cell matching includes terminal line breaks rather than treating them as end anchors',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'alpha\n',style:{}},A2:{raw:'alpha',style:{}}};
  const o={sheetId:s.id,query:'alpha',wholeCell:true};assert.deepEqual(findWorkbook(b,o).results.map(r=>r.ref),['A2']);
  const plan=replaceWorkbookPlan(b,{...o,replacement:'new'});assert.equal(plan.count,1);assert.equal(plan.book.sheets[0].cells.A1.raw,'alpha\n');
});

test('replacing a complete stored array removes its old definition without mutating the source',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'=1',style:{},xlsxFormula:{raw:'=1',formula:'1',shareType:'array',ref:'A1:A3'}},A2:{raw:'1',style:{}},A3:{raw:'1',style:{}}};
  const before=JSON.stringify(b),plan=replaceWorkbookPlan(b,{sheetId:s.id,query:'1',replacement:'2'});
  assert.equal(plan.count,3);assert.equal(plan.book.sheets[0].cells.A1.xlsxFormula,undefined);assert.equal(JSON.stringify(b),before);
});

test('table-header replacement updates cross-sheet formulas and preservation metadata without touching the input',async()=>{
  const {tableFixture}=await import('./tables-fixture.mjs'),{readSpreadsheet}=await import('../src/io.mjs');
  const b=await readSpreadsheet(await tableFixture(),'table.xlsx'),s=b.sheets[0];b.sheets.push({id:'other',name:'Other',cells:{A1:{raw:'=1',cachedFormula:'=SUM(Sales[Amount])',cachedValue:26,style:{}},B1:{raw:'=SUM(Sales[Amount])',style:{}}},charts:[],merges:[],columnWidths:{},rowHeights:{}});
  const before=JSON.stringify(b),plan=replaceWorkbookPlan(b,{sheetId:s.id,query:'Amount',replacement:'Total',wholeCell:true}),patch=replacementPatch(b,plan.book);
  assert.equal(plan.book.sheets[1].cells.B1.raw,'=SUM(Sales[Total])');assert.equal(plan.book.sheets[1].cells.A1.cachedFormula,'=SUM(Sales[Total])');assert.equal(JSON.stringify(b),before);
  const h=createHistory(b);h.begin();applyReplacementPatch(h.book,patch);h.commit();assert.equal(h.book.sheets[1].cells.A1.cachedFormula,'=SUM(Sales[Total])');h.undo();assert.equal(h.book.sheets[1].cells.A1.cachedFormula,'=SUM(Sales[Amount])');
});
