import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator,validateWorkbook} from '../src/engine.mjs';
import {createTablePlan,prepareTableHeaderEdits,applyTableHeaderEdits} from '../src/tables.mjs';
import {prepareTableDataEdits} from '../src/table-data-edits.mjs';
import {resetConstraintIndexes} from '../src/cell-constraints.mjs';
function fixture(){const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'Qty'},B1:{raw:'Price'},C1:{raw:'Amount'},A2:{raw:'2'},B2:{raw:'3'},A3:{raw:'4'},B3:{raw:'5'}};const plan=createTablePlan(book,sheet,{range:'A1:C3',name:'Sales'});sheet.tables=[plan.table];for(const u of plan.updates)sheet.cells[u.key]=u.cell;return {book,sheet};}
function edit(book,sheet,updates){const before=JSON.stringify(book),plan=prepareTableDataEdits(book,sheet,prepareTableHeaderEdits(book,sheet,updates));assert.equal(JSON.stringify(book),before);applyTableHeaderEdits(book,sheet,plan);for(const u of plan.updates)sheet.cells[u.key]=u.cell;validateWorkbook(book);return plan;}
test('the first formula in an empty table column fills every data row, rebased from its entered position',()=>{
  const {book,sheet}=fixture();edit(book,sheet,[{key:'C3',cell:{raw:'=A3*B3'}}]);assert.equal(sheet.tables[0].columns[2].formula,'A2*B2');assert.equal(sheet.cells.C2.raw,'=A2*B2');assert.equal(createEvaluator(book,sheet.id)('C2'),6);assert.equal(createEvaluator(book,sheet.id)('C3'),20);
  edit(book,sheet,[{key:'C2',cell:{raw:'=A2*B2+1'}}]);assert.equal(sheet.cells.C3.raw,'=A3*B3');assert.equal(sheet.tables[0].columns[2].formula,'A2*B2');
});
test('typing and pasting contiguous adjacent rows grow the table, inherit calculated columns and never overwrite input',()=>{
  const {book,sheet}=fixture();edit(book,sheet,[{key:'C2',cell:{raw:'=[@Qty]*[@Price]'}}]);sheet.cells.C3.style={color:'#ff0000'};sheet.cells.C3.note='only this row';
  edit(book,sheet,[{key:'A4',cell:{raw:'3'}}]);assert.equal(sheet.tables[0].range,'A1:C4');assert.equal(sheet.cells.C4.raw,'=[@Qty]*[@Price]');assert.equal(sheet.cells.C4.note,undefined);assert.equal(sheet.cells.C4.style.color,'#ff0000');
  edit(book,sheet,[{key:'B4',cell:{raw:'8'}}]);assert.equal(createEvaluator(book,sheet.id)('C4'),24);
  edit(book,sheet,[{key:'A5',cell:{raw:'2'}},{key:'B5',cell:{raw:'4'}},{key:'C5',cell:{raw:'99'}},{key:'A6',cell:{raw:'3'}},{key:'B6',cell:{raw:'7'}}]);assert.equal(sheet.tables[0].range,'A1:C6');assert.equal(sheet.cells.C5.raw,'99');assert.equal(createEvaluator(book,sheet.id)('C6'),21);assert.equal(sheet.tables[0].filter.range.r2,5);
});
test('automatic edits do not grow through gaps, totals, another table, protection or merged cells, or replace existing column values',()=>{
  const {book,sheet}=fixture();sheet.cells.C3={raw:'5'};edit(book,sheet,[{key:'C2',cell:{raw:'=A2*B2'}}]);assert.equal(sheet.tables[0].columns[2].formula,undefined);assert.equal(sheet.cells.C3.raw,'5');
  edit(book,sheet,[{key:'A5',cell:{raw:'1'}}]);assert.equal(sheet.tables[0].range,'A1:C3');sheet.merges=['B4:C4'];edit(book,sheet,[{key:'A4',cell:{raw:'1'}}]);assert.equal(sheet.tables[0].range,'A1:C3');sheet.merges=[];sheet.tables[0].totalsRow=true;edit(book,sheet,[{key:'A4',cell:{raw:'2'}}]);assert.equal(sheet.tables[0].range,'A1:C3');
});

test('typing beside a table adds stable columns, normalizes headers and autofills a new calculated column',()=>{
 const {book,sheet}=fixture(),original=structuredClone(sheet.tables[0].columns);edit(book,sheet,[{key:'D1',cell:{raw:'Qty',note:'머리글 메모'}}]);assert.equal(sheet.tables[0].range,'A1:D3');assert.deepEqual(sheet.tables[0].columns.slice(0,3),original);assert.equal(sheet.tables[0].columns[3].id,4);assert.equal(sheet.tables[0].columns[3].name,'Qty2');assert.equal(sheet.cells.D1.raw,'Qty2');assert.equal(sheet.cells.D1.note,'머리글 메모');assert.equal(sheet.tables[0].filter.range.c2,3);
 edit(book,sheet,[{key:'E3',cell:{raw:'=[@Qty]*[@Price]'}}]);assert.equal(sheet.tables[0].range,'A1:E3');assert.equal(sheet.cells.E1.raw,'열5');assert.equal(sheet.tables[0].columns[4].formula,'[@Qty]*[@Price]');assert.equal(createEvaluator(book,sheet.id)('E2'),6);assert.equal(createEvaluator(book,sheet.id)('E3'),20);
});

test('horizontal growth respects gaps, existing totals data, arrays, merges and other tables',()=>{
 for(const kind of ['gap','total','array','merge','table','protected']){
  const {book,sheet}=fixture();let key='D2';
  if(kind==='gap')key='E2';
  if(kind==='total'){sheet.tables[0].totalsRow=true;sheet.cells.D3={raw:'표 밖의 자료'};}
  if(kind==='array')sheet.cells.D1={raw:'=1',xlsxFormula:{raw:'=1',formula:'1',shareType:'array',ref:'D1:D3'}};
  if(kind==='merge')sheet.merges=['D1:E1'];
  if(kind==='table'){const p=createTablePlan(book,sheet,{range:'D1:E3',name:'Other'});sheet.tables.push(p.table);for(const u of p.updates)sheet.cells[u.key]=u.cell;}
  if(kind==='protected')sheet.xlsxFeatures={sheetProtection:'<sheetProtection sheet="1"/>'};
  resetConstraintIndexes();edit(book,sheet,[{key,cell:{raw:'7'}}]);assert.equal(sheet.tables[0].range,'A1:C3',kind);
 }
});

test('a wider paste below the table grows only the original columns; an adjacent column paste uses its existing header',()=>{
 const {book,sheet}=fixture();sheet.cells.D1={raw:'추가'};edit(book,sheet,[{key:'A4',cell:{raw:'3'}},{key:'B4',cell:{raw:'4'}},{key:'D4',cell:{raw:'9'}}]);assert.equal(sheet.tables[0].range,'A1:C4');assert.equal(sheet.cells.D4.raw,'9');
 edit(book,sheet,[{key:'D2',cell:{raw:'1'}},{key:'D3',cell:{raw:'2'}}]);assert.equal(sheet.tables[0].range,'A1:D4');assert.equal(sheet.tables[0].columns[3].name,'추가');assert.equal(sheet.cells.D4.raw,'9');
});
