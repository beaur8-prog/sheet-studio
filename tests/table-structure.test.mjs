import test from 'node:test';
import assert from 'node:assert/strict';
import {createEvaluator,address,parseAddress,transformFormulaReferences,validateWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {tableFixture} from './tables-fixture.mjs';
import {assertTableStructureAllowed,removeTablesForStructure,shiftTables} from '../src/table-structure.mjs';
function structure(book,sheet,axis,index,count,delta){
  const remove=assertTableStructureAllowed(book,sheet,axis,index,count,delta);removeTablesForStructure(book,sheet,remove);
  for(let i=0;i<count;i++){
    const cells={};for(const [key,cell] of Object.entries(sheet.cells)){const p=parseAddress(key);if(delta<0&&p[axis]===index)continue;if(p[axis]>=index)p[axis]+=delta;cells[address(p.r,p.c)]=cell;}sheet.cells=cells;
    for(const owner of book.sheets)for(const cell of Object.values(owner.cells))if(cell.raw.startsWith('='))cell.raw=transformFormulaReferences(cell.raw,{axis,index,delta,targetSheet:sheet.name,ownerSheet:owner.name});shiftTables(book,sheet.id,axis,index,delta);
  }
  validateWorkbook(book);
}
test('table rows move and inherit calculated columns; deleting data and totals updates ranges and preserved filters',async()=>{
  let book=await readSpreadsheet(await tableFixture(),'rows.xlsx'),s=book.sheets[0];s.cells.C2.note='do not duplicate';
  structure(book,s,'r',1,1,1);assert.equal(s.tables[0].range,'A1:C5');assert.equal(s.cells.C2.raw,'=[@Qty]*[@Price]');assert.equal(s.cells.C2.note,undefined);assert.equal(createEvaluator(book,s.id)('E1'),26);assert.equal(createEvaluator(book,s.id)('C5'),26);
  s.cells.A2={raw:'3'};s.cells.B2={raw:'4'};assert.equal(createEvaluator(book,s.id)('E1'),38);assert.equal(createEvaluator(book,s.id)('C5'),38);
  structure(book,s,'r',2,1,-1);assert.equal(s.tables[0].range,'A1:C4');assert.equal(createEvaluator(book,s.id)('C4'),32);
  structure(book,s,'r',3,1,-1);assert.equal(s.tables[0].totalsRow,false);assert.equal(s.tables[0].filter.range.r2,2);assert.equal(createEvaluator(book,s.id)('E1'),32);
  structure(book,s,'r',0,2,1);assert.equal(s.tables[0].range,'A3:C5');assert.equal(s.tables[0].filter.range.r1,2);assert.equal(createEvaluator(book,s.id)('E3'),32);
  book=await readSpreadsheet((await writeSpreadsheet(book)).data,'rows-roundtrip.xlsx');s=book.sheets[0];assert.equal(s.tables[0].range,'A3:C5');assert.equal(createEvaluator(book,s.id)('E3'),32);
});
test('table columns retain identities, headers, metadata and explicit deleted-column errors',async()=>{
  const book=await readSpreadsheet(await tableFixture(),'columns.xlsx'),s=book.sheets[0];
  structure(book,s,'c',1,1,1);assert.equal(s.tables[0].range,'A1:D4');assert.deepEqual(s.tables[0].columns.map(c=>c.name),['Qty','열2','Price','Amount']);assert.equal(s.cells.B1.raw,'열2');assert.equal(createEvaluator(book,s.id)('F1'),26);assert.equal(createEvaluator(book,s.id)('D4'),26);
  structure(book,s,'c',0,1,-1);assert.equal(s.tables[0].range,'A1:C4');assert.match(s.cells.C2.raw,/#REF!/);assert.equal(createEvaluator(book,s.id)('E1'),'#REF!');assert.equal(createEvaluator(book,s.id)('C4'),'#REF!');
  const out=await readSpreadsheet((await writeSpreadsheet(book)).data,'cols-roundtrip.xlsx');assert.deepEqual(out.sheets[0].tables[0].columns.map(c=>c.name),['열2','Price','Amount']);assert.equal(out.sheets[0].tables[0].filter.range.c2,2);
});
test('A1 calculated-column templates rebase to a new first data row while cross-sheet references move once',async()=>{
  const book=await readSpreadsheet(await tableFixture(),'templates.xlsx'),s=book.sheets[0],t=s.tables[0];t.columns[2].formula='A2*B2';s.cells.C2.raw='=A2*B2';s.cells.C3.raw='=A3*B3';
  structure(book,s,'r',1,1,1);assert.equal(t.range,'A1:C4');assert.equal(s.tables[0].columns[2].formula,'A2*B2');assert.equal(s.cells.C2.raw,'=A2*B2');assert.equal(s.cells.C3.raw,'=A3*B3');
  structure(book,s,'r',1,1,-1);assert.equal(s.tables[0].columns[2].formula,'A2*B2');assert.equal(s.cells.C2.raw,'=A2*B2');assert.equal(createEvaluator(book,s.id)('C4'),26);
  structure(book,s,'c',1,1,1);assert.equal(s.tables[0].columns[3].formula,'A2*C2');assert.equal(s.cells.D2.raw,'=A2*C2');
});
test('partial header deletion is refused and deleting a whole table produces ref errors outside it',async()=>{
  const book=await readSpreadsheet(await tableFixture(),'delete.xlsx'),s=book.sheets[0];assert.throws(()=>structure(book,s,'r',0,1,-1),/머리글/);assert.equal(s.tables[0].range,'A1:C4');
  structure(book,s,'c',0,3,-1);assert.equal(s.tables.length,0);assert.equal(s.cells.B1.raw,'=SUM(#REF!)');assert.equal(createEvaluator(book,s.id)('B1'),'#REF!');
});
