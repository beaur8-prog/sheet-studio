import test from 'node:test';
import assert from 'node:assert/strict';
import {calendarEntry} from '../src/calendar-entry.mjs';
import {createWorkbook,evaluateFormula} from '../src/engine.mjs';
import {pastePlan} from '../src/cell-paste.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';
const date={format:'date',numberFormat:'yyyy-mm-dd'},time={format:'date',numberFormat:'hh:mm:ss'};
test('date entry uses the documented serials, fake 1900 leap day and 1462-day offset without locale or time zone guessing',()=>{
  for(const [raw,expected] of [['1900-01-01',1],['1900-02-28',59],['1900-02-29',60],['1900-03-01',61],['1904-01-01',1462],['2011-07-05',40729],['2026-11-05',46331]]){
    assert.equal(Number(calendarEntry(raw,date).raw),expected);assert.equal(Number(calendarEntry(raw,date,{date1904:true}).raw),expected-1462);
  }
  assert.equal(calendarEntry('2026-03-08 12:00',date).raw,'46089.5');assert.equal(calendarEntry('2026-11-01T12:00',date).raw,'46327.5');
  for(const raw of ['2026-02-30','1900-02-30','2001-02-29','2026-13-01','2026-11-00','2026-11-05 25:00','2026-11-05 12:60','2026-11-05 12:30:60','1899-12-31','11/5/26'])assert.equal(calendarEntry(raw,date),null,raw);
});
test('time entry supports seconds and milliseconds and remains a fraction of a day',()=>{
  for(const [raw,seconds] of [['00:00',0],['12:30',45000],['23:59:59',86399],['13:45:30.125',49530.125]])assert.ok(Math.abs(Number(calendarEntry(raw,time).raw)-seconds/86400)<1e-12,raw);
  for(const raw of ['24:00','25:00','12:60','12:30:60','12:30:00.1234','12:30 AM','1:2'])assert.equal(calendarEntry(raw,time),null,raw);
});
test('General, explicit General, Text, quoted values and formulas retain their types; validation can opt in a General destination',()=>{
  for(const style of [{},{format:'general'},{format:'general',numberFormat:'yyyy-mm-dd'},{format:'text'},{numberFormat:'@'}])assert.equal(calendarEntry('2026-11-05',style),null);
  assert.equal(calendarEntry("'2026-11-05",date),null);assert.equal(calendarEntry('=DATE(2026,11,5)',date),null);
  assert.deepEqual(calendarEntry('2026-11-05',{}, {kind:'date'}),{raw:'46331',style:{format:'date',numberFormat:'yyyy-mm-dd'}});
  assert.deepEqual(calendarEntry('12:30',{}, {kind:'time'}),{raw:String(12.5/24),style:{format:'date',numberFormat:'hh:mm'}});
  assert.equal(calendarEntry('2026-11-05',{format:'text'},{kind:'date'}),null);
});
test('plain-text paste recognizes date/time destinations while values-only rich paste keeps text identity',async()=>{
  const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'',style:date},B1:{raw:'',style:time},C1:{raw:'',style:{format:'text'}}};
  const plan=pastePlan(sheet,{plainText:true,r:0,c:0,data:[[{raw:'2026-11-05'},{raw:'12:30'},{raw:'2026-11-05'}]]},{r:0,c:0});
  for(const {key,cell} of plan.updates)sheet.cells[key]=cell;assert.equal(sheet.cells.A1.raw,'46331');assert.equal(Number(sheet.cells.B1.raw),12.5/24);assert.equal(sheet.cells.C1.raw,'2026-11-05');
  const values=pastePlan(sheet,{r:0,c:0,data:[[{raw:'2026-11-05'}]],values:[['2026-11-05']]},{r:0,c:0},{mode:'values'});assert.equal(values.updates[0].cell.raw,'2026-11-05');
  const restored=await readSpreadsheet((await writeSpreadsheet(book)).data,'calendar.xlsx');assert.equal(evaluateFormula(restored,restored.sheets[0].id,'=ISNUMBER(A1)'),true);assert.equal(evaluateFormula(restored,restored.sheets[0].id,'=HOUR(B1)'),12);assert.equal(restored.sheets[0].cells.C1.raw,'2026-11-05');
});
