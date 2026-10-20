import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook} from '../src/engine.mjs';
import {createHistory} from '../src/history.mjs';
import {cloneData} from '../src/clone.mjs';

test('delta history restores nested formatting, absent values, cells and redo branches', () => {
  const history = createHistory(createWorkbook()), book = history.book, sheet = book.sheets[0];
  assert.equal(book.sheets.find(item=>item.id===sheet.id),sheet);
  history.begin(); sheet.cells.A1 = {raw: '1', style: {bold: true}}; history.commit();
  history.begin(); sheet.cells.A1.style.color = '#ff0000'; sheet.cells.A1.raw = '2'; sheet.cells.B1 = {raw: '3'}; history.commit();
  history.undo(); assert.deepEqual(cloneData(sheet.cells), {A1: {raw: '1', style: {bold: true}}});
  history.undo(); assert.deepEqual(cloneData(sheet.cells), {});
  history.undo(true); history.undo(true); assert.equal(sheet.cells.A1.raw, '2');
  history.undo(); history.begin(); delete sheet.cells.A1.style.bold; history.commit();
  assert.equal(history.canRedo, false); history.undo(); assert.equal(sheet.cells.A1.style.bold, true);
});

test('array edits, root replacements and coalesced assignments preserve undo order', () => {
  const history = createHistory(createWorkbook()), original = cloneData(history.raw);
  history.begin(); history.book.sheets.push({...cloneData(original.sheets[0]), id: 'two', name: 'Two'}); history.commit();
  history.begin(); history.book.sheets.splice(0, 1); history.book.activeSheetId='two'; history.commit();
  history.undo(); assert.equal(history.book.sheets.length, 2); assert.equal(history.book.sheets[0].id, original.sheets[0].id);
  history.undo(); assert.deepEqual(history.raw, original);
  history.undo(true); history.undo(true); assert.equal(history.book.sheets[0].id, 'two');
  history.begin(); history.replace(createWorkbook()); history.book.name = 'new'; history.commit();
  history.undo(); assert.equal(history.book.sheets[0].id, 'two');
  history.undo(true); assert.equal(history.book.name, 'new');
  history.begin(); history.book.sheets.length = 0; history.rollback(); assert.equal(history.book.sheets.length, 1);
});

test('cell limit rejects an entire multi-sheet change and keeps the last good document', () => {
  const history = createHistory(createWorkbook(), {maxCells: 2});
  history.begin(); history.book.sheets[0].cells.A1 = {raw: '1'}; history.commit();
  history.begin(); history.book.sheets[0].cells.B1 = {raw: '2'}; history.book.sheets[0].cells.C1 = {raw: '3'};
  assert.throws(() => history.commit(), /변경을 취소/);
  assert.deepEqual(cloneData(history.book.sheets[0].cells), {A1: {raw: '1'}});
  assert.equal(history.stats.cells, 1);
  history.undo(); assert.equal(history.stats.cells, 0);
});

test('cell edits keep a small delta regardless of the size of the rest of the book', () => {
  const book = createWorkbook();
  for (let i = 1; i <= 250_000; i++) book.sheets[0].cells['A' + i] = {raw: String(i)};
  const history = createHistory(book);
  for (let i = 0; i < 65; i++) { history.begin(); history.book.sheets[0].cells.A1.raw = String(-i); history.commit(); }
  assert.equal(history.stats.undo, 60); assert.ok(history.stats.bytes < 30_000);
  for (let i = 0; i < 60; i++) history.undo();
  assert.equal(history.book.sheets[0].cells.A1.raw, '-4');
  assert.equal(history.stats.cells, 250_000);
});

test('format and whole-map replacements report exactly the cells requiring persistence', () => {
  const history = createHistory(createWorkbook()), sheet = history.book.sheets[0];
  history.begin(); sheet.cells.A1 = {raw: 'x', style: {}}; history.commit(); history.takeDirty();
  history.begin(); sheet.cells.A1.style.bold = true; history.commit();
  assert.deepEqual([...history.takeDirty().values()].map(set => [...set]), [['A1']]);
  history.undo(); assert.deepEqual([...history.takeDirty().values()].map(set => [...set]), [['A1']]);
  history.begin(); sheet.cells = {B2: {raw: 'new'}}; history.commit();
  history.undo(); assert.equal(sheet.cells.A1.raw, 'x');
});

test('incremental validation rolls back invalid cell text, formatting and document metadata', () => {
  const history=createHistory(createWorkbook()),sheet=history.book.sheets[0];
  for(const mutate of [()=>{sheet.cells.A1={raw:'x'.repeat(32769)};},()=>{sheet.cells.A1={raw:'x',style:{fontSize:0}};},()=>{history.book.name='x'.repeat(256);}]){
    history.begin();mutate();assert.throws(()=>history.commit());assert.deepEqual(cloneData(sheet.cells),{});assert.equal(history.stats.cells,0);
  }
  history.begin();sheet.cells.A1={raw:'good'};history.commit();assert.equal(sheet.cells.A1.raw,'good');
});

test('the text budget is atomic and counts remain exact through coalesced replacements and undo',()=>{
  const history=createHistory(createWorkbook(),{maxText:4}),sheet=history.book.sheets[0];
  history.begin();sheet.cells.A1={raw:'a'};sheet.cells.A1.raw='bb';sheet.cells.A1={raw:'ccc'};history.commit();
  assert.equal(history.stats.textChars,3);history.undo();assert.equal(history.stats.textChars,0);history.undo(true);assert.equal(history.stats.textChars,3);
  history.begin();sheet.cells.B1={raw:'dd'};assert.throws(()=>history.commit(),/텍스트 총량/);assert.equal(history.stats.textChars,3);assert.equal(sheet.cells.B1,undefined);
});
