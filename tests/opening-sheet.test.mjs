import test from 'node:test';
import assert from 'node:assert/strict';
import {openingSheetId} from '../src/opening-sheet.mjs';
test('opening preserves a visible active sheet and otherwise chooses the first visible one without changing the file model',()=>{
  const book={activeSheetId:'hidden',sheets:[{id:'hidden',state:'veryHidden'},{id:'first'},{id:'next',state:'visible'}]},before=JSON.stringify(book);assert.equal(openingSheetId(book),'first');assert.equal(JSON.stringify(book),before);book.activeSheetId='next';assert.equal(openingSheetId(book),'next');book.sheets.forEach(s=>s.state='hidden');assert.throws(()=>openingSheetId(book),/표시되는 시트/);
});
