import test from 'node:test';
import assert from 'node:assert/strict';
import {isAndroidRuntime,createAndroidSession,androidCommandAllowed} from '../src/android-mode.mjs';
import {createHistory} from '../src/history.mjs';
import {createWorkbook} from '../src/engine.mjs';

test('Android UI requires Android runtime, not touch capability or window width',()=>{
  assert.equal(isAndroidRuntime({navigator:{userAgent:'Mozilla/5.0 (Linux; Android 14; Pixel 8)'}}),true);
  assert.equal(isAndroidRuntime({navigator:{userAgentData:{platform:'Android'}}}),true);
  for(const userAgent of ['Mozilla Windows NT 10.0','Macintosh Intel Mac OS X','iPad Mobile','Linux x86_64'])assert.equal(isAndroidRuntime({navigator:{userAgent,maxTouchPoints:10}}),false);
  assert.equal(isAndroidRuntime({desktop:{},navigator:{userAgent:'Android'}}),false);
  assert.equal(isAndroidRuntime({capacitor:{getPlatform:()=> 'ios'},navigator:{userAgent:'Android'}}),false);
  assert.equal(isAndroidRuntime({capacitor:{getPlatform:()=> 'android'},navigator:{}}),true);
});
test('viewer navigation state does not modify workbook metadata, history, or dirty markers',()=>{
  const session=createAndroidSession(),initial=createWorkbook(),before=structuredClone(initial),history=createHistory(initial,{canMutate:()=>session.mode==='edit'});
  session.open(initial.activeSheetId);session.activate('other',{selection:{r:4,c:2,r2:8,c2:3},left:120,top:75,zoom:140});
  const restored=session.activate(initial.activeSheetId);assert.equal(restored.top,75);assert.equal(restored.zoom,140);
  assert.deepEqual(history.raw,before);assert.equal(history.canUndo,false);assert.equal(history.takeDirty().size,0);
  assert.throws(()=>{history.book.sheets[0].cells.A1={raw:'changed'};},/보기 모드/);
  assert.throws(()=>{history.book.activeSheetId='other';},/보기 모드/);
  assert.equal(androidCommandAllowed('copy','view'),true);assert.equal(androidCommandAllowed('find','view'),true);
  for(const command of ['paste','cut','bold','clear','rename-sheet','validation-value:0'])assert.equal(androidCommandAllowed(command,'view'),false);
});
test('returning to view retains edits and undo while rejecting late writes and obsolete results',()=>{
  const session=createAndroidSession(),history=createHistory(createWorkbook(),{canMutate:()=>session.mode==='edit'});
  session.setMode('edit');history.begin();history.book.sheets[0].cells.A1={raw:'한글'};history.commit();session.markEdited();
  const ticket=session.ticket('doc',1);session.setMode('view');
  assert.equal(session.edited,true);assert.equal(history.canUndo,true);assert.equal(history.book.sheets[0].cells.A1.raw,'한글');
  assert.equal(session.current(ticket,'doc',1),false);assert.equal(session.current(session.ticket('doc',2),'other',2),false);
  assert.throws(()=>{delete history.book.sheets[0].cells.A1;},/보기 모드/);assert.equal(history.undo(),false);
  session.setMode('edit');assert.equal(history.undo(),true);assert.equal(history.book.sheets[0].cells.A1,undefined);
});
