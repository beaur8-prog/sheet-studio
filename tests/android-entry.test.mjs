import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator} from '../src/engine.mjs';
import {mobileAggregateFormula,mobileAggregateEntry,mobileAggregateSuggestion} from '../src/android-aggregate.mjs';
import {createAndroidFullscreen} from '../src/android-fullscreen.mjs';

test('mobile aggregates use all selected rows and columns without replacing source cells',()=>{
  const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'1'},B1:{raw:'2'},A2:{raw:'3'},B2:{raw:'4'}};
  const before=JSON.stringify(book),suggestion=mobileAggregateSuggestion(sheet,{r:0,c:0,r2:1,c2:1},createEvaluator(book,sheet.id));
  assert.deepEqual(suggestion,{range:'A1:B2',target:'A3'});
  assert.equal(mobileAggregateFormula('SUM',suggestion.range),'=SUM(A1:B2)');
  assert.equal(mobileAggregateFormula('AVERAGE',suggestion.range),'=AVERAGE(A1:B2)');
  assert.equal(JSON.stringify(book),before);
});
test('aggregate suggestions infer vertical and horizontal runs at a blank destination',()=>{
  const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'1'},A2:{raw:'2'},A3:{raw:'3'},C1:{raw:'10'},D1:{raw:'20'}};
  const evaluate=createEvaluator(book,sheet.id),at=(r,c)=>({r,c,r2:r,c2:c});
  assert.deepEqual(mobileAggregateSuggestion(sheet,at(3,0),evaluate),{range:'A1:A3',target:'A4'});
  assert.deepEqual(mobileAggregateSuggestion(sheet,at(0,4),evaluate),{range:'C1:D1',target:'E1'});
  assert.deepEqual(mobileAggregateSuggestion(sheet,at(2,0),evaluate),{range:'A1:A3',target:'A4'});
});
test('aggregate entry normalizes references and rejects circular, invalid, and injected input',()=>{
  assert.deepEqual(mobileAggregateEntry('AVERAGE','$b$3:$a$1','c4'),{raw:'=AVERAGE(A1:B3)',ref:'C4',point:{r:3,c:2}});
  assert.throws(()=>mobileAggregateEntry('SUM','A1:B3','B2'),/순환 참조/);
  assert.throws(()=>mobileAggregateFormula('SUM','A1:B2:C3'),/계산 범위/);
  assert.throws(()=>mobileAggregateFormula('SUM','A1)+HYPERLINK("url")'),/계산 범위/);
  assert.throws(()=>mobileAggregateFormula('INVALID','A1:A2'),/함수/);
  assert.throws(()=>mobileAggregateEntry('SUM','A1:A3','A1048577'),/결과 셀/);
});
test('Android fullscreen uses an acknowledged native status-bar request or web fullscreen',async()=>{
  const calls=[],doc={fullscreenElement:null,documentElement:{async requestFullscreen(options){calls.push(options);doc.fullscreenElement=this;}},async exitFullscreen(){doc.fullscreenElement=null;},addEventListener(){}};
  let native;
  const win={matchMedia:()=>({matches:false}),sheetStudioAndroid:{protocolVersion:1,platform:'android',async setFullscreen(options){native=options;return {enabled:options.enabled};}}};
  const host=createAndroidFullscreen({document:doc,window:win});await host.initialize();assert.equal(host.active(),true);assert.equal(native.hideStatusBar,true);assert.equal(native.hideNavigationBar,false);await host.toggle();assert.equal(host.active(),false);assert.equal(calls.length,0);
  const browser=createAndroidFullscreen({document:doc,window:{matchMedia:()=>({matches:false})}});await browser.initialize();assert.equal(calls.length,0);await browser.toggle();assert.deepEqual(calls,[{navigationUI:'hide'}]);await browser.toggle();assert.equal(browser.active(),false);
});
test('fullscreen does not report native success when the host rejects or fails to acknowledge it',async()=>{
  const doc={documentElement:{},addEventListener(){}},win={sheetStudioAndroid:{protocolVersion:1,platform:'android',setFullscreen:async()=>({enabled:false})}};
  const control=createAndroidFullscreen({document:doc,window:win});await assert.rejects(control.initialize(),/확인하지/);assert.equal(control.active(),false);
});
