import test from 'node:test';
import assert from 'node:assert/strict';
import {androidInsets,androidViewportFrame} from '../src/android-viewport.mjs';

test('Android inset snapshots separate cutouts from hidden bars and clear obsolete edges',()=>{
  const message=androidInsets({systemBars:{top:40,bottom:44},displayCutout:{right:24}});
  assert.deepEqual(message,{bars:{top:40,right:0,bottom:44,left:0},cutout:{top:0,right:24,bottom:0,left:0}});
  assert.equal(androidInsets({top:20}).bars.bottom,0);
  assert.equal(androidInsets({top:40,bottom:44}).cutout,null);
});
test('Android does not apply insets twice when the host already resized the WebView',()=>{
  const insets=androidInsets({consumedByWebView:true,systemBars:{top:40,bottom:44},displayCutout:{left:24}});
  assert.ok(Object.values(insets.bars).every(value=>value===0));
  assert.ok(Object.values(insets.cutout).every(value=>value===0));
  for(const detail of [null,{}, {top:Infinity,right:-2,bottom:500,left:'24'}])assert.ok(Object.values(androidInsets(detail).bars).every(value=>value===0));
});
test('Android uses the visible keyboard-reduced viewport including browser pan without inventing screen space',()=>{
  const frame=androidViewportFrame({innerWidth:412,innerHeight:842,visualViewport:{width:412,height:300,offsetTop:120,offsetLeft:0,scale:1}});
  assert.deepEqual(frame,{width:412,height:300,top:120,left:0});
  assert.deepEqual(androidViewportFrame({innerWidth:915,innerHeight:412}),{width:915,height:412,top:0,left:0});
});
test('Android preserves the layout viewport during pinch zoom so the document can still be panned',()=>{
  assert.deepEqual(androidViewportFrame({innerWidth:412,innerHeight:842,visualViewport:{width:206,height:421,offsetTop:80,offsetLeft:50,scale:2}}),{width:412,height:842,top:0,left:0});
});
