import test from 'node:test';
import assert from 'node:assert/strict';
import {androidEditorLayout} from '../src/editor-layout.mjs';

test('Android editing stays touch-sized in phone, landscape keyboard and split windows',()=>{
  for(const preference of ['auto','touch','desktop'])for(const [width,height] of [[412,915],[915,412],[839,900],[1280,599]])
    assert.equal(androidEditorLayout({preference,width,height,finePointer:true,keyboard:true}),'touch');
});
test('large Android editors follow an explicit choice or connected input devices',()=>{
  const size={width:1280,height:800};
  assert.equal(androidEditorLayout(size),'touch');
  assert.equal(androidEditorLayout({...size,finePointer:true}),'desktop');
  assert.equal(androidEditorLayout({...size,keyboard:true}),'desktop');
  assert.equal(androidEditorLayout({...size,preference:'touch',keyboard:true}),'touch');
  assert.equal(androidEditorLayout({...size,preference:'desktop'}),'desktop');
});
