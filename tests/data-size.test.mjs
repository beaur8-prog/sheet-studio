import test from 'node:test';
import assert from 'node:assert/strict';
import {estimateDataBytes} from '../src/data-size.mjs';
test('history size estimates stop at the budget without constructing a combined JSON string',()=>{
  const repeated={raw:'x'.repeat(32768),style:{fontFamily:'Font'.repeat(128)}};
  const values=Array.from({length:10000},()=>({...repeated}));
  assert.ok(estimateDataBytes(values,1024*1024)>1024*1024);
  const value={raw:'한글\n"\\',style:{bold:true},values:[1,2,null]};
  assert.ok(estimateDataBytes(value)>=JSON.stringify(value).length*2);
  const cyclic={};cyclic.self=cyclic;assert.ok(Number.isFinite(estimateDataBytes(cyclic)));
});
