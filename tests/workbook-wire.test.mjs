import test from 'node:test';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {createWorkbook,validateWorkbook} from '../src/engine.mjs';
import {encodeWorkbookWire,decodeWorkbookWire} from '../src/workbook-wire.mjs';
import {fixture as pivotFixture,spec} from './fixtures/pivot.mjs';
import {applyPivot} from '../src/pivot.mjs';
import {attachNativeSource} from '../src/native-plan.mjs';
import {applyValidationChange,planValidationChange} from '../src/data-validation.mjs';
const hash=v=>createHash('sha256').update(v).digest('hex');
test('worker transfer keeps the application document representation, Unicode, errors, formulas, number formats and validation XML',()=>{
  const book=createWorkbook(),sheet=book.sheets[0];book.date1904=true;sheet.cells={A1:{raw:'한글 😀 \u0000 \ud800',note:'두 줄\n메모'},B1:{raw:"'001",style:{format:'text'}},C1:{raw:'=#N/A'},D1:{raw:'="#N/A"'},E1:{raw:'=SEQUENCE(3)'},G1:{raw:'0',style:{format:'date',numberFormat:'yyyy-mm-dd'}}};
  applyValidationChange(sheet,planValidationChange(book,sheet,'J1:J9',{type:'list',source:'예,아니오',showInput:true,prompt:'줄 1\n줄 2'}));
  const before=JSON.stringify(book),bytes=encodeWorkbookWire(book),after=decodeWorkbookWire(bytes);assert.equal(hash(JSON.stringify(after)),hash(before));assert.equal(after.sheets[0].cells.A1.raw.charCodeAt(after.sheets[0].cells.A1.raw.length-1),0xd800);assert.equal(after.sheets[0].xlsxFeatures.dataValidations,sheet.xlsxFeatures.dataValidations);validateWorkbook(after);
});
test('pivot metadata and opaque macro source bytes keep identical hashes after worker transfer',()=>{
  const book=pivotFixture(),target={id:'report',name:'피벗',charts:[]};applyPivot(book,target,spec());book.sheets.push(target);const input=new Uint8Array([0,1,2,255,3,200]);attachNativeSource(book,input,'source.xlsm');const before=JSON.stringify(book),after=decodeWorkbookWire(encodeWorkbookWire(book));assert.equal(JSON.stringify(after),before);assert.equal(hash(Buffer.from(after.nativeSource.base64,'base64')),hash(input));assert.deepEqual(after.sheets[1].pivot,target.pivot);
});
test('malformed wire types, UTF-8 and JSON fail instead of returning a partially decoded document',()=>{
  for(const input of ['{}',new Uint8Array([255]),new TextEncoder().encode('{bad')])assert.throws(()=>decodeWorkbookWire(input));
});
