import test from 'node:test';
import assert from 'node:assert/strict';
import {fitNumberDisplay,createNumberMeasurer} from '../src/number-display.mjs';
import {formatValue} from '../src/engine.mjs';
const measure=text=>text.length*10;
const fit=(v,style={},width=60)=>fitNumberDisplay(v,formatValue(v,style),style,width,measure);
test('width-limited numeric display preserves magnitude and complete sign/exponent',()=>{
 assert.equal(fit(123456789,{numberFormat:'0'}),'######');
 assert.equal(fit(-123456789,{numberFormat:'0'}),'######');
 assert.equal(fit(12345.6789,{},80),'12345.68');
 assert.equal(fit(123456789,{},80),'1.23E+08');
 assert.equal(fit(-123456789,{},80),'-1.2E+08');
 assert.equal(fit(123456789012,{},180),'1.23456789012E+11');
 assert.equal(fit(0.000000123456789,{},80),'1.23E-07');
 assert.equal(fit(123456789,{},20),'##');
 assert.equal(fit(0,{},20),'0');
 assert.equal(fit(12.5,{},50),'12.5');
});
test('fixed formats never lose their leading digits, symbols or dates',()=>{
 for(const style of [{format:'currency'},{format:'percent'},{numberFormat:'yyyy-mm-dd'},{numberFormat:'0.00E+00'},{numberFormat:'0;[Red](0)'},{numberFormat:'#,##0.00" 원"'}])assert.match(fit(123456789,style),/^#+$/);
 assert.equal(fit(45678,{numberFormat:'yyyy-mm-dd'},100),'2025-01-21');
 assert.equal(fit(123456789,{numberFormat:'@'}),'123456789');
 assert.equal(fitNumberDisplay('123456789','123456789',{},60,measure),'123456789');
 assert.equal(fitNumberDisplay('#DIV/0!','#DIV/0!',{},60,measure),'#DIV/0!');
 assert.equal(fit(1,{numberFormat:'0;0;""'},50),'1');
 assert.equal(fit(0,{numberFormat:'0;0;""'},50),'');
});
test('display measuring caches by font and text, bounds memory and honors multiline widths',()=>{
 let calls=0;const ctx={font:'',measureText(text){calls++;return {width:text.length*(this.font==='bold 20px serif'?20:10)};}},width=createNumberMeasurer(ctx);
 assert.equal(width('123','10px sans-serif'),30);assert.equal(width('123','10px sans-serif'),30);assert.equal(calls,1);
 assert.equal(width('123','bold 20px serif'),60);assert.equal(width('1\n123','10px sans-serif'),30);
 for(let i=0;i<4100;i++)width(String(i),'10px sans-serif');const before=calls;width('1\n123','10px sans-serif');assert.ok(calls>before);
});

test('literal newlines in numeric formats fit the actual nowrap line, not separate fragments',()=>{
 const multiline=text=>Math.max(...text.split('\n').map(line=>line.length*10));
 assert.equal(fitNumberDisplay(1,'12345\n6789',{numberFormat:'0"2345\n6789"'},60,multiline),'######');
 assert.equal(fitNumberDisplay(1,'12345\n6789',{numberFormat:'0"2345\n6789"'},100,multiline),'12345\n6789');
});
