import test from 'node:test';
import assert from 'node:assert/strict';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {createWorkbook} from '../src/engine.mjs';
import {assertNativeCellTransfers} from '../src/native-cell-features.mjs';
import {styleDefaultsFixture} from './fixtures/style-defaults.mjs';
import {attachNativeSource,buildNativePlan} from '../src/native-plan.mjs';
import {captureClipboardStyles,adoptClipboardStyles} from '../src/cell-clipboard.mjs';
import {pastePlan} from '../src/cell-paste.mjs';
test('native XLSM save must not silently omit copied diagonal borders',async()=>{
 const bytes=await styleDefaultsFixture(),book=await readSpreadsheet(bytes,'original.xlsx'),s=book.sheets[0];attachNativeSource(book,bytes,'original.xlsm');const data=[[structuredClone(s.cells.B2)]],clip={data,r:1,c:1,values:[[7]],xlsxStyles:captureClipboardStyles(book,s,data,[1],[1])},plan=pastePlan(s,clip,{r:1,c:3}),styles=adoptClipboardStyles(book,clip,plan.updates);book.xlsxStyles.push(...styles);for(const u of plan.updates)s.cells[u.key]=u.cell;
 assert.throws(()=>buildNativePlan(book),/상세 서식/);
});

function transfer(book,sheet,from,to,options={}){const ref=/^([A-Z]+)([0-9]+)$/.exec(from),r=Number(ref[2])-1,c=[...ref[1]].reduce((n,l)=>n*26+l.charCodeAt(0)-64,0)-1,data=[[structuredClone(sheet.cells[from]||{raw:''})]],clip={data,r,c,values:[[Number(sheet.cells[from]?.raw)||0]],xlsxStyles:captureClipboardStyles(book,sheet,data,[r],[c])},target=/^([A-Z]+)([0-9]+)$/.exec(to),plan=pastePlan(sheet,clip,{r:Number(target[2])-1,c:[...target[1]].reduce((n,l)=>n*26+l.charCodeAt(0)-64,0)-1},options),added=adoptClipboardStyles(book,clip,plan.updates,options);return {updates:plan.updates,added,styles:[...(book.xlsxStyles||[]),...added]};}
test('unsupported native transfer preflight is atomic while values-only paste retains destination features',async()=>{
 const bytes=await styleDefaultsFixture(),book=await readSpreadsheet(bytes,'styles.xlsx'),s=book.sheets[0];attachNativeSource(book,bytes,'source.xlsm');const before=JSON.stringify(book),plan=transfer(book,s,'B2','D2');assert.throws(()=>assertNativeCellTransfers(book,s,plan.updates,{styles:plan.styles}),/상세 서식/);assert.equal(JSON.stringify(book),before);
 const values=transfer(book,s,'B2','D2',{mode:'values'});assert.doesNotThrow(()=>assertNativeCellTransfers(book,s,values.updates,{styles:values.styles}));assert.equal(values.updates[0].cell.raw,'7');assert.equal(values.updates[0].cell.xlsxStyle,undefined);
 const same=transfer(book,s,'B2','B3');assert.doesNotThrow(()=>assertNativeCellTransfers(book,s,same.updates,{styles:same.styles}));
});
test('ordinary native value and supported font/color copies still produce a save plan',async()=>{
 const book=createWorkbook(),s=book.sheets[0];s.cells={A1:{raw:'7',style:{bold:true,color:'#008800'}}};const bytes=(await writeSpreadsheet(book)).data;attachNativeSource(book,bytes,'basic.xlsm');const plan=transfer(book,s,'A1','B1');assert.doesNotThrow(()=>assertNativeCellTransfers(book,s,plan.updates,{styles:plan.styles}));book.xlsxStyles=plan.styles;for(const u of plan.updates)s.cells[u.key]=u.cell;const native=buildNativePlan(book).sheets[0].cells.find(c=>c.address==='B1');assert.equal(native.raw,'7');assert.equal(native.style.bold,true);assert.equal(native.style.color,'#008800');
});
test('native guard recognizes rotation, gradients, strike/protection and rich links without blocking unchanged originals',async()=>{
 for(const style of [{font:{strike:true}},{alignment:{textRotation:30}},{alignment:{shrinkToFit:true}},{fill:{type:'gradient',gradient:'angle',degree:45,stops:[{position:0,color:{argb:'FFFF0000'}},{position:1,color:{argb:'FF0000FF'}}]}},{protection:{locked:false}},{protection:{hidden:true}}]){
  const book=createWorkbook(),s=book.sheets[0];book.xlsxStyles=[{style,baseline:{}}];s.cells={A1:{raw:'7',xlsxStyle:0}};attachNativeSource(book,new Uint8Array([1]),'guard.xlsm');assert.doesNotThrow(()=>buildNativePlan(book));const plan=transfer(book,s,'A1','B1');assert.throws(()=>assertNativeCellTransfers(book,s,plan.updates,{styles:plan.styles}),/상세 서식/);
 }
 for(const value of [{richText:[{text:'7',font:{bold:true}}]},{text:'7',hyperlink:'https://example.com',tooltip:'reference'}]){const book=createWorkbook(),s=book.sheets[0];s.cells={A1:{raw:'7',xlsxValue:{raw:'7',value}}};attachNativeSource(book,new Uint8Array([1]),'links.xlsm');const plan=transfer(book,s,'A1','B1');assert.throws(()=>assertNativeCellTransfers(book,s,plan.updates,{styles:plan.styles}),/상세 서식/);}
});

test('native source snapshots validate row/column style references with the enclosing style definitions',async()=>{
 const bytes=await styleDefaultsFixture(),book=await readSpreadsheet(bytes,'row-defaults.xlsm');const {serializeWorkbook,parseWorkbook}=await import('../src/engine.mjs');assert.doesNotThrow(()=>serializeWorkbook(book));const restored=parseWorkbook(serializeWorkbook(book));assert.deepEqual(restored.sheets[0].xlsxStyleDefaults,book.sheets[0].xlsxStyleDefaults);assert.deepEqual(restored.nativeSource.baseline[0].xlsxStyleDefaults,book.sheets[0].xlsxStyleDefaults);assert.equal(restored.nativeSource.base64,book.nativeSource.base64);
});
