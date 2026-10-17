import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {buildNativePlan,recordNativeStructure,nativeSourceChanged,nativeCachedValue} from '../src/native-plan.mjs';
const {validatePayload,rememberFile,markExport}=createRequire(import.meta.url)('../electron/vba.cjs');

async function fixture(){
  const x=new ExcelJS.Workbook(),s=x.addWorksheet('Data');s.getCell('A1').value=7;s.getCell('A2').value={formula:'A1*2',result:14};s.getCell('C1').value='keep';s.getCell('C1').font={bold:true};
  const zip=await JSZip.loadAsync(await x.xlsx.writeBuffer());
  // An opaque marker tests byte preservation only, not an executable VBA project.
  zip.file('xl/vbaProject.bin',new Uint8Array([0,1,2,255,128,90]));
  const data=await zip.generateAsync({type:'uint8array'});
  return {data,book:await readSpreadsheet(data,'source.xlsm')};
}

test('XLSM retains the complete original package through JSON and refuses implicit macro loss',async()=>{
  const {data,book}=await fixture();
  assert.deepEqual(Buffer.from(book.nativeSource.base64,'base64'),Buffer.from(data));
  assert.equal(nativeSourceChanged(book),false);
  const restored=parseWorkbook(serializeWorkbook(book));
  assert.equal(restored.nativeSource.base64,book.nativeSource.base64);
  await assert.rejects(writeSpreadsheet(book,'xlsx'),/XLSM/);
  await assert.rejects(writeSpreadsheet(book,'csv'),/XLSM/);
  assert.equal((await writeSpreadsheet(book,'xlsx',{allowMacroLoss:true})).extension,'xlsx');
});

test('native plan changes only edited fields, with removals and sheet renames',async()=>{
  const {book}=await fixture(),s=book.sheets[0];s.name='Renamed';s.cells.A1.raw='11';delete s.cells.C1;
  const p=buildNativePlan(book),sheet=p.sheets[0];
  assert.equal(sheet.sourceName,'Data');assert.equal(sheet.name,'Renamed');
  assert.deepEqual(sheet.cells.find(c=>c.address==='A1'),{address:'A1',raw:'11'});
  assert.equal(sheet.cells.find(c=>c.address==='C1').style.bold,null);
  assert.ok(!sheet.cells.some(c=>c.address==='A2'));
  validatePayload({mode:'save',source:book.nativeSource,plan:p});
});

test('structural operations adjust the comparison baseline to avoid overwriting Excel references',async()=>{
  const {book}=await fixture(),s=book.sheets[0];
  recordNativeStructure(book,s.id,'r',0,1);
  s.cells={A2:s.cells.A1,A3:{...s.cells.A2,raw:'=A2*2'},C2:s.cells.C1};
  const p=buildNativePlan(book);
  assert.deepEqual(p.sheets[0].cells,[]);
  assert.deepEqual(p.operations,[{sheetId:s.id,axis:'r',index:0,delta:1}]);
  assert.equal(book.nativeSource.baseline[0].cells.A3.raw,'=A2*2');
});

test('native bridge rejects invalid addresses, plans and unexpected style properties',async()=>{
  const {book}=await fixture();const payload={mode:'save',source:book.nativeSource,plan:buildNativePlan(book)};
  for(const cell of [{address:'XFE1',raw:'1'},{address:'A1048577',raw:'1'},{address:'A1',style:{Formula:'=EVIL()'}}]){
    const bad=structuredClone(payload);bad.plan.sheets[0].cells=[cell];assert.throws(()=>validatePayload(bad));
  }
  const bad=structuredClone(payload);bad.plan.operations=[{sheetId:'x',axis:'r',index:-1,delta:1}];assert.throws(()=>validatePayload(bad));
});

test('unsupported native functions expose only a saved result of the same formula',async()=>{
  const {book}=await fixture(),s=book.sheets[0];s.cells.A2.raw='=CustomVBA(A1)';book.nativeSource.baseline[0].cells.A2.raw=s.cells.A2.raw;
  assert.equal(nativeCachedValue(book,s.id,'A2','#NAME?'),14);
  s.cells.A2.raw='=CustomVBA(A1)+1';assert.equal(nativeCachedValue(book,s.id,'A2','#NAME?'),undefined);
  assert.equal(nativeCachedValue(book,s.id,'A2',17),undefined);
});

test('Windows exports preserve internet origin and reject trust tokens for changed bytes',{skip:process.platform!=='win32'},async()=>{
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'sheet-studio-zone-test-')),source=path.join(dir,'source.xlsm'),out=path.join(dir,'output.xlsm');
  const bytes=Buffer.from('PK test data'),zone='[ZoneTransfer]\r\nZoneId=3\r\nHostUrl=https://example.invalid/test.xlsm\r\n';
  try{
    await fs.writeFile(source,bytes);await fs.writeFile(source+':Zone.Identifier',zone);await fs.writeFile(out,bytes);
    const token=await rememberFile(source,bytes);await markExport(out,bytes,token);assert.equal(await fs.readFile(out+':Zone.Identifier','utf8'),zone);
    await markExport(out,Buffer.from('PK changed data'),token);assert.equal(await fs.readFile(out+':Zone.Identifier','utf8'),'[ZoneTransfer]\r\nZoneId=3\r\n');
  }finally{await fs.unlink(source);await fs.unlink(out);await fs.rmdir(dir);}
});

test('native chart plan preserves untouched charts and validates size, type, deletion and rebuild requests',async()=>{
  const {book}=await fixture(),s=book.sheets[0];
  const old={id:'old-chart',nativeName:'Chart 1',type:'bar',range:'A1:B2',x:10,y:20,width:500,height:300,title:'Before'};
  s.charts=[structuredClone(old)];book.nativeSource.baseline[0].charts=[structuredClone(old)];
  assert.deepEqual(buildNativePlan(book).sheets[0].charts,[]);
  s.charts[0].width=650;s.charts[0].title='After';
  const p=buildNativePlan(book);assert.deepEqual(p.sheets[0].charts,[{name:'Chart 1',width:650,title:'After'}]);
  validatePayload({mode:'save',source:book.nativeSource,plan:p});
  for(const patch of [{name:'Chart 1',Formula:'=UNSAFE()'},{name:'Chart 1',width:-1},{name:null,range:'A1:B2',chartType:999}]){const bad=structuredClone(p);bad.sheets[0].charts=[patch];assert.throws(()=>validatePayload({mode:'save',source:book.nativeSource,plan:bad}));}
  s.charts=[];assert.deepEqual(buildNativePlan(book).sheets[0].deletedCharts,['Chart 1']);
});
