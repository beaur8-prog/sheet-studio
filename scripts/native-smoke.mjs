import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import JSZip from 'jszip';
import {readSpreadsheet} from '../src/io.mjs';
import {buildNativePlan,recordNativeStructure} from '../src/native-plan.mjs';
const vba=createRequire(import.meta.url)('../electron/vba.cjs'),exec=promisify(execFile);
const root=path.resolve(import.meta.dirname,'..'),dir=path.join(root,'.verification');
const ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
await fs.mkdir(dir,{recursive:true});
const supplied=process.argv[2],source=supplied?path.resolve(supplied):path.join(dir,'native-'+Date.now()+'.xlsm');
let fixture;
if(supplied)fixture={hasMacro:true,macroReason:''};
else {const {stdout}=await exec(ps,['-NoLogo','-NoProfile','-NonInteractive','-STA','-File',path.join(root,'scripts/create-native-fixture.ps1'),'-Destination',source],{windowsHide:true,timeout:60000});fixture=JSON.parse(stdout.trim());}
const original=await fs.readFile(source);
let current;
let macroExecuted=false,macroBlockReason='';
async function waitFor(states,limit=45000){const end=Date.now()+limit;while(Date.now()<end){const s=await vba.status(current);if(s.state==='error')throw Error(s.message);if(states.includes(s.state))return s;await new Promise(r=>setTimeout(r,250));}throw Error('Excel job timeout');}
try {
  const book=await readSpreadsheet(original,'native-source.xlsm');book.nativeSource.originToken=await vba.rememberFile(source,original);
  book.sheets[0].cells.A1.raw='11';book.sheets[0].cells.D1={raw:'Korean 한글',style:{bold:true,background:'#abc'}};
  book.sheets[0].columnWidths[3]=160;book.sheets[0].rowHeights[0]=32;
  recordNativeStructure(book,book.sheets[0].id,'r',10,1);
  book.sheets.push({id:'audit-test',name:'Audit',cells:{A1:{raw:'42',style:{}},A3:{raw:'merged',style:{}}},columnWidths:{},rowHeights:{},merges:['A3:B3'],charts:[]});
  current=(await vba.start({mode:'save',source:book.nativeSource,plan:buildNativePlan(book)})).id;
  await waitFor(['completed']);
  const saved=await vba.result(current),savedBytes=Buffer.from(saved.data),reread=await readSpreadsheet(savedBytes,saved.name);
  assert.equal(reread.sheets[0].cells.A1.raw,'11');assert.equal(reread.sheets[0].cells.A2.cachedValue,22);
  assert.equal(reread.sheets[0].cells.D1.raw,'Korean 한글');assert.equal(reread.sheets[0].cells.D1.style.bold,true);
  assert.equal(reread.sheets[0].cells.C1.raw,'preserve');
  assert.equal(reread.sheets.find(s=>s.name==='Audit').cells.A1.raw,'42');
  assert.deepEqual(reread.sheets.find(s=>s.name==='Audit').merges,['A3:B3']);
  // Excel COM and OOXML widths use different font/padding units.
  assert.ok(Math.abs(reread.sheets[0].columnWidths[3]-160)<7);
  assert.ok(Math.abs(reread.sheets[0].rowHeights[0]-32)<1);
  const zip=await JSZip.loadAsync(savedBytes);assert.ok(zip.file('xl/charts/chart1.xml'));
  if(fixture.hasMacro)assert.ok(zip.file('xl/vbaProject.bin'),'Excel retains the VBA project');
  assert.match(await zip.file('xl/workbook.xml').async('string'),/OriginalInput/);
  assert.deepEqual(await fs.readFile(source),original);
  await fs.writeFile(path.join(dir,'native-saved.xlsm'),savedBytes);
  await new Promise(r=>setTimeout(r,1000));
  reread.nativeSource.originToken=saved.originToken;
  current=(await vba.start({mode:'interactive',source:reread.nativeSource,plan:buildNativePlan(reread)})).id;
  await waitFor(['ready']);
  if(fixture.hasMacro){
    await vba.command(current,{type:'run',name:'Smoke'});
    // Wait until this command has been consumed, including fast macros.
    await new Promise(r=>setTimeout(r,1200));
    const state=await waitFor(['ready']);
    if(state.message.startsWith('Excel:')){
      assert.match(state.message,/매크로를 실행할 수 없습니다|Cannot run the macro/i);
      macroBlockReason=state.message;
    }else macroExecuted=true;
  }
  await vba.command(current,{type:'finish'});await waitFor(['completed']);
  const finished=await vba.result(current),result=await readSpreadsheet(new Uint8Array(finished.data),finished.name);
  if(macroExecuted)assert.equal(result.sheets[0].cells.B2.raw,'33');
  if(macroBlockReason)assert.ok(!result.sheets[0].cells.B2,'Blocked macro did not change the result');
  assert.deepEqual(await fs.readFile(source),original);
  result.nativeSource.originToken=finished.originToken;
  current=(await vba.start({mode:'interactive',source:result.nativeSource,plan:buildNativePlan(result)})).id;
  await waitFor(['ready']);await vba.cancel(current);assert.equal((await vba.status(current)).state,'cancelled');
  const starting=vba.start({mode:'interactive',source:result.nativeSource,plan:buildNativePlan(result)});
  const shutdown=vba.closeAll();current=(await starting).id;await shutdown;
  assert.equal((await vba.status(current)).state,'cancelled');
  const report={ok:true,save:true,interactive:true,cancellation:true,shutdownDuringStart:true,sheetAndStructureEdits:true,sourceUnchanged:true,chartAndNamePreserved:true,vbaProjectRetained:fixture.hasMacro,macroExecuted,macroBlockReason,macroFixtureLimitation:fixture.macroReason};
  await fs.writeFile(path.join(dir,'native-report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));
} finally {await vba.closeAll();}
