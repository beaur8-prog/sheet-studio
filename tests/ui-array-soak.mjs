import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createWorkbook} from '../src/engine.mjs';

// Optional longer-running regression: repeated recalculation, bounded history,
// worker export, reopen and recovery. Memory readings are observations, not a
// cross-platform memory limit or a claim that the process has no possible leak.
const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','array-soak-'+Date.now());
const env={...process.env,SHEET_STUDIO_DATA:profile};delete env.ELECTRON_RUN_AS_NODE;
let app,page;const errors=[],samples=[];
async function launch(){app=await _electron.launch(desktopLaunch(root,env));page=await app.firstWindow();page.setDefaultTimeout(20000);page.on('pageerror',e=>errors.push(e.message));await page.locator('.cell').first().waitFor();}
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
const sourceCount=()=>page.evaluate(()=>sheetStudio.historyStats().cells);
async function resize(count){await go('D1');await page.keyboard.press('F2');await page.locator('#editor').fill(String(count));await page.locator('#editor').press('Enter');await settle();}
async function verify(count){assert.equal(await value('A'+count),count);assert.equal(await value('F'+count),count);assert.equal(await value('F'+(count+1)),'');assert.equal(await value('H1'),count*(count+1)/2);assert.equal(await sourceCount(),4);}
async function sample(step){
 const collect=process.argv.includes('--memory-audit');
 const memory=await app.evaluate(async({app,BrowserWindow},collect)=>{
  let heaps;
  if(collect){
   const debug=BrowserWindow.getAllWindows()[0].webContents.debugger;debug.attach('1.3');
   try{
    await debug.sendCommand('HeapProfiler.collectGarbage');const main=await debug.sendCommand('Runtime.getHeapUsage'),workers=[];
    const {targetInfos}=await debug.sendCommand('Target.getTargets');
    for(const target of targetInfos.filter(t=>t.type==='worker')){
     const {sessionId}=await debug.sendCommand('Target.attachToTarget',{targetId:target.targetId,flatten:true});
     try{await debug.sendCommand('HeapProfiler.collectGarbage',{},sessionId);workers.push({url:target.url,...await debug.sendCommand('Runtime.getHeapUsage',{},sessionId)});}
     finally{await debug.sendCommand('Target.detachFromTarget',{sessionId});}
    }
    heaps={forcedGarbageCollection:true,main,workers};
   }finally{debug.detach();}
  }
  return {processes:app.getAppMetrics().filter(p=>p.type==='Tab').map(p=>({pid:p.pid,workingSetKiB:p.memory?.workingSetSize,privateKiB:p.memory?.privateBytes})),...(heaps?{heaps}:{})};
 },collect);
 samples.push({step,at:new Date().toISOString(),...memory});
}
async function save(file){await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.evaluate(()=>document.querySelector('.toast')?.remove());await page.locator('#save').click();await page.locator('.toast').filter({hasText:'파일을 저장했습니다.'}).waitFor();await settle();}
try{
 const start=Date.now();await fs.mkdir(path.join(root,'test-results'),{recursive:true});await launch();const b=createWorkbook();b.name='Array soak';b.sheets[0].cells={A1:{raw:'=SEQUENCE(D1)'},D1:{raw:'80000'},F1:{raw:'=FILTER(A1#,ISNUMBER(A1#))'},H1:{raw:'=SUM(F1#)'}};
 await page.locator('#file-input').setInputFiles({name:'array-soak.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await settle();await verify(80000);await sample(0);
 const file=path.join(profile,'array-soak.xlsx');
 for(let i=0;i<72;i++){
  const count=80000+(i%21)*1000;await resize(count);await verify(count);
  if((i+1)%12===0){await sample(i+1);console.log('PASS: resize',i+1,'rows',count,'source records',await sourceCount());}
  if((i+1)%24===0)await save(file);
 }
 const history=await page.evaluate(()=>sheetStudio.historyStats());assert.equal(history.undo,60);assert.ok(history.bytes<1000000,JSON.stringify(history));
 for(let i=0;i<60;i++){await page.locator('#undo').click();await settle();}await verify(91000);assert.equal(await page.locator('#undo').isDisabled(),true);
 for(let i=0;i<60;i++){await page.locator('#redo').click();await settle();}await verify(88000);assert.equal(await page.locator('#redo').isDisabled(),true);await sample('history');
 await save(file);await page.locator('#file-input').setInputFiles(file);await settle();await verify(88000);await page.evaluate(()=>sheetStudio.flushRecovery());await app.close();app=null;
 await launch();await settle();await verify(88000);await sample('restart');assert.deepEqual(errors,[]);
 await fs.writeFile(path.join(root,'test-results/array-soak'+(process.argv.includes('--memory-audit')?'-memory':'')+'.json'),JSON.stringify({at:new Date().toISOString(),platform:process.platform,iterations:72,minRows:80000,maxRows:100000,derivedArrays:2,sourceRecords:4,history,undoRedoOperations:120,milliseconds:Date.now()-start,memoryObservations:samples},null,2));
 console.log('PASS: 72 large typed-filter spill resizes, bounded 60-step history and all 120 undo/redo operations, repeated worker XLSX exports, sparse reopen and process recovery');
}catch(error){await page?.screenshot({path:path.join(root,'test-results/array-soak-failure.png')}).catch(()=>{});throw error;}
finally{if(app)await app.close();await fs.rm(profile,{recursive:true,force:true});}
