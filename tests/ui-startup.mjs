import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {desktopLaunch} from './electron-launch.mjs';
import {assertBlankStart,recoverPrevious} from './recovery-ui.mjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),profile=await fs.mkdtemp(path.join(os.tmpdir(),'sheet-startup-'));
const environment={...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'};
let app,page;const errors=[],checks=[];
async function launch(){app=await _electron.launch(desktopLaunch(root,environment));page=await app.firstWindow();page.setDefaultTimeout(60000);page.on('pageerror',e=>errors.push(e.message));await page.locator('.cell').first().waitFor();}
async function restart(){await app.close();app=null;await launch();await assertBlankStart(page);}
async function enter(ref,text){await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await page.keyboard.type(text);await page.keyboard.press('Enter');await page.evaluate(()=>sheetStudio.flushRecovery());}
async function snapshot(){return page.evaluate(async()=>{
 const db=await new Promise((resolve,reject)=>{const open=indexedDB.open('sheet-studio-recovery');open.onsuccess=()=>resolve(open.result);open.onerror=()=>reject(open.error);});
 try{return await new Promise((resolve,reject)=>{const tx=db.transaction(['state','chunks'],'readonly'),current=tx.objectStore('state').get('current'),previous=tx.objectStore('state').get('previous'),keys=tx.objectStore('chunks').getAllKeys();tx.oncomplete=()=>resolve({current:current.result,previous:previous.result,keys:keys.result});tx.onabort=()=>reject(tx.error);});}finally{db.close();}
});}
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await launch();await assertBlankStart(page);
 await recoverPrevious(page);await assertBlankStart(page);assert.match(await page.locator('.toast').innerText(),/복구할 이전 문서가 없습니다/);
 const book=createWorkbook();book.name='Previous large workbook';for(let r=1;r<=100000;r++)book.sheets[0].cells['A'+r]={raw:String(r)};
 const file=path.join(profile,'previous.sheet.json');await fs.writeFile(file,JSON.stringify(book));
 await app.evaluate(({dialog},file)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[file]});},file);
 await page.locator('[data-tab="파일"]').click();await page.locator('.menu [data-action="open"]').click();
 await page.waitForFunction(()=>sheetStudio.historyStats().cells===100000);await page.evaluate(()=>sheetStudio.flushRecovery());
 const before=await snapshot();assert.equal(before.previous,undefined);
 await restart();assert.equal(await page.evaluate(()=>sheetStudio.historyStats().cells),0);
 assert.deepEqual(await snapshot(),before,'cold blank startup must not load or replace previous cells');
 await page.reload();await page.locator('.cell').first().waitFor();await assertBlankStart(page);
 assert.deepEqual(await snapshot(),before,'reloading an unedited blank must remain blank and preserve recovery');
 await restart();assert.deepEqual(await snapshot(),before,'closing an unedited blank must preserve recovery');
 checks.push('cold launch, blank reload and repeated blank close retain the 100000-cell recovery snapshot unchanged');
 await enter('B1','new document');const edited=await snapshot();
 assert.deepEqual(edited.previous,before.current,'the first edit archives the old snapshot by reference');
 const oldGenerations=new Set(Object.values(before.current.generations));
 assert.equal(edited.keys.filter(key=>oldGenerations.has(key[0])).length,before.keys.length);
 await enter('B2','second edit');await recoverPrevious(page);
 assert.equal(await page.evaluate(()=>sheetStudio.historyStats().cells),100000);
 assert.equal(await page.evaluate(()=>sheetStudio.getCell('A100000').raw),'100000');
 assert.equal(await page.evaluate(()=>sheetStudio.getCell('B1')),undefined);
 await page.locator('#undo').click();await page.waitForFunction(()=>sheetStudio.getCell('B1')?.raw==='new document');
 await page.evaluate(()=>sheetStudio.flushRecovery());
 checks.push('new edits do not overwrite previous recovery; explicit restore retains all cells and can be undone');
 await restart();await enter('C1','next session');const rotated=await snapshot();
 const retained=new Set([...Object.values(rotated.current.generations),...Object.values(rotated.previous.generations)]);
 assert.ok(rotated.keys.every(key=>retained.has(key[0])),'old generations are pruned after the previous-document slot rotates');
 assert.ok(rotated.keys.every(key=>!oldGenerations.has(key[0])),'the obsolete 100000-cell document is no longer duplicated in storage');
 await recoverPrevious(page);assert.equal(await page.evaluate(()=>sheetStudio.getCell('B2').raw),'second edit');
 checks.push('retention is bounded to current and previous snapshots without orphaned generations');
 await page.evaluate(()=>sheetStudio.flushRecovery());await restart();
 await page.screenshot({path:path.join(root,'test-results','startup-blank.png')});
 assert.deepEqual(errors,[]);await fs.writeFile(path.join(root,'test-results','startup.json'),JSON.stringify({ok:true,platform:process.platform,packaged:process.argv.includes('--packaged'),checks,errors},null,2));
 console.log('PASS: blank startup, explicit recovery, 100000-cell preservation, undo and bounded snapshot retention.');
}finally{if(app)await app.close();await fs.rm(profile,{recursive:true,force:true});}
