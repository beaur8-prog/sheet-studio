import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const dataDir=path.join(root,'.package','recovery-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const executablePath=process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix()));
let app,page;const errors=[];
async function launch(){app=await _electron.launch({executablePath,args:packaged?[]:[root],env});page=await app.firstWindow();page.setDefaultTimeout(20000);page.on('pageerror',e=>errors.push(e.message));await page.locator('.cell').first().waitFor();}
const cell=()=>page.getByRole('gridcell',{name:'A1',exact:true});
try{
 await launch();
 const legacy=createWorkbook();legacy.name='기존 자동 복구';legacy.sheets[0].cells={A1:{raw:'이전 문서',style:{bold:true}}};
 await page.evaluate(text=>localStorage.setItem('sheet-studio.book',text),JSON.stringify(legacy));
 await page.reload();await cell().waitFor();await page.waitForFunction(()=>sheetStudio.getCell('A1')?.raw==='이전 문서');
 assert.equal(await page.evaluate(()=>localStorage.getItem('sheet-studio.book')),null);
 assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').style.bold),true);
 // Reload during an outstanding write: Electron must wait and retry, rather
 // than let the worker disappear along with the last edit.
 await page.locator('#name-box').fill('A1');await page.locator('#name-box').press('Enter');await page.keyboard.type('reload saved');await page.keyboard.press('Enter');
 await page.reload().catch(error=>{if(!/ERR_ABORTED|interrupted by another navigation/.test(error.message))throw error;});
 await page.waitForFunction(()=>globalThis.sheetStudio?.getCell('A1')?.raw==='reload saved');
 await page.locator('#name-box').fill('A1');await page.locator('#name-box').press('Enter');await page.keyboard.type('close saved');await page.keyboard.press('Enter');
 await app.close();await launch();
 assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'close saved');
 assert.deepEqual(errors,[]);await fs.mkdir('test-results',{recursive:true});await page.screenshot({path:'test-results/recovery-migrated.png'});
 console.log('PASS: legacy migration is durable before deletion; immediate reload and process close retain the last edit.');
}finally{if(app)await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
