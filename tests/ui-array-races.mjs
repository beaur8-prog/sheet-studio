import {assertBlankStart,recoverPrevious} from './recovery-ui.mjs';
import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','array-races-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:profile};delete env.ELECTRON_RUN_AS_NODE;
let app,page;const errors=[],checks=[];async function launch(){app=await _electron.launch(desktopLaunch(root,env));page=await app.firstWindow();page.setDefaultTimeout(20000);page.on('pageerror',e=>errors.push(e.message));await page.locator('.cell').first().waitFor();}
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref),model=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells);
let sequence=0;
async function fresh(){const b=createWorkbook();b.name='Array race '+ ++sequence;await page.locator('#file-input').setInputFiles({name:b.name+'.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,b.name);await settle();}
async function start(formula='=SEQUENCE(100000)'){await page.keyboard.press('F2');await page.locator('#editor').fill(formula);await page.locator('#editor').press('Enter');}
async function nativeSave(name){const file=path.join(profile,name+'.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.evaluate(()=>document.querySelector('.toast')?.remove());await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.send('sheet:menu-action','save'));await page.locator('.toast').filter({hasText:'파일을 저장했습니다.'}).waitFor();await settle();assert.ok((await fs.stat(file)).size>0);return file;}
try{
 await launch();await fs.mkdir(path.join(root,'test-results'),{recursive:true});
 for(let i=0;i<10;i++){
  await fresh();await start();await page.keyboard.type('x');await page.keyboard.press('Enter');await settle();assert.equal(await value('A1'),1);assert.equal((await model()).A2,undefined);
  checks.push('immediate follower edit rejected '+i);
  await fresh();await start();await page.keyboard.press('ArrowRight');const text=i%2?'한글':'42';await page.keyboard.insertText(text);await page.keyboard.press('Enter');await settle();assert.equal(await value('A1'),1);assert.equal((await model()).B2.raw,text);checks.push('unrelated immediate input retained '+i);
 }
 await fresh();await start('=SEQUENCE(100000,2)');await page.keyboard.type('x');await page.keyboard.press('Enter');await settle();assert.equal(await value('B100000'),200000);assert.equal((await model()).A2,undefined);
 await fresh();await start('=SEQUENCE(100000)');await page.evaluate(()=>{const data=new DataTransfer();data.setData('text/plain','paste obstruction');document.querySelector('#grid').dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));});await settle();assert.equal(await value('A1'),1);assert.equal((await model()).A2,undefined);checks.push('immediate synthetic paste cannot overwrite pending array results');
 await fresh();await start();const file=await nativeSave('immediate');await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('A100000'),100000);assert.equal(Object.keys(await model()).length,1);checks.push('native Save waits for fresh array constraints and writes a reopenable file');
 await fresh();await start();await page.keyboard.press('ArrowRight');await page.keyboard.insertText('한글 유지');const edited=await nativeSave('editing');await page.locator('#file-input').setInputFiles(edited);await settle();assert.equal((await model()).B2.raw,'한글 유지');assert.equal(await value('A100000'),100000);checks.push('native Save commits a legitimate in-progress Korean edit after array calculation');
 await fresh();await start();await page.locator('[data-tab="삽입"]').click();await settle();assert.equal(await page.locator('[data-tab="삽입"]').evaluate(el=>el.classList.contains('active')),true);checks.push('ribbon tab intent survives pending array calculation');
 await fresh();await page.locator('[data-action="add-sheet"]').first().click();await settle();await page.getByRole('button',{name:'Sheet1',exact:true}).click();await settle();await start();await page.keyboard.press('ArrowRight');await page.keyboard.insertText('한글 시트');await page.getByRole('button',{name:'Sheet2',exact:true}).click();await settle();assert.equal(await page.evaluate(()=>{const b=sheetStudio.getWorkbook();return b.sheets.find(s=>s.id===b.activeSheetId).name;}),'Sheet2');assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.B2.raw),'한글 시트');checks.push('sheet switch commits a legitimate pending edit and preserves the chosen sheet');
 await fresh();await start();await page.keyboard.press('ArrowRight');await page.keyboard.insertText('종료 직전');await app.close();app=null;await launch();await assertBlankStart(page);await recoverPrevious(page);await settle();assert.equal((await model()).B2.raw,'종료 직전');assert.equal(await value('A100000'),100000);checks.push('normal close flushes a pending edit, which is explicitly recoverable after a blank restart');
 assert.deepEqual(errors,[]);await fs.writeFile(path.join(root,'test-results/array-races.json'),JSON.stringify({at:new Date().toISOString(),platform:process.platform,checks},null,2));console.log('PASS: pending spill followers protected before calculation responses; unrelated numeric/Korean input, two-dimensional output, isolated paste and native Save retain correct data');
}catch(error){await page.screenshot({path:path.join(root,'test-results/array-races-failure.png')});throw error;}
finally{if(app)await app.close();await fs.rm(profile,{recursive:true,force:true});}
