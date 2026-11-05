import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {desktopLaunch} from './electron-launch.mjs';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','entry-formats-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'})),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(20000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const enter=async(ref,raw)=>{await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();};
const text=ref=>page.getByRole('gridcell',{name:ref,exact:true}).innerText(),value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();const book=createWorkbook(),s=book.sheets[0];book.name='Entry formats';s.cells={A1:{raw:'',style:{format:'general'}},B1:{raw:'',style:{format:'text'}},D1:{raw:'',style:{format:'date',numberFormat:'yyyy-mm-dd'}},D2:{raw:'=ISBLANK(D1)'},E1:{raw:'0',style:{format:'date',numberFormat:'yyyy-mm-dd'}}};
 await page.locator('#file-input').setInputFiles({name:'entry-formats.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Entry formats');await settle();await enter('A1','12.5%');await enter('B1','13%');await enter('C1','12%');assert.equal(await text('A1'),'12.5%');assert.equal(await text('C1'),'12%');assert.equal(await value('A1'),.125);assert.equal(await value('B1'),'13%');assert.equal(await value('D1'),'');assert.equal(await value('D2'),true);
 await go('A1');await page.locator('[data-tab="홈"]').click();await page.locator('#number-format').selectOption('general');await settle();assert.equal(await text('A1'),'0.125');await page.locator('#undo').click();await settle();assert.equal(await text('A1'),'12.5%');await page.locator('#redo').click();await settle();assert.equal(await text('A1'),'0.125');
 const file=path.join(profile,'entry-formats.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await text('A1'),'0.125');assert.equal(await text('B1'),'13%');assert.equal(await text('C1'),'12%');assert.equal(await value('D1'),'');assert.equal(await value('D2'),true);assert.equal(await value('E1'),0);await page.screenshot({path:path.join(root,'test-results/entry-formats-roundtrip.png')});assert.deepEqual(errors,[]);
 console.log('PASS: percent entry in General keeps typed precision; explicit General and Text remain authoritative; undo/redo; empty date stays blank; true date zero and ISBLANK survive XLSX reopen.');
}catch(error){await page.screenshot({path:path.join(root,'test-results/entry-formats-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
