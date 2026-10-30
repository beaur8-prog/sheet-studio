import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createWorkbook} from '../src/engine.mjs';
import {writeSpreadsheet} from '../src/io.mjs';
const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','bulk-text-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir},text='x'.repeat(32768),count=3051,characters=count*text.length;
const book=createWorkbook();for(let i=1;i<=count;i++)book.sheets[0].cells['A'+i]={raw:text};
// Shared strings keep the fixture small; the logical document has almost 100M
// characters. Never construct or test a multi-gigabyte overflowing copy.
const fixture=(await writeSpreadsheet(book)).data;
const app=await _electron.launch(desktopLaunch(root,env)),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(60000);
await page.setViewportSize({width:1280,height:900});await app.evaluate(({ipcMain})=>{ipcMain.removeHandler('sheet:write-clipboard');ipcMain.handle('sheet:write-clipboard',(_event,payload)=>{globalThis.bulkTextClipboard=payload;});});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const stats=()=>page.evaluate(()=>sheetStudio.historyStats()),clearToast=()=>page.evaluate(()=>document.querySelector('.toast')?.remove());
const paste=async()=>{await page.locator('#grid').focus();const text=await app.evaluate(()=>bulkTextClipboard.text);await page.evaluate(text=>{const d=new DataTransfer();d.setData('text/plain',text);document.querySelector('#grid').dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:d}));},text);await settle();};
const rejected=async before=>{await page.waitForFunction(()=>document.querySelector('.toast')?.textContent.includes('100,000,000'));assert.deepEqual(await stats(),before);};
try{
 await page.locator('.cell').first().waitFor();await page.locator('#file-input').setInputFiles({name:'bulk-text.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(fixture)});await settle();assert.equal((await stats()).textChars,characters);assert.equal((await stats()).cells,count);
 await go('A1');await page.locator('[data-action="copy"]').click();await settle();const before=await stats();
 await go('C1:C100000');await clearToast();const started=Date.now();await paste();await rejected(before);const rejectedPasteMs=Date.now()-started;assert.equal(await page.evaluate(()=>sheetStudio.getCell('C1')?.raw),undefined);
 await go('A1:A100000');await clearToast();await page.keyboard.press(process.platform==='darwin'?'Meta+d':'Control+d');await settle();await rejected(before);assert.equal(await page.evaluate(()=>sheetStudio.getCell('A3052')?.raw),undefined);
 await go('A3051');await page.locator('#grid').evaluate(el=>{el.scrollTop+=300;});await settle();await clearToast();const handle=await page.locator('.fill-handle').boundingBox(),target=await page.getByRole('gridcell',{name:'A3060',exact:true}).boundingBox();assert.ok(handle&&target);await page.mouse.move(handle.x+handle.width/2,handle.y+handle.height/2);await page.mouse.down();await page.mouse.move(target.x+target.width/2,target.y+target.height/2,{steps:8});await page.mouse.up();await settle();await rejected(before);assert.equal(await page.locator('#name-box').inputValue(),'A3051');
 await page.keyboard.press('F2');await page.locator('#editor').fill('short');await page.locator('#editor').press('Enter');await settle();await go('C1');await paste();assert.equal(await page.evaluate(()=>sheetStudio.getCell('C1').raw.length),text.length);assert.equal((await stats()).textChars,characters+5);
 await page.locator('#undo').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('C1')?.raw),undefined);await page.locator('#undo').click();await settle();assert.equal((await stats()).textChars,characters);await page.locator('#redo').click();await settle();await page.locator('#redo').click();await settle();assert.equal((await stats()).textChars,characters+5);assert.deepEqual(errors,[]);
 const small=createWorkbook();small.name='valid-100k-fill';small.sheets[0].cells.A1={raw:'1'};await page.locator('#file-input').setInputFiles({name:'valid-100k-fill.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(small))});await settle();await go('A1:A100000');const filling=Date.now();await page.keyboard.press(process.platform==='darwin'?'Meta+d':'Control+d');await settle();assert.equal((await stats()).cells,100000);assert.equal(await page.evaluate(()=>sheetStudio.getCell('A100000').raw),'1');const fill100kMs=Date.now()-filling;await page.locator('#undo').click();await settle();assert.equal((await stats()).cells,1);await page.locator('#redo').click();await settle();assert.equal((await stats()).cells,100000);
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/bulk-text.png')});console.log(JSON.stringify({ok:true,case:'near-100M-character document rejects 100k-cell paste/fill and handle fill atomically, then accepts a valid replacement with undo/redo; valid 100k-cell fill remains usable',characters,fixtureBytes:fixture.length,rejectedPasteMs,fill100kMs}));
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/bulk-text-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
