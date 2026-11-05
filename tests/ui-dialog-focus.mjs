import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {desktopLaunch} from './electron-launch.mjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','dialog-focus-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:dataDir,STUDIO_DISABLE_UPDATES:'1'}));
const page=await app.firstWindow(),errors=[];page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(20000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const compose=(selector,type)=>page.locator(selector).evaluate((node,type)=>node.dispatchEvent(new CompositionEvent(type,{bubbles:true,data:'한'})),type);
const imeKey=(selector,key,flag=true)=>page.locator(selector).evaluate((node,{key,flag})=>node.dispatchEvent(new KeyboardEvent('keydown',{key,bubbles:true,cancelable:true,isComposing:flag,keyCode:flag?229:13})),{key,flag});
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();const book=createWorkbook();book.name='dialog-focus';book.sheets[0].cells.A1={raw:'한글',style:{}};
  await page.locator('#file-input').setInputFiles({name:book.name+'.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='dialog-focus');await settle();
  await go('B2');await page.locator('.search-command').click();await page.locator('#find-text').fill('한');const before=await page.locator('#find-status').innerText();
  await compose('#find-text','compositionstart');await imeKey('#find-text','Enter',false);await imeKey('#find-text','Escape',true);assert.equal(await page.getByRole('dialog').count(),1);assert.equal(await page.locator('#find-status').innerText(),before);
  await compose('#find-text','compositionend');await page.locator('#find-text').press('Enter');await settle();assert.match(await page.locator('#find-status').innerText(),/1개 발견/);await page.keyboard.press('Escape');await page.getByRole('dialog').waitFor({state:'detached'});
  assert.ok(['editor','grid'].includes(await page.evaluate(()=>document.activeElement.id)));await page.keyboard.insertText('안전');await page.keyboard.press('Enter');await settle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'안전');
  await go('B2');await page.keyboard.press('ControlOrMeta+1');await page.getByRole('dialog',{name:'셀 서식',exact:true}).waitFor();await page.locator('[data-format-tab="number"]').click();await page.locator('#format-format').selectOption('custom');await page.locator('#format-numberFormat').fill('0.00 "한글"');
  await compose('#format-numberFormat','compositionstart');await imeKey('#format-numberFormat','Enter',false);await imeKey('#format-numberFormat','Escape',true);assert.equal(await page.getByRole('dialog').count(),1);
  await compose('#format-numberFormat','compositionend');await page.keyboard.press('Escape');await page.getByRole('dialog').waitFor({state:'detached'});await settle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('B2')),undefined);
  await page.locator('[data-tab="도움말"]').click();await page.getByRole('button',{name:'파일 호환성',exact:true}).click();await page.getByRole('dialog',{name:'파일 호환성',exact:true}).waitFor();
  assert.equal(await page.evaluate(()=>!!document.activeElement.closest('[role="dialog"]')),true);await page.keyboard.press('Shift+Tab');assert.equal(await page.evaluate(()=>!!document.activeElement.closest('[role="dialog"]')),true);
  await page.screenshot({path:path.join(root,'test-results/dialog-keyboard-focus.png')});await page.keyboard.press('Escape');await page.getByRole('dialog').waitFor({state:'detached'});assert.deepEqual(errors,[]);
  console.log('PASS: composition Enter/Escape do not apply/close dialogs; normal Enter searches; Escape restores editable focus; format cancel unchanged; no-input dialog has initial focus and Tab containment. Synthetic IME events, not an OS IME run.');
}catch(error){await page.screenshot({path:path.join(root,'test-results/dialog-focus-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
