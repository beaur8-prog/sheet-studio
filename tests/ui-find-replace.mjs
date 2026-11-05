import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {desktopLaunch} from './electron-launch.mjs';
import {createWorkbook} from '../src/engine.mjs';
import {readSpreadsheet} from '../src/io.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','find-replace-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:dataDir,STUDIO_DISABLE_UPDATES:'1'}));
const page=await app.firstWindow(),errors=[];page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(20000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const open=async()=>{await page.locator('.search-command').click();await page.getByRole('dialog',{name:'찾기 및 바꾸기',exact:true}).waitFor();};
const scan=async()=>{await page.getByRole('dialog').getByRole('button',{name:'모두 찾기 / 미리 보기',exact:true}).click();await settle();};
const close=async()=>page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();
const book=()=>page.evaluate(()=>sheetStudio.getWorkbook());
async function load(b,name){b.name=name;await page.locator('#file-input').setInputFiles({name:name+'.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,name);await settle();}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.setViewportSize({width:1000,height:680});await page.locator('.cell').first().waitFor();
  const input=createWorkbook();input.sheets[0].name='매출';input.sheets[0].cells={A1:{raw:'ALPHA',style:{bold:true}},A2:{raw:'Alpha Beta',style:{}},A3:{raw:'alpha',style:{}},B1:{raw:'=SUM(1,2)',style:{}},C1:{raw:'0.5',style:{numberFormat:'0%',format:'percent'}}};
  input.sheets.push({id:'second',name:'재고',cells:{A1:{raw:'alpha',style:{}}},charts:[],merges:[],rowHeights:{},columnWidths:{}});await load(input,'find-journey');
  await open();await page.locator('#find-text').fill('alpha');await page.locator('#replace-text').fill('한글');await page.locator('#find-scope').selectOption('workbook');
  await page.getByRole('dialog').getByRole('button',{name:'모두 바꾸기',exact:true}).click();await settle();assert.match(await page.locator('#find-status').innerText(),/4개 발견.*미리 보기/);assert.equal((await book()).sheets[0].cells.A1.raw,'ALPHA');
  await page.getByRole('dialog').getByRole('button',{name:'다음 찾기',exact:true}).click();await settle();assert.match(await page.locator('#find-status').innerText(),/매출!A1/);
  await page.locator('[data-find-result="3"]').click();await settle();assert.match(await page.locator('#find-status').innerText(),/재고!A1/);
  await page.getByRole('dialog').getByRole('button',{name:'모두 바꾸기',exact:true}).click();await settle();assert.match(await page.locator('#find-status').innerText(),/4개 셀을 변경/);assert.equal((await book()).sheets[0].cells.A1.raw,'한글');assert.equal((await book()).sheets[1].cells.A1.raw,'한글');
  // Keep using the same dialog after a document replacement transaction.
  await page.locator('#find-text').fill('한글');await scan();assert.match(await page.locator('#find-status').innerText(),/4개 발견/);
  await page.screenshot({path:path.join(root,'test-results/find-replace-preview.png')});await page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click({timeout:500});
  // Noninteractive notices must pass clicks through, but their save/action
  // buttons must remain usable (e.g. an undo-memory-limit warning).
  await page.evaluate(()=>{document.querySelector('.toast')?.remove();const notice=document.createElement('div');notice.className='toast';notice.innerHTML='<button type="button" id="toast-action-check">알림 작업</button>';notice.querySelector('button').onclick=()=>window.__toastActionClicked=true;document.body.append(notice);});await page.locator('#toast-action-check').click({timeout:500});assert.equal(await page.evaluate(()=>window.__toastActionClicked),true);await page.evaluate(()=>document.querySelector('.toast')?.remove());
  await page.keyboard.press('ControlOrMeta+z');await settle();assert.equal((await book()).sheets[0].cells.A1.raw,'ALPHA');assert.equal((await book()).sheets[1].cells.A1.raw,'alpha');
  await page.keyboard.press('ControlOrMeta+y');await settle();assert.equal((await book()).sheets[0].cells.A2.raw,'한글 Beta');
  await page.locator('[data-sheet]').filter({hasText:'매출'}).click();await settle();await go('A1:A2');await open();await page.locator('#find-scope').selectOption('selection');await page.locator('#find-text').fill('한글');await page.locator('#find-whole').check();await scan();assert.match(await page.locator('#find-status').innerText(),/1개 발견/);await close();
  await open();await page.locator('#find-scope').selectOption('sheet');await page.locator('#find-look-in').selectOption('values');await page.locator('#find-text').fill('50%');await scan();assert.match(await page.locator('#find-status').innerText(),/1개 발견/);assert.equal(await page.getByRole('dialog').getByRole('button',{name:'모두 바꾸기',exact:true}).isDisabled(),true);await close();
  // A protected match on another sheet rejects the whole transaction.
  const protectedBook=createWorkbook();protectedBook.sheets[0].cells.A1={raw:'keep',style:{}};protectedBook.sheets.push({id:'protected',name:'보호',cells:{A1:{raw:'keep',style:{}}},charts:[],merges:[],rowHeights:{},columnWidths:{},xlsxFeatures:{sheetProtection:'<sheetProtection sheet="1"/>'}});await load(protectedBook,'find-protected');
  await open();await page.locator('#find-scope').selectOption('workbook');await page.locator('#find-look-in').selectOption('formulas');await page.locator('#find-text').fill('keep');await page.locator('#replace-text').fill('bad');await scan();
  await page.getByRole('dialog').getByRole('button',{name:'모두 바꾸기',exact:true}).click();await settle();assert.match(await page.locator('#find-status').innerText(),/변경하지 않았습니다.*보호/);assert.equal((await book()).sheets[0].cells.A1.raw,'keep');assert.equal((await book()).sheets[1].cells.A1.raw,'keep');await close();
  const file=path.join(dataDir,'find-saved.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');const output=await readSpreadsheet(await fs.readFile(file),'saved.xlsx');assert.equal(output.sheets[0].cells.A1.raw,'keep');
  assert.deepEqual(errors,[]);console.log('PASS: case-insensitive preview/replacement; workbook navigation; one undo/redo; same-dialog reuse; selection/full-cell/value search; protected atomic failure; XLSX saved.');
}catch(error){await page.screenshot({path:path.join(root,'test-results/find-replace-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
