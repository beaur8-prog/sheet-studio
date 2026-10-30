import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import platform from '../scripts/platform.cjs';
import {worksheetExtrasFixture} from './fixtures/worksheet-extras.mjs';
import {readSpreadsheet} from '../src/io.mjs';
import {printOrientation} from '../src/worksheet-extra.mjs';
const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','worksheet-extras-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch(desktopLaunch(root,env)),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(15000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));};
const model=()=>page.evaluate(()=>sheetStudio.getWorkbook()),act=async action=>{await page.locator('[data-action="'+action+'"]').first().click();await settle();};
try{
  await page.locator('.cell').first().waitFor();await page.locator('#file-input').setInputFiles({name:'print-original.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(await worksheetExtrasFixture())});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='print-original');await settle();assert.equal(printOrientation((await model()).sheets[0]),'landscape');
  await page.locator('[data-tab="페이지 레이아웃"]').click();await act('portrait');assert.equal(printOrientation((await model()).sheets[0]),'portrait');await page.locator('#undo').click();await settle();assert.equal(printOrientation((await model()).sheets[0]),'landscape');await page.locator('#redo').click();await settle();assert.equal(printOrientation((await model()).sheets[0]),'portrait');
  await act('print');assert.equal(await page.locator('#print-orientation').inputValue(),'portrait');await page.getByRole('button',{name:'취소',exact:true}).click();await page.locator('#name-box').fill('A2');await page.locator('#name-box').press('Enter');await page.keyboard.press('F2');await page.locator('#editor').fill('4');await page.locator('#editor').press('Enter');await settle();
  const output=path.join(dataDir,'print-saved.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');const saved=await readSpreadsheet(await fs.readFile(output),'saved.xlsx');assert.equal(printOrientation(saved.sheets[0]),'portrait');assert.equal(saved.xlsxProperties.creator,'Original Author');assert.equal(saved.xlsxProperties.created,'2018-04-03T02:01:00.000Z');assert.equal(saved.sheets[0].xlsxFeatures.linkedParts.rels.length,2);assert.match(saved.sheets[0].xlsxFeatures.headerFooter,/Confidential/);assert.match(saved.sheets[0].xlsxFeatures.sheetPr,/OriginalCode/);
  await page.locator('#file-input').setInputFiles(output);await settle();assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('B2')),12);await page.locator('[data-tab="페이지 레이아웃"]').click();await act('print');assert.equal(await page.locator('#print-orientation').inputValue(),'portrait');await page.getByRole('button',{name:'취소',exact:true}).click();assert.deepEqual(errors,[]);
  console.log('PASS: actual original print settings, persisted orientation/undo/redo and print-dialog default, edit/save/reopen with headers, printer/header-image relationships and original author/creation date');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/worksheet-extras-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
