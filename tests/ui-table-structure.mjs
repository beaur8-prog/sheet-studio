import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import platform from '../scripts/platform.cjs';
import {tableFixture} from './tables-fixture.mjs';
const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','table-structure-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix()),args:[root],env}),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(15000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const act=async name=>{await page.locator(`[data-action="${name}"]`).first().click();await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref),table=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].tables[0]);
const edit=async(ref,raw)=>{await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();};
try{
  await page.locator('.cell').first().waitFor();await page.locator('#file-input').setInputFiles({name:'structure.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(await tableFixture())});await settle();await page.locator('[data-tab="홈"]').click();
  await go('A2');await act('insert-row');assert.equal((await table()).range,'A1:C5');assert.equal(await value('C5'),26);await edit('A2','3');await edit('B2','4');assert.equal(await value('C2'),12);assert.equal(await value('C5'),38);
  await go('B2');await act('insert-col');assert.deepEqual((await table()).columns.map(c=>c.name),['Qty','열2','Price','Amount']);assert.equal(await value('D5'),38);await page.locator('#undo').click();await settle();assert.equal((await table()).range,'A1:C5');assert.equal(await value('C5'),38);await page.locator('#redo').click();await settle();assert.equal((await table()).range,'A1:D5');
  await go('A1');await act('delete-menu');await act('delete-col');assert.equal(await value('E1'),'#REF!');await page.locator('#undo').click();await settle();assert.equal(await value('F1'),38);
  await go('A1');await act('delete-menu');await act('delete-row');assert.equal((await table()).range,'A1:D5');assert.match(await page.locator('.toast').textContent(),/머리글/);
  await act('table-style');await page.getByRole('dialog').getByRole('button',{name:'일반 범위로 변환',exact:true}).click();await page.getByRole('dialog').waitFor({state:'detached'});await settle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].tables.length),0);assert.equal(await value('D2'),12);assert.equal(await value('D5'),38);assert.equal(await value('F1'),38);assert.equal(await page.getByRole('gridcell',{name:'A1',exact:true}).evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(68, 114, 196)');
  await page.locator('#undo').click();await settle();assert.equal((await table()).name,'Sales');await page.locator('#redo').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].tables.length),0);
  const output=path.join(dataDir,'converted.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await act('save');await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');await page.locator('#file-input').setInputFiles(output);await settle();assert.equal(await value('F1'),38);assert.equal(await value('D5'),38);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].tables?.length||0),0);
  assert.deepEqual(errors,[]);console.log('PASS: actual table row/column insertion, calculated-row inheritance, deleted-column errors, header protection, undo/redo, conversion retaining styles/references and XLSX reopen');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/table-structure-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
