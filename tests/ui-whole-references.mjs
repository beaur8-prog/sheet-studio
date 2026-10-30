import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','whole-references-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix()),args:[root],env}),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(20000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const edit=async(ref,raw)=>{await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
try{
  await page.locator('.cell').first().waitFor();const book=createWorkbook();book.name='whole-ui';book.sheets[0].cells={A1:{raw:'2'},A3:{raw:'4'},A1048576:{raw:'8'},B1:{raw:'x'},B3:{raw:'y'},B1048576:{raw:'last'}};await page.locator('#file-input').setInputFiles({name:'whole-ui.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='whole-ui');await settle();
  for(const [ref,formula,wanted] of [['D1','=SUM(A:A)',14],['D2','=COUNTBLANK(A:A)',1048573],['D3','=SUMIF(B:B,"x",A:A)',2],['D4','=XLOOKUP(8,A:A,B:B)','last'],['D5','=INDEX(A:A,1048576)',8],['D6','=SUM(INDIRECT("A:A"))',14]]){await edit(ref,formula);assert.equal(await value(ref),wanted,formula);}
  await go('D1');await page.keyboard.press('F2');await page.locator('#editor').evaluate(el=>el.setSelectionRange(7,7));await page.locator('#editor').press('F4');assert.equal(await page.locator('#editor').inputValue(),'=SUM($A:$A)');assert.ok(await page.locator('[data-reference="$A:$A"]').count());await page.locator('#editor').press('Enter');await settle();assert.equal(await value('D1'),14);
  await edit('A1048576','10');await go('D1');assert.equal(await value('D1'),16);assert.equal(await value('D5'),10);await page.locator('#undo').click();await settle();assert.equal(await value('D1'),14);await page.locator('#redo').click();await settle();assert.equal(await value('D1'),16);
  await go('A1');await page.locator('[data-action="insert-col"]').click();await settle();assert.equal(await value('E1'),16);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.E1.raw),'=SUM($B:$B)');await page.locator('#undo').click();await settle();assert.equal(await value('D1'),16);
  const output=path.join(dataDir,'whole.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');await page.locator('#file-input').setInputFiles(output);await settle();assert.equal(await value('D1'),16);assert.equal(await value('D5'),10);assert.equal(await value('D2'),1048573);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.A1048576.raw),'10');assert.deepEqual(errors,[]);
  console.log('PASS: actual whole-column aggregates/criteria/lookup/INDEX/INDIRECT, F4 and reference highlights, last-row editing, undo/redo, column insertion and XLSX reopen');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/whole-reference-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
