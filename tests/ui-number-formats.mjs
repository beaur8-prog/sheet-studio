import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import {createWorkbook} from '../src/engine.mjs';
import platform from '../scripts/platform.cjs';
const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','number-formats-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix()),args:[root],env}),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(15000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();},cell=ref=>page.getByRole('gridcell',{name:ref,exact:true});
const openFormat=async()=>{await page.locator('#grid').focus();await page.keyboard.press(process.platform==='darwin'?'Meta+1':'Control+1');await page.getByRole('dialog',{name:'셀 서식',exact:true}).waitFor();await page.getByRole('tab',{name:'표시 형식',exact:true}).click();};
const apply=async()=>{await page.getByRole('button',{name:'적용',exact:true}).click();await page.getByRole('dialog').waitFor({state:'detached'});await settle();};
try{
  await page.locator('.cell').first().waitFor();const b=createWorkbook();b.name='number-format-ui';b.sheets[0].cells={A1:{raw:'12'},A2:{raw:'-1234.5'},A3:{raw:'0'},A4:{raw:'text'},B1:{raw:'=TEXT(A2,"#,##0.00;(#,##0.00)")'},B2:{raw:'5.25',style:{numberFormat:'# ?/?',format:'number'}}};
  await page.locator('#file-input').setInputFiles({name:'number-format.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='number-format-ui');await settle();
  await go('A1');await openFormat();await page.locator('#format-format').selectOption('custom');await page.locator('#format-numberFormat').fill('00000');assert.equal(await page.locator('#format-sample').textContent(),'00012');await apply();assert.equal(await cell('A1').textContent(),'00012');assert.equal(await page.locator('#formula').inputValue(),'12');
  await page.locator('#undo').click();await settle();assert.equal(await cell('A1').textContent(),'12');await page.locator('#redo').click();await settle();assert.equal(await cell('A1').textContent(),'00012');
  await go('A2:A4');await openFormat();await page.locator('#format-format').selectOption('custom');await page.locator('#format-numberFormat').fill('#,##0.00;[Red](#,##0.00);"-";"item: "@');assert.equal(await page.locator('#format-sample').textContent(),'(1,234.50)');assert.equal(await page.locator('#format-sample').evaluate(el=>getComputedStyle(el).color),'rgb(255, 0, 0)');await apply();
  assert.equal(await cell('A2').textContent(),'(1,234.50)');assert.equal(await cell('A2').evaluate(el=>getComputedStyle(el).color),'rgb(255, 0, 0)');assert.equal(await cell('A3').textContent(),'-');assert.equal(await cell('A4').textContent(),'item: text');assert.equal(await cell('B1').textContent(),'(1,234.50)');assert.equal(await cell('B2').textContent(),'5 1/4');
  await go('A1');await openFormat();await page.locator('#format-format').selectOption('custom');await page.locator('#format-numberFormat').fill('0"broken');await page.getByRole('button',{name:'적용',exact:true}).click();assert.match(await page.locator('#format-error').textContent(),/큰따옴표/);await page.getByRole('button',{name:'취소',exact:true}).click();assert.equal(await cell('A1').textContent(),'00012');
  const output=path.join(dataDir,'formatted.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');
  const excel=new ExcelJS.Workbook();await excel.xlsx.load(await fs.readFile(output));assert.equal(excel.worksheets[0].getCell('A1').value,12);assert.equal(excel.worksheets[0].getCell('A2').value,-1234.5);assert.equal(excel.worksheets[0].getCell('B1').result,'(1,234.50)');
  await page.locator('#file-input').setInputFiles(output);await settle();assert.equal(await cell('A1').textContent(),'00012');assert.equal(await cell('A2').textContent(),'(1,234.50)');assert.equal(await cell('A4').textContent(),'item: text');await page.screenshot({path:path.join(root,'test-results/number-formats.png')});assert.deepEqual(errors,[]);
  console.log('PASS: actual custom-code entry and sample, numeric raw value, section/color/text display, fractions and TEXT, validation/cancel, undo/redo, worker XLSX save and reopen');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/number-formats-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
