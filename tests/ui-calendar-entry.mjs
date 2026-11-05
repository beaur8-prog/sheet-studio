import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createWorkbook,evaluateFormula} from '../src/engine.mjs';
import {applyValidationChange,planValidationChange} from '../src/data-validation.mjs';
const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','calendar-entry-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile})),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(15000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const edit=async(ref,text)=>{await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(text);await page.locator('#editor').press('Enter');await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  for(const date1904 of [false,true]){
    const book=createWorkbook(),sheet=book.sheets[0];book.name='calendar-'+date1904;book.date1904=date1904;
    sheet.cells={A1:{raw:'',style:{format:'date',numberFormat:'yyyy-mm-dd'}},B1:{raw:'',style:{format:'date',numberFormat:'hh:mm:ss'}},C1:{raw:'',style:{format:'text'}},G1:{raw:'',style:{format:'general'}},H1:{raw:'=ISNUMBER(A1)'},I1:{raw:'=HOUR(B1)'}};
    applyValidationChange(sheet,planValidationChange(book,sheet,'D1:D5',{type:'date',formulae:['2026-01-01','2026-12-31']}));applyValidationChange(sheet,planValidationChange(book,sheet,'E1:E5',{type:'time',formulae:['09:00','18:00']}));
    await page.locator('#file-input').setInputFiles({name:book.name+'.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,book.name);await settle();
    await edit('A1','2026-11-05');assert.equal(await value('A1'),evaluateFormula(book,sheet.id,'=DATE(2026,11,5)'));
    await edit('B1','13:45:30');assert.ok(Math.abs(await value('B1')-(13*3600+45*60+30)/86400)<1e-12);assert.equal(await value('H1'),true);assert.equal(await value('I1'),13);
    await edit('C1','2026-11-05');assert.equal(await value('C1'),'2026-11-05');await edit('G1','2026-11-05');assert.equal(await value('G1'),'2026-11-05');
    await edit('D1','2026-06-15');assert.equal(await value('D1'),evaluateFormula(book,sheet.id,'=DATE(2026,6,15)'));await edit('E1','12:30');assert.equal(await value('E1'),12.5/24);
    await edit('D2','2026-02-30');assert.equal(await value('D2'),'');await page.keyboard.press('Escape');await edit('E2','25:00');assert.equal(await value('E2'),'');await page.keyboard.press('Escape');
    const output=path.join(profile,'calendar-'+date1904+'.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');await page.locator('#file-input').setInputFiles(output);await settle();assert.equal(await value('D1'),evaluateFormula(book,sheet.id,'=DATE(2026,6,15)'));assert.equal(await value('E1'),12.5/24);assert.equal(await value('C1'),'2026-11-05');
  }
  assert.deepEqual(errors,[]);await page.screenshot({path:path.join(root,'test-results/calendar-entry.png')});console.log('PASS: deterministic ISO date/time entry in formatted or validated cells, date1904, invalid input, General/Text identity and XLSX reopen');
}catch(error){await page.screenshot({path:path.join(root,'test-results/calendar-entry-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
