import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','let-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix()),args:[root],env}),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(15000);
await app.evaluate(({ipcMain})=>{ipcMain.removeHandler('sheet:write-clipboard');ipcMain.handle('sheet:write-clipboard',(_event,payload)=>{globalThis.testTypedClipboard=payload;});});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();},value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref),raw=ref=>page.evaluate(ref=>sheetStudio.getWorkbook().sheets[0].cells[ref]?.raw,ref);
const edit=async(ref,text)=>{await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(text);await page.locator('#editor').press('Enter');await settle();};
const copy=async()=>{await page.locator('[data-action="copy"]').click();await settle();return app.evaluate(()=>globalThis.testTypedClipboard.text);};
const paste=async text=>{await page.locator('#grid').focus();await page.evaluate(text=>{const data=new DataTransfer();data.setData('text/plain',text);document.querySelector('#grid').dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));},text);await settle();};
try{
 await page.locator('.cell').first().waitFor();const b=createWorkbook(),s=b.sheets[0];b.name='let-ui';s.cells.A1={raw:'10'};s.cells.A2={raw:'20'};
 await page.locator('#file-input').setInputFiles({name:'let-ui.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await settle();
 await edit('B1','=LET(data,A1:A2,total,SUM(data),total/COUNT(data))');assert.equal(await value('B1'),15);assert.equal(await page.getByRole('gridcell',{name:'B1',exact:true}).textContent(),'15');
 await edit('A2','30');assert.equal(await value('B1'),20);await page.locator('#undo').click();await settle();assert.equal(await value('B1'),15);await page.locator('#redo').click();await settle();assert.equal(await value('B1'),20);
 await edit('B2','=LET(x,5,LET(x,x+1,x)+x)');assert.equal(await value('B2'),11);await edit('B3','=LET(x,RAND(),x-x)');assert.equal(await value('B3'),0);await page.keyboard.press('F9');await settle();assert.equal(await value('B3'),0);
 await go('B1');const text=await copy();await go('C1');await paste(text);assert.equal(await value('C1'),15.5);
 const output=path.join(dataDir,'let.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');const out=new ExcelJS.Workbook();await out.xlsx.load(await fs.readFile(output));assert.match(out.worksheets[0].getCell('B1').formula,/^_xlfn.LET/);assert.equal(out.worksheets[0].getCell('B1').result,20);assert.equal(out.worksheets[0].getCell('C1').result,15.5);await page.locator('#file-input').setInputFiles(output);await settle();assert.equal(await value('B1'),20);assert.equal(await value('B2'),11);assert.equal(await value('B3'),0);assert.equal(await value('C1'),15.5);assert.deepEqual(errors,[]);
 console.log('PASS: actual LET entry, lexical scope, dependency recalculation, undo/redo, volatile reuse/F9, relative copy, worker XLSX caches/future prefix and reopen');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/let-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
