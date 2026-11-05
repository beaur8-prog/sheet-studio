import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','open-hidden-active-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile})),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(15000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
async function check(name){await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,name);await settle();const book=await page.evaluate(()=>sheetStudio.getWorkbook()),active=book.sheets.find(s=>s.id===book.activeSheetId);assert.equal(active.name,'표시 시트');assert.equal(active.cells.A1.raw,'보이는 값');assert.equal(book.sheets[0].state==='hidden'||book.sheets[0].state==='veryHidden',true);assert.equal(book.sheets[0].cells.A1.raw,'숨긴 값');assert.equal(await page.getByRole('gridcell',{name:'A1',exact:true}).innerText(),'보이는 값');}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  for(const state of ['hidden','veryHidden']){
    const book=createWorkbook();book.name='hidden-active-'+state;book.sheets[0].name='숨긴 시트';book.sheets[0].state=state;book.sheets[0].cells={A1:{raw:'숨긴 값'}};book.sheets.push({...createWorkbook().sheets[0],id:'visible',name:'표시 시트',cells:{A1:{raw:'보이는 값'}}});
    await page.locator('#file-input').setInputFiles({name:book.name+'.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await check(book.name);
    const source=new ExcelJS.Workbook();source.views=[{activeTab:0}];const hidden=source.addWorksheet('숨긴 시트',{state}),visible=source.addWorksheet('표시 시트');hidden.getCell('A1').value='숨긴 값';visible.getCell('A1').value='보이는 값';const name='xlsx-hidden-'+state;await page.locator('#file-input').setInputFiles({name:name+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(await source.xlsx.writeBuffer())});await check(name);
  }
  await page.evaluate(()=>sheetStudio.flushRecovery());await page.evaluate(async()=>{const db=await new Promise((resolve,reject)=>{const request=indexedDB.open('sheet-studio-recovery');request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);});const tx=db.transaction('state','readwrite'),store=tx.objectStore('state'),request=store.get('current');request.onsuccess=()=>{const value=request.result;value.metadata.activeSheetId=value.metadata.sheets[0].id;store.put(value,'current');};await new Promise((resolve,reject)=>{tx.oncomplete=resolve;tx.onabort=()=>reject(tx.error);});db.close();});
  await page.reload();await page.locator('.cell').first().waitFor();await settle();await check('xlsx-hidden-veryHidden');assert.deepEqual(errors,[]);console.log('PASS: hidden/veryHidden active sheets choose a visible sheet on JSON/XLSX open and legacy recovery reload, preserving hidden data and visibility');
}catch(error){await page.screenshot({path:path.join(root,'test-results/open-hidden-active-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
