import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {desktopLaunch} from './electron-launch.mjs';
import {createWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {chartSeriesEditorSeed} from '../src/chart-series-editor.mjs';
import {chartData} from '../src/chart-model.mjs';
const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','chart-conversion-guard-'+Date.now()),baseline=process.argv.includes('--baseline-unguarded');
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'})),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(20000);await page.setViewportSize({width:1000,height:680});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const ready=()=>page.locator('#chart-preview-status').filter({hasText:/개 계열 사용/}).waitFor();
const chart=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].charts[0]);
async function open(){await page.locator('[data-edit-chart]').click();await ready();await page.locator('#chart-convert').check();}
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'항목'},B1:{raw:'금액'},A2:{raw:'동부'},B2:{raw:'10'},A3:{raw:'서부'},B3:{raw:'20'}};
 sheet.charts=[{id:'many',type:'bar',title:'원본 33계열',range:'A1:B3',x:320,y:40,width:600,height:360,series:Array.from({length:33},(_,i)=>({name:'계열 '+(i+1),type:'bar',valRef:{sheetId:sheet.id,range:'B2:B3'},catRef:{sheetId:sheet.id,range:'A2:A3'},color:'#4472c4'}))}];
 await page.locator('#file-input').setInputFiles({name:'many-series.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from((await writeSpreadsheet(book)).data)});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='many-series');await settle();await page.locator('#name-box').fill('A1:B3');await page.locator('#name-box').press('Enter');await settle();const before=await chart();assert.equal(before.series.length,32);await open();
 if(baseline){await ready();assert.equal(await page.locator('[data-series-row]').count(),32);await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await settle();assert.equal((await chart()).source,undefined);assert.equal((await chart()).series.length,32);console.log('REPRODUCED: original 33-series chart converts to the limited 32-series preview without a completeness guard.');}
 else{
  await page.locator('#chart-series-editor .worksheet-error').filter({hasText:'33개 계열'}).waitFor();assert.equal(await page.locator('[data-series-row]').count(),0);await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();assert.match(await page.locator('#chart-error').innerText(),/1~32/);assert.deepEqual(await chart(),before);await page.screenshot({path:path.join(root,'test-results/chart-conversion-guard.png')});
  await page.locator('[data-series-reset]').click();await ready();assert.equal(await page.locator('[data-series-row]').count(),1);await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await settle();assert.equal((await chart()).source,undefined);let current=await page.evaluate(()=>sheetStudio.getWorkbook());assert.deepEqual(chartData(current.sheets[0].charts[0],current.sheets[0],current)[0].values,[10,20]);
  await page.locator('#undo').click();await settle();assert.ok((await chart()).source);const file=path.join(profile,'preserved-33.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');const loaded=await readSpreadsheet(await fs.readFile(file),'preserved.xlsx');assert.match(chartSeriesEditorSeed(loaded.sheets[0].charts[0],loaded.sheets[0],loaded).warning,/33개 계열/);assert.deepEqual(errors,[]);console.log('PASS: partial original preview cannot silently convert; explicit range rebuild works, undo and XLSX save retain all 33 original series.');
 }
}catch(error){await page.screenshot({path:path.join(root,'test-results/chart-conversion-guard-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
