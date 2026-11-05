import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import {desktopLaunch} from './electron-launch.mjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','chart-blanks-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'})),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(20000);await page.setViewportSize({width:1480,height:960});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const model=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].charts[0]);
const lengths=(root='.chart-card')=>page.locator(root+' .chart-data polyline').evaluateAll(nodes=>nodes.map(n=>n.getAttribute('points').trim().split(/\s+/).length));
async function open(){await page.locator('[data-edit-chart]').click();await ready();}
async function ready(){await page.locator('#chart-preview-status').filter({hasText:'개 계열 사용'}).waitFor();}
async function apply(){await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await page.getByRole('dialog').waitFor({state:'detached'});await settle();}
async function save(name){const file=path.join(profile,name+'.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');return file;}
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();const book=createWorkbook(),s=book.sheets[0];book.name='Chart blanks';s.cells={A1:{raw:'항목'},B1:{raw:'매출'}};['갑','을','병','정','무'].forEach((v,i)=>s.cells['A'+(i+2)]={raw:v});[10,null,30,'=NA()',50].forEach((v,i)=>{if(v!==null)s.cells['B'+(i+2)]={raw:String(v)};});s.charts=[{id:'line',type:'line',range:'A1:B6',title:'빈 셀 처리',width:600,height:360,x:300,y:30,firstRowAsHeader:true,firstColumnAsLabels:true,seriesBy:'columns'}];
 await page.locator('#file-input').setInputFiles({name:'blanks.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Chart blanks');await settle();assert.deepEqual(await lengths(),[1,1,1]);await open();assert.equal(await page.locator('#chart-blanks').inputValue(),'gap');await page.locator('#chart-blanks').selectOption('zero');await ready();assert.equal(await page.locator('#chart-preview-image .chart-point').count(),4);assert.match(await page.locator('#chart-preview-data').innerText(),/0 \(빈 셀\)/);assert.match(await page.locator('#chart-preview-data').innerText(),/숫자 아님 \/ 오류/);await apply();assert.deepEqual(await lengths(),[3,1]);await page.locator('#undo').click();await settle();assert.deepEqual(await lengths(),[1,1,1]);await page.locator('#redo').click();await settle();assert.equal((await model()).blankCells,'zero');
 await open();await page.locator('#chart-blanks').selectOption('span');await ready();assert.deepEqual(await lengths('#chart-preview-image'),[2,1]);await page.screenshot({path:path.join(root,'test-results/chart-blanks-span.png')});await apply();assert.deepEqual(await lengths(),[2,1]);
 await go('B3');await page.keyboard.press('F2');await page.locator('#editor').fill('=""');await page.locator('#editor').press('Enter');await settle();assert.deepEqual(await lengths(),[1,1,1]);await page.locator('#undo').click();await settle();assert.deepEqual(await lengths(),[2,1]);
 const file=await save('span');await page.locator('#file-input').setInputFiles(file);await settle();assert.equal((await model()).blankCells,'span');assert.ok((await model()).source);await open();assert.equal(await page.locator('#chart-blanks').inputValue(),'span');await page.locator('#chart-blanks').selectOption('zero');await ready();await apply();assert.ok((await model()).source);assert.deepEqual(await lengths(),[3,1]);const native=await save('native-zero'),zip=await JSZip.loadAsync(await fs.readFile(native)),part=zip.file(/(?:Chart|chart)\d+\.xml$/).find(p=>p.name.includes('chart'));assert.match(await part.async('string'),/dispBlanksAs[^>]*val="zero"/);
 await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('.chart-card').waitFor();await settle();assert.equal((await model()).blankCells,'zero');assert.deepEqual(await lengths(),[3,1]);assert.deepEqual(errors,[]);
 console.log('PASS: physical gap/zero/span in live preview and chart; errors remain gaps; formula-empty distinction; undo/redo; native XLSX policy edit without conversion; recovery.');
}catch(error){await page.screenshot({path:path.join(root,'test-results/chart-blanks-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
