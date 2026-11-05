import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {desktopLaunch} from './electron-launch.mjs';
import {createWorkbook} from '../src/engine.mjs';
import {createTablePlan} from '../src/tables.mjs';
import {chartData} from '../src/chart-model.mjs';
import {validationRules} from '../src/cell-constraints.mjs';

const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','linked-report-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'})),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(20000);await page.setViewportSize({width:1280,height:900});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const data=()=>page.evaluate(()=>sheetStudio.getWorkbook());
const go=async range=>{await page.locator('#name-box').fill(range);await page.locator('#name-box').press('Enter');await settle();};
const edit=async(ref,raw)=>{await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();};
const switchTo=async name=>{await page.locator('[data-sheet]').filter({hasText:name}).click();await settle();};
const previewReady=()=>page.locator('#chart-preview-status').filter({hasText:/개 계열 사용/}).waitFor();
const chartValues=async()=>{const book=await data(),sheet=book.sheets[0];return chartData(sheet.charts[0],sheet,book)[0].values;};
const total=async()=>{const book=await data(),sheet=book.sheets.find(s=>s.pivot);return +sheet.cells['B'+Object.keys(sheet.cells).map(key=>+key.replace(/\D/g,'')).reduce((a,b)=>Math.max(a,b),0)].raw;};
async function save(file){await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');}
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();const book=createWorkbook(),source=book.sheets[0];book.name='Linked report journey';source.name='자료';source.cells={A1:{raw:'지역'},B1:{raw:'금액'},A2:{raw:'동부'},B2:{raw:'10'},A3:{raw:'서부'},B3:{raw:'20'}};
 const plan=createTablePlan(book,source,{name:'Sales',range:'A1:B3'});source.tables=[plan.table];for(const u of plan.updates)source.cells[u.key]=u.cell;book.chartNames=[{name:'PlotValues',formula:'Sales[금액]'},{name:'PlotItems',formula:'Sales[지역]'}];
 await page.locator('#file-input').setInputFiles({name:'linked-report.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Linked report journey');await settle();
 await go('B2:B1048576');await page.locator('[data-tab="데이터"]').click();await page.locator('[data-action="validation"]').click();await page.locator('#dv-type').selectOption('whole');await page.locator('#dv-value').fill('0');await page.locator('#dv-value2').fill('100');await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await settle();assert.ok(validationRules((await data()).sheets[0])[0].ranges[0].r2===1048575);
 await go('B2:B100');await page.locator('[data-tab="홈"]').click();await page.locator('[data-action="conditional"]').click();await page.getByRole('button',{name:'새 규칙',exact:true}).click();await page.locator('#cf-type').selectOption('dataBar');await page.locator('#cf-bar-color').fill('#638ec6');await page.getByRole('dialog').getByRole('button',{name:'확인',exact:true}).click();await page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();await settle();
 await go('A1:B3');await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="chart-insert"]').click();await previewReady();await page.locator('#chart-data-mode').selectOption('series');await page.locator('[data-series-field="name"]').fill('매출');await page.locator('[data-series-field="values"]').fill('PlotValues');await page.locator('[data-series-field="categories"]').fill('PlotItems');await previewReady();await page.getByRole('dialog').getByRole('button',{name:'삽입',exact:true}).click();await settle();assert.deepEqual(await chartValues(),[10,20]);
 await edit('A4','동부');await edit('B4','30');assert.equal((await data()).sheets[0].tables[0].range,'A1:B4');assert.deepEqual(await chartValues(),[10,20,30]);assert.match(await page.getByRole('gridcell',{name:'B4',exact:true}).evaluate(n=>getComputedStyle(n).backgroundImage),/linear-gradient/);
 await go('B2');await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="pivot"]').click();assert.equal(await page.locator('#pivot-source-kind').inputValue(),'table');await page.waitForFunction(()=>!document.querySelector('[data-modal-button="0"]')?.disabled);await page.getByRole('dialog').getByRole('button',{name:'만들기',exact:true}).click();await settle();const report=(await data()).sheets.find(s=>s.pivot).name;assert.equal(await total(),60);
 await switchTo('자료');await edit('A5','서부');await edit('B5','150');assert.equal((await data()).sheets[0].cells.B5?.raw??'','');assert.match(await page.locator('.toast').innerText(),/입력|조건/);await page.locator('#editor').fill('40');await page.locator('#editor').press('Enter');await settle();assert.deepEqual(await chartValues(),[10,20,30,40]);
 await switchTo(report);await page.locator('#pivot-bar').getByRole('button',{name:'새로 고침',exact:true}).click();await settle();assert.equal(await total(),100);await page.locator('#undo').click();await settle();assert.equal(await total(),60);await page.locator('#redo').click();await settle();assert.equal(await total(),100);
 const file=path.join(profile,'linked-report.xlsx');await save(file);await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await total(),100);assert.ok((await data()).sheets.find(s=>s.pivot).pivot.sourceTableId);await switchTo('자료');assert.deepEqual(await chartValues(),[10,20,30,40]);assert.ok((await data()).sheets[0].charts[0].source);assert.equal(validationRules((await data()).sheets[0])[0].type,'whole');assert.match(await page.getByRole('gridcell',{name:'B5',exact:true}).evaluate(n=>getComputedStyle(n).backgroundImage),/linear-gradient/);
 await page.screenshot({path:path.join(root,'test-results/linked-report.png')});await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('.cell').first().waitFor();await settle();assert.deepEqual(await chartValues(),[10,20,30,40]);assert.equal(await total(),100);assert.deepEqual(errors,[]);console.log('PASS: full-column validation, data bars, named table chart, automatic new rows, rejected/retried entry, linked pivot refresh undo/redo, XLSX reopen and recovery.');
}catch(error){await page.screenshot({path:path.join(root,'test-results/linked-report-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
