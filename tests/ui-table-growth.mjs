import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {desktopLaunch} from './electron-launch.mjs';
import {createWorkbook} from '../src/engine.mjs';
import {createTablePlan,calculatedColumnPlan} from '../src/tables.mjs';
import {readSpreadsheet} from '../src/io.mjs';

const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','table-growth-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'})),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(20000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const edit=async(ref,raw)=>{await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();};
const model=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets[0]),value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
async function manage(){await go('B2');await page.locator('#table-bar').getByRole('button',{name:'표 관리',exact:true}).click();await page.getByRole('dialog',{name:'표 관리',exact:true}).waitFor();}
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();const book=createWorkbook(),s=book.sheets[0];book.name='Table growth';s.cells={A1:{raw:'Qty'},B1:{raw:'Price'},C1:{raw:'Amount'},A2:{raw:'2'},B2:{raw:'3'},A3:{raw:'4'},B3:{raw:'5'},D2:{raw:'10',note:'기존 자료'},D3:{raw:'20'},F1:{raw:'보너스'},F2:{raw:'7'},F3:{raw:'8'},H1:{raw:'병합 유지'},J1:{raw:'=COLUMNS(Sales[#Headers])'}};s.merges=['H1:I1'];
 const p=createTablePlan(book,s,{range:'A1:C3',name:'Sales'});s.tables=[p.table];for(const u of p.updates)s.cells[u.key]=u.cell;const calc=calculatedColumnPlan(book,s,s.tables[0],{column:2,formula:'=[@Qty]*[@Price]'});s.tables=[calc.table];for(const u of calc.updates)s.cells[u.key]=u.cell;const ids=s.tables[0].columns.map(c=>c.id);
 await page.locator('#file-input').setInputFiles({name:'table-growth.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Table growth');await settle();assert.equal(await value('J1'),3);
 await edit('D1','Qty');assert.equal((await model()).tables[0].range,'A1:D3');assert.equal((await model()).tables[0].columns[3].name,'Qty2');assert.equal((await model()).cells.D2.note,'기존 자료');assert.equal(await value('J1'),4);await page.locator('#undo').click();await settle();assert.equal((await model()).tables[0].range,'A1:C3');assert.equal(await value('D2'),10);await page.locator('#redo').click();await settle();assert.equal(await value('J1'),4);
 await edit('E3','=[@Qty]+[@Price]');assert.equal((await model()).tables[0].range,'A1:E3');assert.equal(await value('E2'),5);assert.equal(await value('E3'),9);assert.equal((await model()).tables[0].columns[4].name,'열5');
 await manage();await page.locator('#table-range').fill('A1:F3');await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await page.getByRole('dialog').waitFor({state:'detached'});await settle();assert.equal((await model()).tables[0].range,'A1:F3');assert.equal(await value('J1'),6);assert.equal(await value('F2'),7);assert.deepEqual((await model()).tables[0].columns.slice(0,3).map(c=>c.id),ids);
 await manage();await page.locator('#table-range').fill('A1:H3');await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await page.locator('#table-error').filter({hasText:'병합'}).waitFor();assert.equal((await model()).tables[0].range,'A1:F3');await page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();
 await go('F1');await page.locator('[data-filter-column="5"]').click();await page.locator('#filter-all').uncheck();await page.getByRole('checkbox',{name:'7',exact:true}).check();await page.locator('#filter-ok').click();await settle();assert.deepEqual((await model()).tables[0].filter.hiddenRows,[2]);
 const file=path.join(profile,'expanded.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');const read=await readSpreadsheet(await fs.readFile(file),'expanded.xlsx');assert.equal(read.sheets[0].tables[0].range,'A1:F3');assert.equal(read.sheets[0].cells.D2.note,'기존 자료');assert.deepEqual(read.sheets[0].tables[0].columns.map(c=>c.id),[1,2,3,4,5,6]);assert.deepEqual(read.sheets[0].tables[0].filter.hiddenRows,[2]);
 await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('J1'),6);await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('.cell').first().waitFor();assert.equal((await model()).tables[0].range,'A1:F3');assert.equal(await value('E3'),9);await page.screenshot({path:path.join(root,'test-results/table-growth.png')});assert.deepEqual(errors,[]);
 console.log('PASS: adjacent header growth; stable IDs; calculated new column; explicit right resize; outside data/notes preserved; conflict unchanged; new-column filter; undo/redo; XLSX and recovery.');
}catch(error){await page.screenshot({path:path.join(root,'test-results/table-growth-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
