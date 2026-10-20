import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook,address} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const dataDir=path.join(root,'.package','chart-selection-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(8000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const get=()=>page.evaluate(()=>sheetStudio.getWorkbook());
const bars=()=>page.locator('.chart-card>svg rect:has(title)');
const titles=()=>bars().locator('title').allTextContents();
let fixtureId=0;
const load=async rows=>{const b=createWorkbook();b.name='chart-selection '+ ++fixtureId;rows.forEach((row,r)=>row.forEach((v,c)=>b.sheets[0].cells[address(r,c)]={raw:String(v)}));await page.locator('#file-input').setInputFiles({name:'selection.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(n=>sheetStudio.getWorkbook().name===n,b.name);await settle();};
const go=async range=>{await page.locator('#name-box').fill(range);await page.locator('#name-box').press('Enter');await settle();};
const insert=async(range,{drag=false,reverse=false,action='chart-bar'}={})=>{
 if(drag){const [first,last]=range.split(':'),refs=reverse?[last,first]:[first,last],points=await page.evaluate(refs=>refs.map(ref=>{const q=document.querySelector(`.cell[aria-label="${ref}"]`).getBoundingClientRect();return{x:q.x+q.width/2,y:q.y+q.height/2};}),refs);await page.mouse.move(points[0].x,points[0].y);await page.mouse.down();await page.mouse.move(points[1].x,points[1].y,{steps:10});await page.mouse.up();await settle();assert.equal(await page.locator('#name-box').inputValue(),range);}
 else await go(range);
 await page.locator('[data-tab="삽입"]').click();await page.locator(`[data-action=${action}]`).click();await page.locator('.chart-card').waitFor();await settle();assert.equal((await get()).sheets[0].charts[0].range,range);
};
const settings=()=>page.locator('[data-edit-chart]').click();
const apply=async()=>{await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await page.locator('.backdrop').waitFor({state:'detached'});await settle();};
try{
 await page.locator('.cell').first().waitFor();await fs.mkdir('test-results',{recursive:true});
 await load([[10,20,30],[40,50,60],[70,80,90]]);await insert('A1:C3',{drag:true});assert.equal(await bars().count(),9);assert.ok((await titles()).includes('계열 1 · 1: 10'));
 let chart=(await get()).sheets[0].charts[0];assert.equal(chart.firstRowAsHeader,false);assert.equal(chart.firstColumnAsLabels,false);
 await load([['매출','비용','수량'],[40,22,8],[25,18,12],[35,10,15]]);await insert('A1:C4',{drag:true,reverse:true});assert.equal(await bars().count(),9);assert.ok((await titles()).includes('매출 · 1: 40'));
 await go('A2');await page.keyboard.press('F2');await page.locator('#editor').fill('99');await page.locator('#editor').press('Enter');await settle();assert.ok((await titles()).includes('매출 · 1: 99'));
 // Save and reopen through the UI; generated XLSX must contain all three series.
 const output=path.join(dataDir,'selected.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');await page.locator('#file-input').setInputFiles(output);await page.waitForFunction(()=>!!sheetStudio.getWorkbook().sheets[0].charts[0]?.source);await settle();assert.equal((await get()).sheets[0].charts[0].series.length,3);assert.equal(await bars().count(),9);assert.ok((await titles()).includes('매출 · 1: 99'));
 await settings();assert.equal(await page.locator('#chart-first-row').isDisabled(),true);assert.equal(await page.locator('#chart-first-column').isDisabled(),true);await page.getByRole('button',{name:'닫기',exact:true}).click();
 await load([['상품','매출','비용'],['커피',40,22],['차',25,18],['주스',35,10]]);await insert('A1:C4',{drag:true});assert.equal(await bars().count(),6);assert.ok((await titles()).includes('매출 · 커피: 40'));
 await load([['커피',40,22],['차',25,18]]);await insert('A1:C2');assert.equal(await bars().count(),4);assert.ok((await titles()).includes('계열 1 · 커피: 40'));
 await load([['매출'],[10],[20],[30]]);await insert('A1:A4',{drag:true});assert.equal(await bars().count(),3);
 await settings();await page.locator('#chart-first-column').check();await page.getByRole('button',{name:'적용',exact:true}).click();assert.match(await page.locator('#chart-error').innerText(),/값/);await page.getByRole('button',{name:'닫기',exact:true}).click();
 // Numeric years/IDs can deliberately serve as labels. Undo, cancel and reload
 // must retain that choice rather than infer a different layout on each edit.
 await load([['연도','매출'],[2024,10],[2025,20]]);await insert('A1:B3');assert.equal(await bars().count(),4);
 await settings();assert.equal(await page.locator('#chart-first-column').isChecked(),false);await page.locator('#chart-first-column').check();await apply();assert.equal(await bars().count(),2);assert.ok((await titles()).includes('매출 · 2024: 10'));
 await page.locator('#undo').click();await settle();assert.equal(await bars().count(),4);await page.locator('#redo').click();await settle();assert.equal(await bars().count(),2);
 await settings();await page.locator('#chart-first-column').uncheck();await page.getByRole('button',{name:'닫기',exact:true}).click();assert.equal((await get()).sheets[0].charts[0].firstColumnAsLabels,true);
 await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('.chart-card').waitFor();await settle();assert.equal(await bars().count(),2);assert.ok((await titles()).includes('매출 · 2024: 10'));
 await page.screenshot({path:'test-results/chart-selection-labels.png'});
 // Numeric headings are supported by opting in to the first-row setting.
 await load([[2024,2025],[10,20],[30,40]]);await insert('A1:B3');assert.equal(await bars().count(),6);await settings();await page.locator('#chart-first-row').check();await apply();assert.equal(await bars().count(),4);assert.ok((await titles()).includes('2024 · 1: 10'));
 // The insert dialog and line shortcut use the same source-range interpretation.
 await load([[10,20],[30,40]]);await go('A1:B2');await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action=chart-insert]').click();assert.equal(await page.locator('#chart-first-column').isChecked(),false);assert.equal(await page.locator('#chart-first-row').isChecked(),false);await page.getByRole('dialog').getByRole('button',{name:'삽입',exact:true}).click();await settle();assert.equal(await bars().count(),4);
 await load([[10,20],[30,40]]);await insert('A1:B2',{action:'chart-line'});assert.equal(await page.locator('.chart-card>svg polyline').count(),2);assert.equal(await page.locator('.chart-card>svg circle').count(),4);
 assert.deepEqual(errors,[]);console.log('PASS: drag/reverse selection, first numeric column and row, text headings/categories, single column, label settings, live edit, undo/cancel/reload, XLSX save/reopen and line charts');
}finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
