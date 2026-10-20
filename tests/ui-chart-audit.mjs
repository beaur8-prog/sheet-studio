import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook,address} from '../src/engine.mjs';
import {chartCases} from './fixtures/charts.mjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const dataDir=path.join(root,'.package','chart-audit-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(8000);await page.setViewportSize({width:1280,height:900});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const get=()=>page.evaluate(()=>sheetStudio.getWorkbook());
const go=async range=>{await page.locator('#name-box').fill(range);await page.locator('#name-box').press('Enter');await settle();};
let fixture=0;
const load=async(rows=[],charts=[])=>{const b=createWorkbook();b.name='chart audit '+ ++fixture;rows.forEach((row,r)=>row.forEach((value,c)=>{if(value!==null)b.sheets[0].cells[address(r,c)]={raw:String(value)};}));b.sheets[0].charts=charts;await page.locator('#file-input').setInputFiles({name:'chart-audit.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(n=>sheetStudio.getWorkbook().name===n,b.name);await settle();};
const open=async(range,type,{seriesBy='auto',grouping='clustered',drag=false}={})=>{
 if(drag){const [first,last]=range.split(':'),points=await page.evaluate(refs=>refs.map(ref=>{const r=document.querySelector(`.cell[aria-label="${ref}"]`).getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2};}),[first,last]);await page.mouse.move(points[0].x,points[0].y);await page.mouse.down();await page.mouse.move(points[1].x,points[1].y,{steps:6});await page.mouse.up();await settle();assert.equal(await page.locator('#name-box').inputValue(),range);}
 else await go(range);
 await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="chart-insert"]').click();await page.locator('#chart-type').selectOption(type);await page.locator('#chart-series-by').selectOption(seriesBy);await page.locator('#chart-grouping').selectOption(grouping);
};
const submit=async name=>{await page.getByRole('dialog').getByRole('button',{name,exact:true}).click();await settle();};
const check=async count=>{await page.waitForFunction(n=>document.querySelectorAll('.chart-card>svg .chart-point').length===n,count);await settle();assert.equal(await page.locator('.chart-card>svg .chart-point').count(),count);const invalid=await page.locator('.chart-card>svg').evaluate(svg=>[...svg.querySelectorAll('.chart-point')].some(p=>{const b=p.getBBox();return !Number.isFinite(b.x+b.y+b.width+b.height)||b.width<=0||b.height<=0;}));assert.equal(invalid,false);assert.doesNotMatch(await page.locator('.chart-card>svg').innerHTML(),/NaN|Infinity/);};
const roundtrip=async()=>{const output=path.join(dataDir,'horizontal.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');await page.locator('#file-input').setInputFiles(output);await page.waitForFunction(()=>!!sheetStudio.getWorkbook().sheets[0].charts[0]?.source);await settle();};
try{
 await page.locator('.cell').first().waitFor();await fs.mkdir('test-results',{recursive:true});
 // Reproduce the user's exact four-cell row through native typing and drag selection.
 await load();await go('A1');for(const value of ['10','20','30','40']){await page.keyboard.type(value);await page.keyboard.press('Tab');}
 await open('A1:D1','pie',{drag:true});await submit('삽입');await check(4);assert.equal((await get()).sheets[0].charts[0].seriesBy,'rows');assert.deepEqual(await page.locator('.chart-point title').allTextContents(),['1: 10','2: 20','3: 30','4: 40']);await page.screenshot({path:'test-results/chart-audit-pie.png'});
 await roundtrip();await check(4);await page.locator('[data-edit-chart]').click();assert.equal(await page.locator('#chart-series-by').isDisabled(),true);await page.getByRole('button',{name:'닫기',exact:true}).click();
 // Every menu type and grouping uses four selected values. Bubble gets complete X/Y/size data.
 for(const spec of chartCases){
  const rows=spec.type==='bubble'?[[1,10,5],[2,20,10],[3,30,15],[4,40,20]]:[[10,20,30,40]],range=spec.type==='bubble'?'A1:C4':'A1:D1';
  await load(rows);await open(range,spec.type,{grouping:spec.grouping,drag:true});await submit('삽입');await check(4);
  if(spec.grouping==='clustered'){await page.screenshot({path:`test-results/chart-audit-${spec.type}.png`});if(['doughnut','scatter','bubble'].includes(spec.type)){await roundtrip();await check(4);}}
 }
 // A second data orientation must keep all four points too.
 for(const type of ['bar','bar-horizontal','line','area','pie','doughnut','radar','combo','scatter']){await load([[10],[20],[30],[40]]);await open('A1:A4',type);await submit('삽입');await check(4);}
 // A missing size vector gives an actionable error, rather than drawing one or two arbitrary bubbles.
 await load([[10,20,30,40]]);await open('A1:D1','bubble');assert.match(await page.locator('#chart-data-help').innerText(),/X값.*Y값·크기/);await submit('삽입');assert.match(await page.locator('#chart-error').innerText(),/X값·Y값·크기/);assert.equal((await get()).sheets[0].charts.length,0);await page.getByRole('button',{name:'닫기',exact:true}).click();
 await load([[1,2,3,4],[10,20,30,40],[5,10,15,20]]);await open('A1:D3','bubble',{seriesBy:'rows'});await submit('삽입');await check(4);
 // Explicit row/column switching, names/categories, multiple rings, cancellation, undo and restart.
 await load([['항목','가','나','다','라'],['매출',10,20,30,40],['비용',5,8,12,15]]);await open('A1:E3','pie',{seriesBy:'rows'});assert.match(await page.locator('#chart-data-help').innerText(),/2개 계열.*첫 계열/);assert.equal(await page.locator('#chart-first-row').isChecked(),true);assert.equal(await page.locator('#chart-first-column').isChecked(),true);await submit('삽입');await check(4);
 await page.locator('[data-edit-chart]').click();await page.locator('#chart-type').selectOption('doughnut');await submit('적용');await check(8);await page.locator('#undo').click();await check(4);await page.locator('#redo').click();await check(8);
 await page.locator('[data-edit-chart]').click();await page.locator('#chart-series-by').selectOption('columns');await submit('적용');await check(8);assert.equal((await get()).sheets[0].charts[0].seriesBy,'columns');await page.locator('#undo').click();assert.equal((await get()).sheets[0].charts[0].seriesBy,'rows');
 await page.locator('[data-edit-chart]').click();await page.locator('#chart-series-by').selectOption('columns');await page.getByRole('button',{name:'닫기',exact:true}).click();assert.equal((await get()).sheets[0].charts[0].seriesBy,'rows');await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('.chart-card').waitFor();await check(8);
 await roundtrip();await check(8);
 // The added direction control must not push Apply/Close outside a short window.
 await page.setViewportSize({width:1024,height:600});await load([[10,20,30,40]]);await open('A1:D1','pie');
 const dialogBounds=await page.getByRole('dialog').evaluate(el=>{const r=el.getBoundingClientRect(),body=el.querySelector('.modal-body'),footer=el.querySelector('.modal-actions').getBoundingClientRect();return {top:r.top,bottom:r.bottom,height:innerHeight,scrolls:body.scrollHeight>body.clientHeight,footerBottom:footer.bottom};});assert.ok(dialogBounds.top>=0&&dialogBounds.bottom<=dialogBounds.height,JSON.stringify(dialogBounds));assert.ok(dialogBounds.scrolls);assert.ok(dialogBounds.footerBottom<=dialogBounds.height);await page.screenshot({path:'test-results/chart-audit-compact-dialog.png'});await submit('삽입');await check(4);await page.locator('[data-edit-chart]').click();await page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();await page.locator('.backdrop').waitFor({state:'detached'});await page.setViewportSize({width:1280,height:900});
 // Four legend categories remain visible in a narrow chart instead of truncating to two.
 await load([[10,20,30,40]],[{id:'narrow',type:'doughnut',range:'A1:D1',title:'네 값 도넛형',x:60,y:40,width:280,height:300,legend:true}]);await check(4);assert.equal(await page.locator('.chart-legend-item').count(),4);assert.equal(await page.locator('.chart-legend-row').count(),2);
 const centered=await page.locator('.chart-card>svg').evaluate(svg=>[...svg.querySelectorAll('.chart-legend-row')].every(row=>{const b=row.getBBox();return Math.abs(b.x+b.width/2-svg.viewBox.baseVal.width/2)<.5;}));assert.equal(centered,true);await page.screenshot({path:'test-results/chart-audit-narrow.png'});
 assert.deepEqual(errors,[]);console.log('PASS: 10 types / 18 variants, four horizontal/vertical values, actual typing and drag selection, pie/doughnut slices, XY and bubble requirements, row/column switching, labels, legends, native XLSX round trips, undo/cancel/reload');
}catch(error){await page.screenshot({path:'test-results/chart-audit-failure.png'});throw error;}finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
