import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook,address} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const dataDir=path.join(root,'.package','bar-style-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(8000);await page.setViewportSize({width:1280,height:900});
const settle=()=>page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const get=()=>page.evaluate(()=>sheetStudio.getWorkbook());
let fixture=0;
const load=async rows=>{const b=createWorkbook();b.name='막대형 미리보기 '+ ++fixture;rows.forEach((row,r)=>row.forEach((v,c)=>b.sheets[0].cells[address(r,c)]={raw:String(v)}));await page.locator('#file-input').setInputFiles({name:'bar-style.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(n=>sheetStudio.getWorkbook().name===n,b.name);await settle();};
const insert=async(range,dialog=false)=>{await page.locator('#name-box').fill(range);await page.locator('#name-box').press('Enter');await page.locator('[data-tab="삽입"]').click();await page.locator(`[data-action="${dialog?'chart-insert':'chart-bar'}"]`).click();if(dialog){assert.equal(await page.locator('#chart-legend').isChecked(),false);assert.equal(await page.locator('#chart-legend').isDisabled(),true);assert.match(await page.locator('#chart-data-help').innerText(),/가로축/);await page.getByRole('dialog').getByRole('button',{name:'삽입',exact:true}).click();}await page.waitForFunction(()=>document.querySelectorAll('.chart-point').length===4);await settle();};
const check=async(names,{wrapped=false}={})=>{
 assert.equal(await page.locator('.chart-legend-item').count(),0);assert.equal(await page.locator('.chart-category').count(),4);assert.deepEqual(await page.locator('.chart-category').evaluateAll(items=>items.map(el=>el.getAttribute('aria-label'))),names);
 const geometry=await page.locator('.chart-card>svg').evaluate(svg=>{const bars=[...svg.querySelectorAll('rect.chart-point')],labels=[...svg.querySelectorAll('.chart-category')];return {colors:[...new Set(bars.map(b=>b.getAttribute('fill')))],rows:labels.map((el,i)=>{const b=bars[i].getBBox(),l=el.getBBox(),t=el.querySelector('text');return {barCenter:b.x+b.width/2,labelCenter:l.x+l.width/2,left:l.x,right:l.x+l.width,labelTop:l.y,bottom:l.y+l.height,width:b.width,lines:t.querySelectorAll('tspan').length};}),height:svg.viewBox.baseVal.height};});
 assert.equal(geometry.colors.length,1);for(let i=0;i<geometry.rows.length;i++){const row=geometry.rows[i];assert.ok(Math.abs(row.barCenter-row.labelCenter)<.6,JSON.stringify(row));assert.ok(row.width<=56.1);assert.ok(row.bottom<=geometry.height);if(i)assert.ok(row.left>=geometry.rows[i-1].right+8);if(wrapped)assert.equal(row.lines,2);}
};
try{
 await page.locator('.cell').first().waitFor();await fs.mkdir('test-results',{recursive:true});
 const names=['커피','차','주스','물'];await load([names,[120,90,150,80]]);await insert('A1:D2');await check(names);assert.equal((await get()).sheets[0].charts[0].seriesBy,'rows');
 await page.locator('[data-edit-chart]').click();await page.locator('#chart-title').fill('음료별 판매량');await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await settle();await page.screenshot({path:'test-results/bar-style-clean.png'});
 const output=path.join(dataDir,'clean-bar.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');await page.locator('#file-input').setInputFiles(output);await page.waitForFunction(()=>!!sheetStudio.getWorkbook().sheets[0].charts[0]?.source);await settle();await check(names);assert.equal((await get()).sheets[0].charts[0].legend,false);
 await page.reload();await page.locator('.chart-card').waitFor();await settle();await check(names);
 const long=['아주 긴 항목 이름이 있는 첫 번째 비교 대상','아주 긴 항목 이름이 있는 두 번째 비교 대상','아주 긴 항목 이름이 있는 세 번째 비교 대상','아주 긴 항목 이름이 있는 네 번째 비교 대상'];await load([long,[10,20,30,40]]);await insert('A1:D2',true);await check(long,{wrapped:true});await page.screenshot({path:'test-results/bar-style-wrapped.png'});
 await page.locator('#zoom').evaluate(el=>{el.value=150;el.dispatchEvent(new Event('input',{bubbles:true}));});await settle();await check(long,{wrapped:true});await page.locator('#zoom').evaluate(el=>{el.value=100;el.dispatchEvent(new Event('input',{bubbles:true}));});
 await load([['항목','매출','비용'],['커피',120,70],['차',90,40],['주스',150,80]]);await page.locator('#name-box').fill('A1:C4');await page.locator('#name-box').press('Enter');await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="chart-bar"]').click();await page.waitForFunction(()=>document.querySelectorAll('rect.chart-point').length===6);await settle();assert.equal(await page.locator('.chart-legend-item').count(),2);assert.equal(await page.locator('rect.chart-point').evaluateAll(bars=>new Set(bars.map(b=>b.getAttribute('fill'))).size),2);
 assert.deepEqual(errors,[]);console.log('PASS: category names directly under centered uniform bars, no redundant legend, readable spacing, wrapped/zoomed labels, dialog defaults, multi-series distinction and native XLSX/reload');
}catch(error){await page.screenshot({path:'test-results/bar-style-failure.png'});throw error;}finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
