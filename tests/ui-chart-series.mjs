import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {desktopLaunch} from './electron-launch.mjs';
import {createWorkbook} from '../src/engine.mjs';
import {chartData} from '../src/chart-model.mjs';

const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','chart-series-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'})),page=await app.firstWindow(),errors=[];
page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(20000);await page.setViewportSize({width:1440,height:960});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const ready=()=>page.locator('#chart-preview-status').filter({hasText:/개 계열 사용/}).waitFor();
const getBook=()=>page.evaluate(()=>sheetStudio.getWorkbook());
const charts=async()=>(await getBook()).sheets[0].charts;
const row=i=>page.locator('[data-series-row="'+i+'"]');
const field=(i,key)=>row(i).locator('[data-series-field="'+key+'"]');
async function go(ref){await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();}
async function open(){await go('B1:B5');await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="chart-insert"]').click();await ready();}
async function apply(label){const dialog=page.getByRole('dialog');await dialog.getByRole('button',{name:label,exact:true}).click();await dialog.waitFor({state:'detached'});await settle();}
async function configure(i,{name,values,categories='A2:A5',color='#4472c4',size}){await field(i,'name').fill(name);await field(i,'values').fill(values);await field(i,'categories').fill(categories);await field(i,'color').fill(color);if(size)await field(i,'size').fill(size);}
async function editChart(index){await page.locator('[data-edit-chart="'+(await charts())[index].id+'"]').click();await ready();}
async function save(file){await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');}
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();const book=createWorkbook(),sheet=book.sheets[0];book.name='Explicit chart series';sheet.cells={A1:{raw:'항목'},B1:{raw:'실적'},D1:{raw:'목표'}};
 for(let i=0;i<4;i++){sheet.cells['A'+(i+2)]={raw:['가','나','다','라'][i]};sheet.cells['B'+(i+2)]={raw:String((i+1)*10)};sheet.cells['D'+(i+2)]={raw:String((i+1)*12)};sheet.cells['F'+(i+2)]={raw:String(i+1)};}
 await page.locator('#file-input').setInputFiles({name:'series.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Explicit chart series');await settle();await open();
 if(process.argv.includes('--baseline')){assert.equal(await page.locator('#chart-data-mode').count(),0);console.log('REPRODUCED: baseline chart settings have no explicit series editor.');}
 else{
  const before=JSON.stringify(await charts());await page.locator('#chart-data-mode').selectOption('series');await configure(0,{name:'=B1',values:'B2:B5'});await page.locator('[data-series-add]').click();await configure(1,{name:'목표',values:'D2:D5',color:'#e15759'});await page.locator('[data-series-up="1"]').click();await ready();assert.deepEqual(await page.locator('#chart-preview-data tbody tr td:last-child').allTextContents(),['12','24','36','48','10','20','30','40']);
  await field(0,'values').fill('D2:D4');await page.locator('#chart-preview-status').filter({hasText:/셀 수|같은 수/}).waitFor();await page.getByRole('dialog').getByRole('button',{name:'삽입',exact:true}).click();assert.match(await page.locator('#chart-error').innerText(),/셀 수|같은 수/);assert.equal(JSON.stringify(await charts()),before);await field(0,'values').fill('D2:D5');await ready();
  await page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();await settle();assert.equal(JSON.stringify(await charts()),before);
  await open();await page.locator('#chart-data-mode').selectOption('series');await configure(0,{name:'=B1',values:'B2:B5'});await page.locator('[data-series-add]').click();await configure(1,{name:'목표',values:'D2:D5',color:'#e15759'});await page.locator('[data-series-up="1"]').click();await ready();await apply('삽입');
  let data=await getBook(),chart=data.sheets[0].charts[0];assert.deepEqual(chart.series.map(s=>s.valRef.range),['D2:D5','B2:B5']);assert.deepEqual(chartData(chart,data.sheets[0],data).map(s=>s.name),['목표','실적']);assert.deepEqual(chart.series.map(s=>s.color),['#e15759','#4472c4']);
  await editChart(0);assert.equal(await page.locator('#chart-data-mode').inputValue(),'series');await page.locator('[data-series-down="0"]').click();await ready();await apply('적용');assert.equal((await charts())[0].series[0].valRef.range,'B2:B5');await page.locator('#undo').click();await settle();assert.equal((await charts())[0].series[0].valRef.range,'D2:D5');await page.locator('#redo').click();await settle();assert.equal((await charts())[0].series[0].valRef.range,'B2:B5');
  await editChart(0);await page.locator('#chart-type').selectOption('scatter');await field(0,'categories').fill('F2:F5');await field(1,'categories').fill('F2:F5');await field(0,'lines').check();await field(1,'lines').uncheck();await ready();assert.equal(await page.locator('#chart-preview-image .chart-point').count(),8);await apply('적용');assert.deepEqual((await charts())[0].series.map(s=>s.lines),[true,false]);
  await editChart(0);await page.locator('#chart-type').selectOption('bubble');await field(0,'size').fill('D2:D5');await field(1,'size').fill('B2:B5');await ready();assert.equal(await page.locator('#chart-preview-image .chart-point').count(),8);await apply('적용');data=await getBook();assert.deepEqual(chartData(data.sheets[0].charts[0],data.sheets[0],data)[0].sizes,[12,24,36,48]);
  await editChart(0);await page.locator('#chart-type').selectOption('combo');await field(0,'categories').fill('A2:A5');await field(1,'categories').fill('A2:A5');await field(0,'type').selectOption('bar');await field(1,'type').selectOption('line');await field(1,'secondary').check();await ready();await page.setViewportSize({width:1000,height:680});await field(1,'secondary').scrollIntoViewIfNeeded();await page.screenshot({path:path.join(root,'test-results/chart-series-narrow.png')});assert.equal(await page.locator('.modal').evaluate(n=>n.scrollWidth<=n.clientWidth+1),true);await apply('적용');assert.deepEqual((await charts())[0].series.map(s=>[s.type,s.secondary]),[['bar',false],['line',true]]);
  const file=path.join(profile,'series.xlsx');await save(file);await page.locator('#file-input').setInputFiles(file);await settle();await editChart(0);assert.ok((await charts())[0].source);assert.equal(await page.locator('#chart-data-mode-option').isVisible(),false);await page.locator('#chart-convert').check();await ready();assert.equal(await page.locator('#chart-data-mode').inputValue(),'series');assert.match(await field(0,'values').inputValue(),/\$B\$2:\$B\$5/);assert.match(await field(1,'values').inputValue(),/\$D\$2:\$D\$5/);assert.equal(await field(1,'secondary').isChecked(),true);await field(1,'name').fill('계획');await ready();await apply('적용');assert.equal((await charts())[0].source,undefined);await save(file);await page.locator('#file-input').setInputFiles(file);await settle();data=await getBook();assert.deepEqual(chartData(data.sheets[0].charts[0],data.sheets[0],data).map(s=>s.name),['실적','계획']);assert.deepEqual(chartData(data.sheets[0].charts[0],data.sheets[0],data).map(s=>s.values),[[10,20,30,40],[12,24,36,48]]);assert.deepEqual(errors,[]);
  console.log('PASS: non-adjacent series/name references/color/order, read-only invalid/cancel preview, undo/redo, scatter X/Y/connections, bubble sizes, combo axes, 1000x680 dialog and XLSX reopen/explicit conversion.');
 }
}catch(error){await page.screenshot({path:path.join(root,'test-results/chart-series-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
