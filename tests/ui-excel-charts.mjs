import {_electron} from 'playwright';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import platform from '../scripts/platform.cjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {chartFixture,addChart} from './fixtures/charts.mjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const env={...process.env,SHEET_STUDIO_DATA:path.join(root,'.package','excel-charts-'+Date.now())};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));
const go=async range=>{await page.locator('#name-box').fill(range);await page.locator('#name-box').press('Enter');};
const get=()=>page.evaluate(()=>sheetStudio.getWorkbook());
async function insert(type,grouping='clustered'){
  await go('A1:C5');await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="chart-insert"]').click();
  await page.locator('#chart-type').selectOption(type);await page.locator('#chart-grouping').selectOption(grouping);
  await page.locator('#chart-title').fill('매출 분석 '+type);await page.locator('#chart-labels').check();
  await page.getByRole('dialog').getByRole('button',{name:'삽입',exact:true}).click();await page.locator('.chart-card').waitFor();
}
try{
 await page.locator('#cells .cell').first().waitFor();
 await page.locator('#file-input').setInputFiles({name:'매출.csv',mimeType:'text/csv',buffer:Buffer.from('상품,매출,비용\n커피,40,22\n차,25,18\n주스,-10,5\n물,35,10')});
 await page.waitForFunction(()=>sheetStudio.getWorkbook().sheets[0].cells.A1?.raw==='상품');
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});
 for(const type of ['bar','bar-horizontal','line','area','pie','doughnut','scatter','bubble','radar','combo']){
   await insert(type,type==='bar'?'stacked':'clustered');const book=await get();assert.equal(book.sheets[0].charts[0].type,type);
   const svg=await page.locator('.chart-card>svg').innerHTML();assert.ok(!/NaN|Infinity/.test(svg));assert.ok(svg.length>300);
   await page.screenshot({path:path.join(root,'test-results/excel-chart-'+type+'.png')});
   await page.locator('[data-remove-chart]').click();
 }
 await insert('combo');
 const destination=path.join(root,'.package/ui-chart-save.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},destination);
 await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');
 const saved=await readSpreadsheet(await fs.readFile(destination),'saved.xlsx');assert.equal(saved.sheets[0].charts[0].type,'combo');assert.equal(saved.sheets[0].charts[0].series.length,2);
 await page.locator('#file-input').setInputFiles(destination);await page.waitForFunction(()=>!!sheetStudio.getWorkbook().sheets[0].charts[0]?.source);
 await page.locator('[data-edit-chart]').click();assert.ok(await page.locator('#chart-secondary').isDisabled());await page.locator('#chart-title').fill('보존된 혼합 차트');await page.locator('#chart-legend').uncheck();await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();
 assert.equal((await get()).sheets[0].charts[0].title,'보존된 혼합 차트');assert.ok((await get()).sheets[0].charts[0].source);
 await go('B2');await page.keyboard.press('F2');await page.locator('#editor').fill('99');await page.locator('#editor').press('Enter');await page.waitForFunction(()=>document.querySelector('.chart-card>svg')?.innerHTML.includes('99'));assert.ok((await page.locator('.chart-card>svg').innerHTML()).includes('99'));
 await page.locator('[data-edit-chart]').click();await page.locator('#chart-type').selectOption('area');await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();assert.ok((await page.locator('#chart-error').innerText()).includes('변환'));
 await page.locator('#chart-convert').check();await page.locator('#chart-range').fill('A1:C5');await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();assert.equal((await get()).sheets[0].charts[0].source,undefined);assert.equal((await get()).sheets[0].charts[0].type,'area');
 await page.locator('#grid').focus();await page.keyboard.press((process.platform==='darwin'?'Meta':'Control')+'+z');assert.ok((await get()).sheets[0].charts[0].source);
 await page.reload();await page.locator('.chart-card').waitFor();assert.equal((await get()).sheets[0].charts[0].title,'보존된 혼합 차트');
 const unknown=chartFixture();addChart(unknown);const zip=await JSZip.loadAsync((await writeSpreadsheet(unknown)).data),part='xl/charts/sheetStudioChart1.xml';zip.file(part,(await zip.file(part).async('string')).replaceAll('barChart','stockChart'));
 await page.locator('#file-input').setInputFiles({name:'unsupported.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:await zip.generateAsync({type:'nodebuffer'})});await page.waitForFunction(()=>sheetStudio.getWorkbook().sheets[0].charts[0]?.type==='unsupported');
 await page.locator('[data-edit-chart]').click();assert.ok(await page.locator('#chart-title').isDisabled());await page.locator('#chart-convert').check();await page.locator('#chart-type').selectOption('bar');await page.locator('#chart-title').fill('변환한 차트');await page.locator('#chart-range').fill('A1:C5');await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();assert.equal((await get()).sheets[0].charts[0].title,'변환한 차트');assert.equal((await get()).sheets[0].charts[0].source,undefined);
 assert.deepEqual(errors,[]);console.log('Excel chart UI passed: 10 types, multi-series, settings, native XLSX save/reopen, live updates, explicit conversion, undo and reload.');
}finally{await app.close();}
