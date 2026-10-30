import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createWorkbook} from '../src/engine.mjs';

// A realistic composite journey through the visible app. A private profile and
// file-dialog responses keep the user's documents, clipboard and printer intact.
const root=path.resolve(import.meta.dirname,'..');
const profile=path.join(root,'.package','dashboard-journey-'+Date.now());
const env={...process.env,SHEET_STUDIO_DATA:profile};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch(desktopLaunch(root,env));
const page=await app.firstWindow(),errors=[],checks=[];
page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(20000);
await page.setViewportSize({width:1440,height:1000});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
const book=()=>page.evaluate(()=>sheetStudio.getWorkbook());
const tab=async name=>{await page.getByRole('button',{name,exact:true}).click();await settle();};
const action=async name=>{await page.locator(`[data-action="${name}"]`).first().click();await settle();};
const edit=async(ref,raw)=>{await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();};
const apply=async(name='적용')=>{await page.getByRole('dialog').getByRole('button',{name,exact:true}).click();await page.getByRole('dialog').waitFor({state:'detached'});await settle();};
const record=label=>{checks.push(label);console.log('PASS:',label);};
async function save(file){
 await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);
 await page.evaluate(()=>document.querySelector('.toast')?.remove());await page.locator('#save').click();
 await page.locator('.toast').filter({hasText:'파일을 저장했습니다.'}).waitFor();await settle();
 assert.ok((await fs.stat(file)).size>0);
}
async function open(file){
 await app.evaluate(({dialog},file)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[file]});},file);
 await page.evaluate(()=>document.querySelector('.toast')?.remove());
 await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.send('sheet:menu-action','open'));
 await page.locator('.toast').filter({hasText:'파일을 열었습니다.'}).waitFor();await settle();
}
async function verifyDashboard(top=200){
 await tab('Sheet2');assert.equal(await value('A1'),'Region');assert.equal(await value('D1'),'Amount');
 assert.equal(await value('D2'),top);assert.equal(await value('F2'),top+430);
 assert.equal(await page.locator('rect.chart-point').count(),12);
 assert.doesNotMatch(await page.locator('.chart-card>svg').innerHTML(),/NaN|Infinity/);
 const dashboard=(await book()).sheets[1];assert.equal(Object.keys(dashboard.cells).length,4);
 const chart=dashboard.charts[0];
 if(chart.source)assert.deepEqual(chart.series.map(series=>series.valRef.range),['B2:B5','C2:C5','D2:D5']);
 else assert.equal(chart.range,'A1:D5');
}
async function verifyPivot(total){
 await tab('피벗 1');assert.equal(await value('B5'),total);
 assert.equal(await value('A2'),'Busan');assert.equal(await value('A3'),'Jeju');assert.equal(await value('A4'),'Seoul');
}
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
 const b=createWorkbook();b.name='Sales dashboard';const source=b.sheets[0];source.name='Source';
 const rows=[['Region','Qty','Price','Amount'],['Seoul',2,100,''],['Busan',1,80,''],['Seoul',3,50,''],['Jeju',4,40,''],['Busan',2,60,'']];
 for(let r=0;r<rows.length;r++)for(let c=0;c<4;c++)if(rows[r][c]!=='')source.cells[String.fromCharCode(65+c)+(r+1)]={raw:String(rows[r][c])};
 await page.locator('#file-input').setInputFiles({name:'sales-dashboard.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await settle();
 await go('A1:D6');await tab('삽입');await action('table-style');await page.locator('#table-name').fill('Sales');await apply('표 만들기');
 await go('D2');await action('table-calc');await page.locator('#table-calc-formula').fill('=[@Qty]*[@Price]');await apply('전체 행에 적용');
 assert.equal(await value('D2'),200);assert.equal(await value('D6'),120);
 await action('add-sheet');
 await edit('A1','=Sales[#Headers]');await edit('A2','=SORT(FILTER(Sales,Sales[Amount]>=100),4,-1)');await edit('F1','Total');await edit('F2','=SUM(INDEX(A2#,,4))');
 assert.equal(await value('D5'),120);assert.equal(await value('F2'),630);
 await go('A1:D5');await tab('삽입');await action('chart-insert');await page.locator('#chart-type').selectOption('bar');await page.locator('#chart-title').fill('Sales dashboard');await apply('삽입');
 const originalChart=(await book()).sheets[1].charts[0],title=await page.locator('[data-chart-title]').boundingBox();assert.ok(title);
 await page.mouse.move(title.x+title.width/2,title.y+title.height/2);await page.mouse.down();await page.mouse.move(title.x+title.width/2-originalChart.x,title.y+title.height/2+160-originalChart.y,{steps:8});await page.mouse.up();await settle();
 assert.ok(Math.abs((await book()).sheets[1].charts[0].y-160)<1);
 await verifyDashboard();record('visible table creation and calculated column feed a sorted/filtered cross-sheet dashboard and chart');
 await go('A1:D5');await action('pivot');await page.getByRole('dialog').getByLabel('값 필드 1',{exact:true}).selectOption('Amount');await apply('만들기');
 await verifyPivot(630);record('pivot groups calculated array results without editable follower records');

 await tab('Source');await edit('B2','4');assert.equal(await value('D2'),400);await verifyDashboard(400);
 await tab('Source');await page.locator('#undo').click();await settle();await verifyDashboard();
 await tab('Source');await page.locator('#redo').click();await settle();await verifyDashboard(400);
 await tab('피벗 1');await page.locator('#pivot-bar').getByRole('button',{name:'새로 고침',exact:true}).click();await settle();await verifyPivot(830);
 await page.locator('#undo').click();await settle();await verifyPivot(630);await page.locator('#redo').click();await settle();await verifyPivot(830);
 record('source edits, undo/redo and explicit pivot refresh keep independent chart and report results consistent');

 // Automatic table expansion should propagate its formula; the low-value row
 // stays outside this dashboard's filter rather than changing the chart range.
 await tab('Source');await edit('A7','Jeju');await edit('B7','1');await edit('C7','30');
 assert.equal(await value('D7'),30);assert.equal((await book()).sheets[0].tables[0].range,'A1:D7');
 await go('B2');await tab('삽입');await action('table-style');await page.locator('#table-name').fill('Revenue');await apply();
 await verifyDashboard(400);assert.match((await book()).sheets[1].cells.A2.raw,/FILTER\(Revenue,Revenue\[Amount\]/);
 record('table expansion fills the calculated column and table rename rewrites the dashboard formulas');

 const file=path.join(profile,'dashboard.xlsx');await save(file);await open(file);await verifyDashboard(400);await verifyPivot(830);
 assert.equal((await book()).sheets[0].tables[0].name,'Revenue');assert.equal((await book()).sheets[0].tables[0].range,'A1:D7');
 await verifyDashboard(400);const originalSeries=(await book()).sheets[1].charts[0].series;
 await page.locator('[data-edit-chart]').click();assert.equal(await page.locator('#chart-range').isVisible(),false);
 assert.match(await page.locator('#chart-source-references').textContent(),/'Sheet2'!\$D\$2:\$D\$5/);
 await page.locator('#chart-convert').check();assert.equal(await page.locator('#chart-range').isVisible(),true);assert.equal(await page.locator('#chart-source-references').isVisible(),false);
 await page.locator('#chart-convert').uncheck();assert.equal(await page.locator('#chart-range').isVisible(),false);await apply();
 assert.deepEqual((await book()).sheets[1].charts[0].series,originalSeries);
 record('reopened chart settings show original per-series references and retain them after conversion is cancelled');
 await tab('Source');await edit('C2','110');await verifyDashboard(440);await tab('피벗 1');await page.locator('#pivot-bar').getByRole('button',{name:'새로 고침',exact:true}).click();await settle();await verifyPivot(870);
 await save(file);await open(file);await verifyDashboard(440);await verifyPivot(870);
 record('two real XLSX save/reopen cycles preserve table definitions, dynamic source sparsity, chart and editable pivot source');

 await verifyDashboard(440);await page.evaluate(()=>{globalThis.dashboardPrints=[];new MutationObserver(records=>{for(const record of records)for(const node of record.addedNodes)if(node instanceof HTMLIFrameElement&&node.classList.contains('sheet-print-frame'))node.contentWindow.print=()=>dashboardPrints.push(node.contentDocument.documentElement.outerHTML);}).observe(document.body,{childList:true});});
 await tab('페이지 레이아웃');await action('page-setup');await page.locator('#ps-mode').selectOption('fit');await page.locator('#ps-fit-width').fill('1');await page.locator('#ps-fit-height').fill('1');await page.locator('#ps-orientation').selectOption('landscape');await apply();
 const before=await book();await action('print');await apply('인쇄');await page.waitForFunction(()=>dashboardPrints.length===1);
 const html=await page.evaluate(()=>dashboardPrints[0]);assert.match(html,/Sales dashboard/);assert.match(html,/>440</);assert.match(html,/>870</);assert.doesNotMatch(html,/>계산 중…</);assert.equal((html.match(/class="print-page(?:\s|")/g)||[]).length,1);assert.deepEqual(await book(),before);
 const htmlFile=path.join(profile,'dashboard.html');await fs.writeFile(htmlFile,html);
 const pdf=await app.evaluate(async({BrowserWindow},file)=>{const win=new BrowserWindow({show:false,webPreferences:{sandbox:true,nodeIntegration:false,contextIsolation:true}});try{await win.loadFile(file);await win.webContents.executeJavaScript('document.fonts.ready');return (await win.webContents.printToPDF({printBackground:true,preferCSSPageSize:true})).toString('base64');}finally{win.destroy();}},htmlFile);
 const bytes=Buffer.from(pdf,'base64');assert.equal((bytes.toString('latin1').match(/\/Type \/Page\b/g)||[]).length,1);await fs.writeFile(path.join(root,'test-results/dashboard-journey.pdf'),bytes);
 record('final dashboard table, total and chart print together on one physical PDF page without changing the document');
 assert.deepEqual(errors,[]);await page.screenshot({path:path.join(root,'test-results/dashboard-journey.png')});await fs.writeFile(path.join(root,'test-results/dashboard-journey.json'),JSON.stringify({at:new Date().toISOString(),platform:process.platform,checks},null,2));
}catch(error){await page.screenshot({path:path.join(root,'test-results/dashboard-journey-failure.png')});console.error(JSON.stringify({dialogs:await page.getByRole('dialog').allTextContents(),errors}));throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
