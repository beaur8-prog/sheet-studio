import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import platform from '../scripts/platform.cjs';
import {chartsheetBook} from './fixtures/sheet-kinds.mjs';
import {drawingFixture} from './fixtures/drawings.mjs';
import {createWorkbook} from '../src/engine.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';
const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','print-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch(desktopLaunch(root,env)),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(15000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));};
const load=async(name,bytes)=>{await page.locator('#file-input').setInputFiles({name:name+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(bytes)});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,name);await settle();};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const openPrint=async()=>{await page.locator('[data-tab="페이지 레이아웃"]').click();await page.locator('[data-action="print"]').click();await page.getByRole('dialog',{name:'인쇄',exact:true}).waitFor();};
async function print(){const before=await page.evaluate(()=>testPrintDocuments.length);await page.getByRole('dialog',{name:'인쇄',exact:true}).getByRole('button',{name:'인쇄',exact:true}).click();await page.waitForFunction(n=>testPrintDocuments.length>n,before);return page.evaluate(()=>testPrintDocuments.at(-1));}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  // Intercept only this isolated test page's iframe print calls, so tests never
  // open a native printer dialog or touch a user's printer configuration.
  await page.evaluate(()=>{globalThis.testPrintDocuments=[];new MutationObserver(records=>{for(const record of records)for(const node of record.addedNodes)if(node instanceof HTMLIFrameElement&&node.classList.contains('sheet-print-frame'))node.contentWindow.print=()=>testPrintDocuments.push(node.contentDocument.documentElement.outerHTML);}).observe(document.body,{childList:true});});
  const b=await readSpreadsheet(await drawingFixture(),'print.xlsx'),s=b.sheets[0];s.cells.B1={raw:'=RAND()'};s.cells.A3={raw:'-12',style:{numberFormat:'0;[Red](0)',format:'number'}};s.merges=['A4:B4'];s.cells.A4={raw:'Merged print title',style:{background:'#e2efda',bold:true}};s.charts=[{id:'c1',type:'line',title:'Current values',range:'A2:B3',series:[{name:'Random',valRef:{sheetId:s.id,range:'B1:B1'}}],width:500,height:300}];
  await load('print-flow',(await writeSpreadsheet(b)).data);await settle();const random=await page.evaluate(()=>sheetStudio.evaluateAsync('B1'));await openPrint();await page.locator('#print-orientation').selectOption('landscape');await page.locator('#print-grid-lines').check();const html=await print();assert.match(html,/Merged print title/);assert.match(html,/rowspan="1" colspan="2"/);assert.match(html,/class="print-chart"/);assert.match(html,/class="print-image"/);assert.match(html,/미리 보기를 지원하지 않는 이미지·도형 1개/);assert.match(html,/color:#ff0000/);assert.match(html,/size:A4 landscape/);assert.ok(html.includes(String(random)));assert.doesNotMatch(html,/chart-settings|chart-resize/);
  const printedHTML=path.join(dataDir,'printed.html');await fs.writeFile(printedHTML,html);
  const rendered=await app.evaluate(async({BrowserWindow},file)=>{const win=new BrowserWindow({show:false,width:1123,height:900,webPreferences:{sandbox:true,contextIsolation:true,nodeIntegration:false}});try{await win.loadFile(file);await win.webContents.executeJavaScript('document.fonts.ready.then(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))))');const pdf=await win.webContents.printToPDF({printBackground:true,preferCSSPageSize:true}),png=(await win.webContents.capturePage()).toPNG();return {pdf:pdf.toString('base64'),png:png.toString('base64')};}finally{win.destroy();}},printedHTML);
  await fs.writeFile(path.join(root,'test-results/sheet-print.pdf'),Buffer.from(rendered.pdf,'base64'));await fs.writeFile(path.join(root,'test-results/sheet-print.png'),Buffer.from(rendered.png,'base64'));assert.ok(Buffer.from(rendered.pdf,'base64').length>3000);
  await go('A3:B4');await openPrint();await page.locator('#print-scope').selectOption('selection');await page.locator('#print-charts').uncheck();await page.locator('#print-images').uncheck();const selected=await print();assert.match(selected,/Merged print title/);assert.doesNotMatch(selected,/class="print-chart"|class="print-image"|data-ref="B1"/);
  await load('print-dashboard',(await writeSpreadsheet(chartsheetBook())).data);await openPrint();assert.equal(await page.locator('#print-scope').inputValue(),'charts');const chart=await print();assert.match(chart,/class="chart-point"/);assert.doesNotMatch(chart,/<table/);
  const large=createWorkbook();large.sheets[0].cells={A1:{raw:'first'},A1000000:{raw:'last'}};await load('large-print',(await writeSpreadsheet(large)).data);await openPrint();await page.getByRole('dialog',{name:'인쇄',exact:true}).getByRole('button',{name:'인쇄',exact:true}).click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent.includes('20,000'));assert.equal(await page.getByRole('dialog',{name:'인쇄',exact:true}).count(),1);await page.getByRole('button',{name:'취소',exact:true}).click();assert.deepEqual(errors,[]);
  console.log('PASS: actual print scope/options, current RAND session, merges/format colors, chart/image flow, unsupported object note, PDF generation, selection-only printing, chart sheet SVG and large-range guard');
}catch(error){await page.screenshot({path:path.join(root,'test-results/sheet-print-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
