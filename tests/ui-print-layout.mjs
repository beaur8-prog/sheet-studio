import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createWorkbook,address} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';

const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','print-layout-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:profile};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch(desktopLaunch(root,env)),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(15000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
async function print({selection=false,ignore=false}={}){
  await page.locator('[data-tab="페이지 레이아웃"]').click();await page.locator('[data-action="print"]').click();const dialog=page.getByRole('dialog',{name:'인쇄',exact:true});await dialog.waitFor();
  if(selection)await page.locator('#print-scope').selectOption('selection');if(ignore)await page.locator('#print-use-area').uncheck();
  await page.locator('#print-charts').uncheck();await page.locator('#print-images').uncheck();await page.locator('#print-grid-lines').check();const n=await page.evaluate(()=>testPrintDocuments.length);
  await dialog.getByRole('button',{name:'인쇄',exact:true}).click();await page.waitForFunction(n=>testPrintDocuments.length>n,n);return page.evaluate(()=>testPrintDocuments.at(-1));
}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await fs.mkdir(profile,{recursive:true});await page.locator('.cell').first().waitFor();
  await page.evaluate(()=>{globalThis.testPrintDocuments=[];new MutationObserver(records=>{for(const record of records)for(const node of record.addedNodes)if(node instanceof HTMLIFrameElement&&node.classList.contains('sheet-print-frame'))node.contentWindow.print=()=>testPrintDocuments.push(node.contentDocument.documentElement.outerHTML);}).observe(document.body,{childList:true});});
  let book=createWorkbook(),s=book.sheets[0];book.name='Print-layout';s.name='Monthly report';s.defaultColumnWidth=130;
  for(let r=0;r<81;r++)for(let c=0;c<8;c++)s.cells[address(r,c)]={raw:r===0?['Row','Value','Total','Rate','Dept','Month','Target','Actual'][c]:c===0?'ROW'+String(r).padStart(4,'0'):String(r*(c+1)),style:r===0?{background:'#24476b',color:'#ffffff',bold:true}:{}};
  s.cells.C2.raw='=B2*2';s.cells.A1000000={raw:'EXCLUDED_PRIVATE_CONTENT'};s.xlsxFeatures={pageSetup:'<pageSetup paperSize="9" scale="100" orientation="portrait"/>',pageMargins:'<pageMargins top="0.4724409448818898" bottom="0.4724409448818898" left="0.4724409448818898" right="0.4724409448818898"/>',printOptions:'<printOptions gridLines="1"/>'};book.chartNames=[{name:'_xlnm.Print_Area',sheetId:s.id,formula:"'Monthly report'!$A$1:$H$81"},{name:'_xlnm.Print_Titles',sheetId:s.id,formula:"'Monthly report'!$1:$1,'Monthly report'!$A:$A"}];
  book=await readSpreadsheet((await writeSpreadsheet(book)).data,'Print-layout.xlsx');
  await page.locator('#file-input').setInputFiles({name:'Print-layout.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from((await writeSpreadsheet(book)).data)});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Print-layout');await settle();
  // The worker uses one shared calculation session; printing must not mutate
  // print names or any source cell, including the distant excluded cell.
  const before=await page.evaluate(()=>JSON.stringify(sheetStudio.getWorkbook())),html=await print();assert.equal(await page.evaluate(()=>JSON.stringify(sheetStudio.getWorkbook())),before);assert.equal((html.match(/class="print-page"/g)||[]).length,4);assert.doesNotMatch(html,/EXCLUDED_PRIVATE_CONTENT/);
  const file=path.join(profile,'layout.html');await fs.writeFile(file,html);
  const rendered=await app.evaluate(async({BrowserWindow},file)=>{
    const win=new BrowserWindow({show:false,width:794,height:1200,webPreferences:{sandbox:true,contextIsolation:true,nodeIntegration:false}});
    try{await win.loadFile(file);await win.webContents.executeJavaScript('document.fonts.ready');const metrics=await win.webContents.executeJavaScript(`Array.from(document.querySelectorAll('.print-page')).map(p=>{const r=p.getBoundingClientRect(),t=p.querySelector('table').getBoundingClientRect();return {pageHeight:r.height,tableHeight:t.height,pageWidth:r.width,tableWidth:t.width,first:p.querySelector('[data-ref]')?.textContent,last:p.querySelector('tr:last-child')?.textContent};})`);const pdf=await win.webContents.printToPDF({printBackground:true,preferCSSPageSize:true,displayHeaderFooter:false});return {metrics,pdf:pdf.toString('base64')};}finally{win.destroy();}
  },file);
  for(const m of rendered.metrics){assert.ok(m.tableHeight<=m.pageHeight+.1,JSON.stringify(m));assert.ok(m.tableWidth<=703,JSON.stringify(m));assert.equal(m.first,'Row');assert.match(m.last,/ROW0040|ROW0080/);}
  await fs.writeFile(path.join(root,'test-results/print-layout.pdf'),Buffer.from(rendered.pdf,'base64'));await fs.writeFile(path.join(root,'test-results/print-layout-metrics.json'),JSON.stringify(rendered.metrics,null,2));
  await page.locator('#name-box').fill('B2:C3');await page.locator('#name-box').press('Enter');await settle();const selected=await print({selection:true});assert.equal((selected.match(/class="print-page"/g)||[]).length,1);assert.doesNotMatch(selected,/data-ref="A1"|data-ref="A2"|data-ref="A1000000"/);assert.match(selected,/data-ref="C2"[^>]*><div[^>]*>4</);
  // Ignoring a saved small area makes the original large-range guard visible,
  // rather than silently falling back to that large range on parse failures.
  await page.locator('[data-action="print"]').click();await page.locator('#print-use-area').uncheck();await page.getByRole('dialog',{name:'인쇄',exact:true}).getByRole('button',{name:'인쇄',exact:true}).click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent.includes('20,000'));await page.getByRole('button',{name:'취소',exact:true}).click();
  assert.deepEqual(errors,[]);console.log('PASS: saved print area excludes distant content, four physical pages repeat rows/columns, all table geometry fits, formula snapshot is immutable, selection bypasses saved titles/area, ignore-area guard, PDF generated');
}catch(error){await page.screenshot({path:path.join(root,'test-results/print-layout-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
