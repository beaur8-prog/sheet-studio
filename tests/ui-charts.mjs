import {chromium,_electron} from 'playwright';
import path from 'node:path';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';

const root=path.resolve(import.meta.dirname,'..'),version=createRequire(import.meta.url)('../package.json').version;
const packaged=process.argv.includes('--packaged'),desktop=packaged||process.argv.includes('--desktop');
let application,browser,page;
if(desktop){
  const env={...process.env,SHEET_STUDIO_DATA:path.join(root,'.package','chart-smoke-'+Date.now())};delete env.ELECTRON_RUN_AS_NODE;
  application=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
  page=await application.firstWindow();await application.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].setSize(1024,768));
}else{browser=await chromium.launch({headless:true,channel:'msedge'});page=await browser.newPage({viewport:{width:1540,height:970}});}
const errors=[];page.on('pageerror',e=>errors.push(e.message));
const settle=()=>page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
const go=async address=>{await page.locator('#name-box').fill(address);await page.locator('#name-box').press('Enter');await settle();};
const chart=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].charts[0]);
const key=process.platform==='darwin'?'Meta':'Control';
const undo=async()=>{await page.locator('#grid').focus();await page.keyboard.press(key+'+z');await settle();};
const redo=async()=>{await page.locator('#grid').focus();await page.keyboard.press(key+'+Shift+z');await settle();};
const zoom=async value=>{await page.locator('#zoom').evaluate((el,value)=>{el.value=value;el.dispatchEvent(new Event('input',{bubbles:true}));},value);await settle();};
async function drag(edge,dx,dy,{cancel=false}={}){
  // Scrolling/zoom rebuilds chart DOM on animation frames. Measure after that render.
  await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
  const handle=page.locator('[data-chart-resize="'+edge+'"]').first();
  await handle.scrollIntoViewIfNeeded();
  await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
  const rect=await handle.evaluate(el=>{const {x,y,width,height}=el.getBoundingClientRect();return{x,y,width,height};});
  assert.ok(rect,'Chart resize handle is visible after rendering');
  await page.mouse.move(rect.x+rect.width/2,rect.y+rect.height/2);await page.mouse.down();
  await page.mouse.move(rect.x+rect.width/2+dx,rect.y+rect.height/2+dy,{steps:8});
  if(cancel)await page.keyboard.press('Escape');
  await page.mouse.up();await settle();
}
try{
  if(!desktop)await page.goto('http://127.0.0.1:5192');await page.locator('#cells .cell').first().waitFor();
  await page.locator('#file-input').setInputFiles({name:'차트.csv',mimeType:'text/csv',buffer:Buffer.from('상품,매출\n커피,40\n차,25\n주스,-10')});
  await page.waitForFunction(()=>sheetStudio.getWorkbook().sheets[0].cells.A1?.raw==='상품');
  await go('A1:B4');await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="chart-bar"]').click();
  await page.locator('[data-chart-resize]').first().waitFor();assert.equal(await page.locator('[data-chart-resize]').count(),8);
  const initial=await chart();assert.equal(initial.width,500);
  await drag('se',150,90);assert.equal((await chart()).width,650);assert.equal((await chart()).height,390);
  const svg=await page.locator('.chart-card svg').getAttribute('viewBox');assert.equal(svg,'0 0 614 320');
  await undo();assert.equal((await chart()).width,500);await redo();assert.equal((await chart()).width,650);
  await drag('nw',40,30);const moved=await chart();assert.equal(moved.x,initial.x+40);assert.equal(moved.y,initial.y+30);assert.equal(moved.width,610);assert.equal(moved.height,360);
  await undo();await undo();assert.equal((await chart()).width,500);
  await zoom(150);await drag('se',120,60);assert.equal((await chart()).width,580);assert.equal((await chart()).height,340);
  const box=await page.locator('.chart-card').evaluate(el=>({width:el.getBoundingClientRect().width}));assert.equal(Math.round(box.width),870);
  await drag('se',40,20,{cancel:true});assert.equal((await chart()).width,580);
  await undo();assert.equal((await chart()).width,500);await redo();assert.equal((await chart()).width,580);
  await zoom(100);
  await page.locator('#grid').evaluate(el=>{el.scrollLeft=100;el.scrollTop=20;});
  await drag('e',50,0);assert.equal((await chart()).width,630);
  await page.locator('[data-chart-resize="se"]').focus();const selection=await page.locator('#name-box').inputValue();
  await page.keyboard.press('ArrowRight');await page.waitForFunction(()=>document.activeElement?.dataset.chartResize==='se');
  await page.keyboard.press('Shift+ArrowDown');assert.equal((await chart()).width,640);assert.equal((await chart()).height,341);assert.equal(await page.locator('#name-box').inputValue(),selection);
  await page.locator('[data-chart-title]').dblclick();await page.getByRole('dialog').locator('input').fill('분기 매출');await page.getByRole('button',{name:'확인',exact:true}).click();
  assert.equal((await chart()).title,'분기 매출');
  await settle();const beforeMove=await chart(),title=await page.locator('[data-chart-title]').evaluate(el=>{const {x,y}=el.getBoundingClientRect();return{x,y};});
  await page.mouse.move(title.x+40,title.y+12);await page.mouse.down();await page.mouse.move(title.x+60,title.y+37,{steps:5});await page.mouse.up();
  assert.equal((await chart()).x,beforeMove.x+20);assert.equal((await chart()).y,beforeMove.y+25);
  await drag('se',-900,-600);assert.equal((await chart()).width,280);assert.equal((await chart()).height,180);await undo();
  const saved=await chart();await page.reload();await page.locator('.chart-card').waitFor();assert.deepEqual(await chart(),saved);
  // Save through the UI and reopen the JSON document, including legacy charts without sizes.
  let json;
  if(application){const file=path.join(root,'.package/chart-document.sheet.json');await application.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('[data-tab="파일"]').click();await page.locator('[data-action="save-json"]').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');json=await fs.readFile(file);}
  else{await page.locator('[data-tab="파일"]').click();const event=page.waitForEvent('download');await page.locator('[data-action="save-json"]').click();json=await fs.readFile(await (await event).path());}
  const doc=JSON.parse(json.toString());assert.deepEqual(doc.sheets[0].charts[0],saved);
  await page.locator('#file-input').setInputFiles({name:'다시.sheet.json',mimeType:'application/json',buffer:json});await page.waitForFunction(()=>document.querySelector('.chart-card h3')?.textContent==='분기 매출');assert.deepEqual(await chart(),saved);
  delete doc.sheets[0].charts[0].width;delete doc.sheets[0].charts[0].height;
  await page.locator('#file-input').setInputFiles({name:'이전.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(doc))});
  await page.waitForFunction(()=>document.querySelector('.chart-card')?.style.width==='500px');
  await drag('se',100,40);assert.equal((await chart()).width,600);assert.equal((await chart()).height,340);
  // Line charts share the same interaction layer.
  await page.locator('[data-remove-chart]').click();await go('A1:B4');await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="chart-line"]').click();
  await drag('se',60,40);assert.equal((await chart()).width,560);assert.equal(await page.locator('.chart-card polyline').count(),1);
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/charts-'+(packaged?'packaged':desktop?'desktop':'browser')+'.png')});
  assert.deepEqual(errors,[]);console.log('Chart UI passed: 8 handles, resize/move, zoom, scroll, undo/redo, Escape, keyboard, title, persistence, JSON, legacy documents, line chart.');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/charts-failed.png')}).catch(()=>{});console.error('Chart test viewport',await page.evaluate(()=>({width:innerWidth,height:innerHeight,scrollLeft:document.querySelector('#grid')?.scrollLeft,scrollTop:document.querySelector('#grid')?.scrollTop})).catch(()=>null));throw error;}finally{if(application)await application.close();if(browser)await browser.close();}
