import {chromium} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import http from 'node:http';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),dist=path.join(root,'dist'),baseline=process.argv.includes('--baseline'),errors=[];
const server=http.createServer(async(req,res)=>{try{const name=new URL(req.url,'http://local').pathname,file=path.resolve(dist,'.'+(name==='/'?'/index.html':name));if(!file.startsWith(dist+path.sep))throw Error();res.setHeader('Content-Type',file.endsWith('.js')?'text/javascript':file.endsWith('.css')?'text/css':file.endsWith('.json')?'application/json':'text/html');res.end(await fs.readFile(file));}catch{res.statusCode=404;res.end();}});
await new Promise(r=>server.listen(0,'127.0.0.1',r));const browser=await chromium.launch({channel:'chrome',headless:true});let context,page;
const idle=()=>page.evaluate(()=>sheetStudio.whenIdle());
async function boot(android=true){
 await context?.close();context=await browser.newContext({viewport:{width:1280,height:800},hasTouch:android,isMobile:android,userAgent:android?'Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/142.0.0.0 Mobile Safari/537.36':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/142.0.0.0 Safari/537.36'});
 await context.addInitScript(()=>Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async()=>{},readText:async()=>''}}));
 page=await context.newPage();page.on('pageerror',e=>errors.push(e.message));await page.goto('http://127.0.0.1:'+server.address().port);await page.waitForFunction(()=>!!window.sheetStudio);await idle();
 if(android){await page.evaluate(()=>window.dispatchEvent(new CustomEvent('sheet-studio-android-input',{detail:{hardwareKeyboard:true}})));await page.waitForFunction(()=>document.body.classList.contains('android-desktop-input'));}
 const book=createWorkbook();book.name='Number width checks';const s=book.sheets[0];s.columnWidths={0:64,1:170,2:90,3:0};s.rowHeights={10:0};
 s.cells={A1:{raw:'123456789',style:{format:'number',decimals:0}},A2:{raw:'1234567',style:{format:'currency'}},A3:{raw:'12345.6789'},A4:{raw:'123456789012'},A5:{raw:'45678',style:{numberFormat:'yyyy-mm-dd'}},A6:{raw:'=A1*10',style:{numberFormat:'#,##0'}},A7:{raw:"'123456789"},A8:{raw:'123456789',style:{numberFormat:'@'}},A9:{raw:'-123456789',style:{numberFormat:'0',fontSize:18,bold:true,italic:true,align:'left',wrap:true}},A10:{raw:'-0.000000123456789'},B1:{raw:'123456789',style:{numberFormat:'0'}}};
 // Realistic sparse large worksheet: all million row positions still participate in layout.
 for(let i=100;i<100100;i++)s.cells['C'+i]={raw:String(i)};
 await page.locator('#file-input').setInputFiles({name:'width.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Number width checks'&&!document.querySelector('.workbook-busy'));await idle();
}
const cell=r=>page.locator(`.cell[data-r="${r}"][data-c="0"]`);
const action=async name=>{await page.evaluate(name=>{const b=document.createElement('button');b.dataset.action=name;document.body.append(b);b.click();b.remove();},name);await idle();};
async function start(axis,index){const b=await page.locator(`[data-resize-${axis}="${index}"]`).boundingBox(),p={x:b.x+b.width/2,y:b.y+b.height/2};await page.mouse.move(p.x,p.y);await page.mouse.down();return p;}
async function finish(p,dx=0,dy=0){await page.mouse.move(p.x+dx,p.y+dy);await page.mouse.up();await idle();}
try{
 await fs.mkdir(path.join(root,'test-results/android-number-width'),{recursive:true});await boot();
 const narrow=await cell(0).evaluate(el=>{const range=document.createRange();range.selectNodeContents(el);const b=range.getBoundingClientRect(),r=el.getBoundingClientRect();return {text:el.textContent,textLeft:b.left,cellLeft:r.left,hiddenLeadingPixels:Math.max(0,r.left-b.left),aria:el.getAttribute('aria-label')};});
 const p=await start('col',0);
 await page.evaluate(()=>{window.widthRenders=0;window.widthObserver=new MutationObserver(records=>widthRenders+=records.length);widthObserver.observe(document.querySelector('#cells'),{childList:true});});
 const drag=await page.evaluate(p=>{const start=performance.now();for(let n=1;n<=60;n++)document.dispatchEvent(new PointerEvent('pointermove',{pointerId:1,clientX:p.x+n,clientY:p.y,bubbles:true}));return {elapsedMs:performance.now()-start,width:sheetStudio.getWorkbook().sheets[0].columnWidths[0]};},p);
 await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));drag.gridRedraws=await page.evaluate(()=>widthRenders);drag.guide=await page.locator('.dimension-guide-label').textContent();await page.evaluate(()=>widthObserver.disconnect());
 console.log(JSON.stringify({baseline,narrow,drag}));
 if(baseline){assert.ok(narrow.hiddenLeadingPixels>1);assert.equal(drag.width,124);await finish(p,60);}
 else{
  assert.match(narrow.text,/^#+$/);assert.equal(narrow.hiddenLeadingPixels,0);assert.ok(narrow.aria.includes('123,456,789'));assert.equal(drag.width,64);assert.equal(drag.gridRedraws,0);assert.equal(drag.guide,'열 너비: 124px');assert.ok(drag.elapsedMs<100,'60 moves must not rebuild a million-row layout');
  await finish(p,60);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].columnWidths[0]),124);assert.equal(await cell(0).textContent(),'123,456,789');await action('undo');assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].columnWidths[0]),64);await action('redo');assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].columnWidths[0]),124);await action('undo');
  for(const r of [1,4,5,8])assert.match(await cell(r).textContent(),/^#+$/);
  for(const r of [2,3,9]){const text=await cell(r).textContent();assert.ok(/E[+-]\d{2,}$/.test(text)||/^\d+(\.\d+)?$/.test(text),text);}
  assert.equal(await cell(6).textContent(),'123456789');assert.equal(await cell(7).textContent(),'123456789');
  await page.evaluate(()=>sheetStudio.select(0,0));await idle();assert.equal(await page.locator('#formula').inputValue(),'123456789');assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'123456789');assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('A6')),1234567890);
  // Escape, blur, lost capture, pointer cancellation and no-op cannot leave model/history mutations.
  const before=await page.evaluate(()=>({history:sheetStudio.historyStats(),width:sheetStudio.getWorkbook().sheets[0].columnWidths[0]}));
  for(const cancel of ['escape','blur','capture','pointercancel']){const p=await start('col',0);await page.mouse.move(p.x+120,p.y);if(cancel==='escape')await page.keyboard.press('Escape');else await page.evaluate(kind=>kind==='blur'?window.dispatchEvent(new Event('blur')):kind==='capture'?document.querySelector('#grid-shell').releasePointerCapture(1):document.querySelector('#grid-shell').dispatchEvent(new PointerEvent('pointercancel',{pointerId:1})),cancel);await page.mouse.up();await idle();assert.deepEqual(await page.evaluate(()=>({history:sheetStudio.historyStats(),width:sheetStudio.getWorkbook().sheets[0].columnWidths[0]})),before,cancel);}
  let q=await start('col',0);await page.mouse.move(q.x+30,q.y);await finish(q);assert.deepEqual(await page.evaluate(()=>({history:sheetStudio.historyStats(),width:sheetStudio.getWorkbook().sheets[0].columnWidths[0]})),before);
  q=await start('row',0);await finish(q,0,35);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].rowHeights[0]),60);await action('undo');assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].rowHeights[0]),undefined);
  // AutoFit restores the full formatted amount, and zoom still maps pointer pixels to sheet pixels.
  await page.locator('[data-resize-col="0"]').dblclick();await idle();assert.equal(await cell(1).textContent(),'₩1,234,567.00');await action('undo');
  await page.locator('#zoom').evaluate(el=>{el.value='150';el.dispatchEvent(new Event('input',{bubbles:true}));});await idle();q=await start('col',0);await finish(q,60);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].columnWidths[0]),104);assert.equal(await cell(0).textContent(),'123,456,789');assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].columnWidths[3]),0);
  await page.screenshot({path:path.join(root,'test-results/android-number-width/tablet.png')});
  await page.evaluate(()=>sheetStudio.select(1,1));await idle();
  await action('freeze-selection');await page.locator('#grid').evaluate(el=>{el.scrollLeft=80;el.scrollTop=80;});await idle();q=await start('col',0);await finish(q,30);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].columnWidths[0]),124);q=await start('row',0);await finish(q,0,30);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].rowHeights[0]),45);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].rowHeights[10]),0);
  await action('unfreeze');await page.locator('#grid').evaluate(el=>{el.scrollLeft=0;el.scrollTop=0;});await idle();
  await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.waitForFunction(()=>!!window.sheetStudio);await idle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].columnWidths[0]),124);assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'123456789');

  const special=createWorkbook();special.name='Special numeric widths';const sp=special.sheets[0];sp.columnWidths={0:64,1:64,2:64};sp.merges=['B3:C3'];sp.autoFilter={range:{r1:0,c1:0,r2:1,c2:0},columns:{},hiddenRows:[]};sp.cells={A1:{raw:'1234567',style:{numberFormat:'#,##0'}},A2:{raw:'42'},B3:{raw:'123456789',style:{numberFormat:'0'}},A4:{raw:'0.123456789',style:{format:'percent',decimals:8}},A5:{raw:'=1/0'},A6:{raw:'=123456789'}};
  await page.locator('#file-input').setInputFiles({name:'special.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(special))});await page.getByRole('button',{name:'복구본 보관 후 계속',exact:true}).click();await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Special numeric widths'&&!document.querySelector('.workbook-busy'));await idle();
  assert.match(await cell(0).evaluate(el=>el.firstChild.textContent),/^#+$/);assert.equal(await cell(0).locator('.filter-arrow').count(),1);assert.equal(await page.locator('#android-cell-2-1').textContent(),'123456789');assert.match(await cell(3).textContent(),/^#+$/);assert.equal(await cell(4).textContent(),'#DIV/0!');assert.match(await cell(5).textContent(),/E\+08$/);
  await page.setViewportSize({width:412,height:915});await page.waitForFunction(()=>!document.body.classList.contains('android-desktop-input'));await idle();await page.locator('.android-header [data-mobile=mode]').click();await page.waitForFunction(()=>sheetStudio.isViewer());await idle();assert.match(await cell(0).evaluate(el=>el.firstChild.textContent),/^#+$/);await page.keyboard.press('Escape');await page.waitForFunction(()=>!sheetStudio.isViewer());await idle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'1234567');await page.screenshot({path:path.join(root,'test-results/android-number-width/phone.png')});
  console.log('PASS frozen/scrolled headings, recovery reload, merged numeric width, filter affordance, percentage, formula errors and phone/fullscreen numeric display');
  await boot(false);assert.equal(await cell(0).textContent(),'123,456,789');q=await start('col',0);await page.mouse.move(q.x+30,q.y);await idle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].columnWidths[0]),94);await page.mouse.up();await idle();console.log('PASS numeric widths and 100k-cell drag: General/scientific, fixed/currency/date/formula hashes, text preservation, raw/full value, deferred rows/columns, undo/redo, four cancellation paths, no-op, AutoFit, zoom, hidden columns, desktop isolation');
 }
 assert.deepEqual(errors,[]);await fs.writeFile(path.join(root,'test-results/android-number-width',baseline?'baseline.json':'verified.json'),JSON.stringify({narrow,drag},null,2));
}catch(error){await page?.screenshot({path:path.join(root,'test-results/android-number-width/failure.png')});console.error(await page?.evaluate(()=>({name:sheetStudio.getWorkbook().name,body:document.body.innerText.slice(-900)})));throw error;}finally{await context?.close();await browser.close();await new Promise(r=>server.close(r));}
