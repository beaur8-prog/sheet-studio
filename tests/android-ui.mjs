import {chromium} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import http from 'node:http';
import {createWorkbook,address} from '../src/engine.mjs';
import {readSpreadsheet} from '../src/io.mjs';
const root=path.resolve(import.meta.dirname,'..'),dist=path.join(root,'dist');
const mime={'.html':'text/html','.js':'text/javascript','.css':'text/css','.json':'application/json'};
const server=http.createServer(async(req,res)=>{try{if(req.url==='/__test__/recovery-store.mjs'){res.setHeader('Content-Type','text/javascript');res.end(await fs.readFile(path.join(root,'src/recovery-store.mjs')));return;}const requested=decodeURIComponent(new URL(req.url,'http://local').pathname),file=path.resolve(dist,'.'+(requested==='/'?'/index.html':requested));if(!file.startsWith(dist+path.sep))throw Error();res.setHeader('Content-Type',mime[path.extname(file)]||'application/octet-stream');res.end(await fs.readFile(file));}catch{res.statusCode=404;res.end();}});
await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
const url='http://127.0.0.1:'+server.address().port;
const browser=await chromium.launch({channel:'chrome',headless:true});
const errors=[];let context,page;
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const boot=async(options={})=>{
 await context?.close();context=await browser.newContext({viewport:{width:412,height:915},hasTouch:true,isMobile:true,userAgent:'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 Chrome/142.0.0.0 Mobile Safari/537.36',...options});
 await context.addInitScript(()=>{
  window.testClipboard='';Object.defineProperty(navigator,'clipboard',{value:{writeText:async t=>window.testClipboard=t,write:async items=>window.testClipboard=await(await items[0].getType('text/plain')).text(),readText:async()=>window.testClipboard},configurable:true});
  window.sharedFiles=[];navigator.canShare=()=>true;navigator.share=async({files})=>{window.sharedFiles.push({name:files[0].name,bytes:Array.from(new Uint8Array(await files[0].arrayBuffer()))});};
 });
 page=await context.newPage();page.setDefaultTimeout(8000);page.on('pageerror',e=>errors.push(e.message));await page.goto(url);await page.waitForFunction(()=>!!window.sheetStudio,null,{timeout:10000});await settle();
};
const load=async(book,name='android-test.sheet.json')=>{await page.locator('#file-input').setInputFiles({name,mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,book.name);await settle();};
const command=async name=>{await page.evaluate(name=>{const el=document.createElement('button');el.dataset.action=name;document.body.append(el);el.click();el.remove();},name);await settle();};
const mobile=async name=>{await page.locator('[data-mobile="'+name+'"]').filter({visible:true}).first().click();await settle();};
const close=async()=>{await page.locator('.backdrop [data-close]').click();await settle();};
const select=async text=>{await mobile('address');await page.locator('#android-address').fill(text);await page.getByRole('button',{name:'이동',exact:true}).click();await settle();};
const workbook=()=>page.evaluate(()=>sheetStudio.getWorkbook());
const snapshot=()=>page.evaluate(()=>({book:sheetStudio.getWorkbook(),history:sheetStudio.historyStats(),dirty:sheetStudio.isDirty()}));
try{
 await boot();assert.equal(await page.evaluate(()=>sheetStudio.getMode()),'view');
 let b=createWorkbook();b.name='Android viewer';b.sheets[0].cells={A1:{raw:'제목'},A2:{raw:'10'},A3:{raw:'20'},B2:{raw:'=SUM(A2:A3)'},C1:{raw:'긴 메모',note:'확인 메모'},D1:{raw:'=FutureFunction()',cachedFormula:'=FutureFunction()',cachedValue:42}};
 const other=structuredClone(createWorkbook().sheets[0]);other.id='second';other.name='두 번째';other.cells.A1={raw:'other'};b.sheets.push(other);
 b.sheets[0].charts=[{id:'chart-1',type:'bar',range:'A1:A3',title:'보기 차트',x:200,y:100,width:360,height:230}];
 await load(b);await page.locator('[data-chart="chart-1"]').tap({position:{x:50,y:50}});await page.getByRole('dialog').waitFor();assert.equal(await page.locator('.android-chart-data table').count(),1);await close();const original=await snapshot();assert.equal(original.dirty,false);assert.equal(original.history.undo,0);assert.equal(await page.evaluate(async()=>{const db=await new Promise(resolve=>{const r=indexedDB.open('sheet-studio-android-recovery',1);r.onsuccess=()=>resolve(r.result);});return new Promise(resolve=>{const r=db.transaction('state').objectStore('state').getAll();r.onsuccess=()=>{db.close();resolve(r.result.length);};});}),0);
 await select('D1');await mobile('detail');assert.match(await page.locator('.android-detail .android-help').innerText(),/저장된 계산값/);await close();
 await mobile('sheets');await page.locator('[data-mobile-sheet="second"]').click();await mobile('sheets');await page.locator('[data-mobile-sheet="sheet-1"]').click();
 await select('A2:B3');await command('copy');await command('bold');await command('clear');await command('paste');await command('rename-sheet');
 await page.locator('#grid').focus();await page.keyboard.press('Delete');await page.keyboard.press('Control+b');await page.keyboard.press('Control+v');
 assert.deepEqual(await snapshot(),original);assert.equal(await page.evaluate(()=>document.activeElement.id),'grid');
 await mobile('find');await page.locator('#find-text').fill('제목');await page.getByRole('button',{name:'다음 찾기',exact:true}).click();await settle();assert.equal(await page.locator('#replace-text').isVisible(),false);await close();
 await page.locator('#grid').dispatchEvent('contextmenu');await page.getByRole('dialog').getByRole('button',{name:'셀 상세',exact:true}).focus();await page.keyboard.press('Enter');await page.locator('.android-detail').waitFor();await close();
 await mobile('view');await mobile('freeze-first-row');assert.deepEqual((await workbook()).sheets[0].freeze,original.book.sheets[0].freeze);
 await mobile('share');await page.getByRole('dialog').getByRole('button',{name:'공유',exact:true}).click();await settle();
 const shared=await page.evaluate(()=>window.sharedFiles);assert.equal(shared.length,1);assert.deepEqual(Buffer.from(shared[0].bytes),Buffer.from(JSON.stringify(b)));
 console.log('PASS Android viewer: original bytes, navigation/search/copy, command guards, viewing freeze, original share');
 // Browser-owned touch panning must not also run desktop range selection.
 const grid=await page.locator('#grid').boundingBox(),beforeSelection=await page.evaluate(()=>sheetStudio.getSelection());
 const cdp=await context.newCDPSession(page),x=grid.x+100,y=grid.y+Math.min(400,grid.height-70);
 await cdp.send('Input.dispatchTouchEvent',{type:'touchStart',touchPoints:[{x,y}]});
 for(let i=1;i<=8;i++)await cdp.send('Input.dispatchTouchEvent',{type:'touchMove',touchPoints:[{x,y:y-i*20}]});
 await cdp.send('Input.dispatchTouchEvent',{type:'touchEnd',touchPoints:[]});await page.waitForTimeout(200);await settle();
 assert.deepEqual(await page.evaluate(()=>sheetStudio.getSelection()),beforeSelection);assert.ok(await page.locator('#grid').evaluate(el=>el.scrollTop)>0);assert.deepEqual(await snapshot(),original);
 console.log('PASS touch swipe: scroll changes without selection or model mutations');
 await select('A1:B2');const oldRange=await page.evaluate(()=>sheetStudio.getSelection());const handle=await page.locator('[data-range-handle=end]').boundingBox();await page.mouse.move(handle.x+24,handle.y+24);await page.mouse.down();await page.mouse.move(handle.x+24,handle.y+95);await page.waitForTimeout(60);await page.setViewportSize({width:400,height:850});await page.waitForFunction(()=>innerWidth===400&&innerHeight===850);await page.waitForTimeout(60);await page.mouse.up();await settle();assert.deepEqual(await page.evaluate(()=>sheetStudio.getSelection()),oldRange);await page.setViewportSize({width:412,height:915});
 assert.equal(await page.locator('#grid').getAttribute('aria-readonly'),'true');assert.match(await page.locator('#grid').getAttribute('aria-activedescendant'),/^android-cell-/);
 console.log('PASS range/accessibility: canceled handle restores selection; logical active cell and read-only metadata');

 await mobile('mode');assert.equal(await page.evaluate(()=>sheetStudio.getMode()),'edit');
 await select('A2');await mobile('edit-cell');await page.locator('#formula').fill('한글 입력');await page.locator('#formula').dispatchEvent('input');await page.locator('[data-action="commit-edit"]').click();await settle();
 assert.equal((await workbook()).sheets[0].cells.A2.raw,'한글 입력');assert.equal(await page.evaluate(()=>sheetStudio.historyStats().undo),1);
 await command('undo');assert.equal((await workbook()).sheets[0].cells.A2.raw,'10');await command('redo');
 await select('A2:B3');await mobile('text-color');await page.locator('#palette-color').fill('#ff0000');await page.locator('#palette-color').dispatchEvent('input');await page.getByRole('button',{name:'적용',exact:true}).click();await settle();
 assert.equal((await workbook()).sheets[0].cells.B3.style.color,'#ff0000');
 await select('A1:A3');await mobile('column-filter');await page.getByRole('button',{name:'추가하고 열기',exact:true}).click();await page.locator('.filter-popup').waitFor();assert.equal(await page.locator('#filter-search').isVisible(),true);await page.locator('#filter-cancel').click();await settle();
 const edited=await snapshot();await mobile('mode');assert.equal(await page.evaluate(()=>sheetStudio.getMode()),'view');assert.deepEqual(await snapshot(),edited);
 await mobile('share');assert.match(await page.getByRole('dialog').getAttribute('aria-label'),/편집본/);await page.getByRole('dialog').getByRole('button',{name:'공유',exact:true}).click();await settle();assert.match((await page.evaluate(()=>sharedFiles.at(-1).name)),/\.xlsx$/);const exported=await page.evaluate(()=>sharedFiles.at(-1));const reopened=await readSpreadsheet(Buffer.from(exported.bytes),exported.name);assert.equal(reopened.sheets[0].cells.A2.raw,'한글 입력');assert.equal(reopened.sheets[0].cells.B3.style.color.toLowerCase(),'#ff0000');
 assert.equal(await page.evaluate(()=>sheetStudio.isDirty()),true);
 await page.evaluate(()=>sheetStudio.flushRecovery());await mobile('open');await page.getByRole('button',{name:'복구본 보관 후 계속',exact:true}).click();
 const next=createWorkbook();next.name='Other viewer';await load(next);assert.equal(await page.evaluate(()=>sheetStudio.isDirty()),false);
 await mobile('more');await mobile('recover');await page.locator('[data-recover-document]').filter({hasText:'Android viewer'}).click();await settle();assert.equal((await workbook()).sheets[0].cells.A2.raw,'한글 입력');
 console.log('PASS editing: Korean input, undo/redo, range color, edit→view→share, protected document recovery after browsing');
 // Native save responses are simulated; the file export and failure paths are real.
 await page.evaluate(()=>{window.saveAttempts=[];window.sheetStudioAndroid={protocolVersion:1,platform:'android',saveCopy:async file=>{saveAttempts.push(file.name);throw Error('provider failed');}};});
 const beforeFailure=await snapshot();await command('save');assert.deepEqual(await snapshot(),beforeFailure);assert.match(await page.locator('.toast').innerText(),/provider failed/);
 await page.evaluate(()=>{sheetStudioAndroid.saveCopy=async()=>({cancelled:true});});await command('save');assert.deepEqual(await snapshot(),beforeFailure);
 await page.evaluate(()=>{sheetStudioAndroid.saveCopy=async file=>{window.savedCopy={name:file.name,bytes:Array.from(file.data)};return {completed:true};};});await command('save');assert.equal((await page.evaluate(()=>savedCopy.name)).endsWith('.xlsx'),true);
 // A completed draft checkpoint survives renderer reload without becoming a cell edit.
 await select('C4');await mobile('edit-cell');await page.locator('#formula').fill('복구 초안');await page.locator('#formula').dispatchEvent('input');await page.evaluate(()=>sheetStudio.flushRecovery());
 await page.reload();await page.waitForFunction(()=>window.sheetStudio&&document.querySelector('#editor')?.value==='복구 초안');await settle();assert.equal((await workbook()).sheets[0].cells.C4,undefined);await page.locator('[data-action="cancel-edit"]').click();await settle();
 console.log('PASS save/draft: provider failure and cancellation preserve edits; completed cell draft survives reload');
 // A bounded archive retains each document's chunks and requires explicit removal.
 const recoveryResults=await page.evaluate(async books=>{
  const {openRecoveryStore}=await import('/__test__/recovery-store.mjs');const store=await openRecoveryStore({databaseName:'android-recovery-archive-test',retainDocuments:true,maxDocuments:2});await store.startBlank();
  for(let i=0;i<2;i++)await store.save(books[i],new Map(),new Set(books[i].sheets.map(s=>s.id)),true,'doc-'+i);
  const first=await store.recover('doc-0');let limited=false;try{await store.save(books[2],new Map(),new Set(books[2].sheets.map(s=>s.id)),true,'doc-2');}catch{limited=true;}
  const second=await store.recover('doc-1');await store.forget('doc-0');await store.save(books[2],new Map(),new Set(books[2].sheets.map(s=>s.id)),true,'doc-2');const list=await store.list();const retained=await store.recover('doc-1');store.close();return {first:first.book.name,second:second.book.name,limited,list:list.map(x=>x.name),retained:retained.book.name};
 },['first','second','third'].map(name=>{const b=createWorkbook();b.name=name;b.sheets[0].cells.A1={raw:name};return b;}));
 assert.deepEqual(recoveryResults,{first:'first',second:'second',limited:true,list:['third','second'],retained:'second'});
 console.log('PASS recovery archive: separate document generations, capacity rejection and explicit deletion');

 await fs.mkdir(path.join(root,'test-results/android-ui'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/android-ui/phone.png')});
 // Android selection is not activated on other platforms, even at phone width.
 for(const userAgent of ['Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/142.0.0.0 Safari/537.36','Mozilla/5.0 (iPad; CPU OS 18_0 like Mac OS X) AppleWebKit/605.1.15 Mobile/15E148']){
  await boot({userAgent,isMobile:false});assert.equal(await page.locator('.android-header').count(),0);assert.equal(await page.evaluate(()=>document.body.classList.contains('sheet-android')),false);assert.equal(await page.evaluate(()=>sheetStudio.getMode()),'edit');
 }
 await boot({viewport:{width:1280,height:800},isMobile:false});await mobile('mode');await mobile('more');await mobile('ribbon');assert.equal(await page.locator('.tabs').isVisible(),true);
 await page.locator('#grid').focus();await page.keyboard.press('ArrowDown');await page.locator('#editor').dispatchEvent('compositionstart',{data:''});await page.locator('#editor').fill('한글');await page.locator('#editor').dispatchEvent('input');await page.locator('#editor').dispatchEvent('compositionend',{data:'한글'});await page.locator('#editor').press('Enter');await settle();assert.equal((await workbook()).sheets[0].cells.A2.raw,'한글');
 await page.screenshot({path:path.join(root,'test-results/android-ui/tablet.png')});
 console.log('PASS platform isolation: narrow Windows/iPad unchanged; Android tablet ribbon and native composition event path');
 await boot();const large=createWorkbook();large.name='100k Android viewer';for(let i=0;i<100000;i++)large.sheets[0].cells[address(Math.floor(i/10),i%10)]={raw:String(i)};
 const start=performance.now();await load(large);const elapsed=performance.now()-start;await select('J10000');assert.equal(await page.evaluate(()=>sheetStudio.getCell('J10000').raw),'99999');assert.equal(await page.evaluate(()=>sheetStudio.isDirty()),false);
 await page.locator('#file-input').setInputFiles({name:'broken.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from('invalid zip')});await settle();assert.equal((await page.evaluate(()=>sheetStudio.getWorkbook().name)),'100k Android viewer');assert.equal(await page.evaluate(()=>sheetStudio.getCell('J10000').raw),'99999');
 await page.evaluate(()=>window.dispatchEvent(new CustomEvent('sheet-studio-android-back',{cancelable:true})));await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='통합 문서1');assert.equal(await page.evaluate(()=>sheetStudio.getMode()),'view');
 console.log('PASS large viewer/failure/back: 100,000 cells in '+Math.round(elapsed)+'ms on desktop Chrome; malformed import preserves the open document');

 assert.deepEqual(errors,[]);
}catch(error){console.error('BROWSER ERRORS',errors);await fs.mkdir(path.join(root,'test-results/android-ui'),{recursive:true});await page?.screenshot({path:path.join(root,'test-results/android-ui/failure.png')});throw error;}finally{await context?.close();await browser.close();await new Promise(resolve=>server.close(resolve));}
