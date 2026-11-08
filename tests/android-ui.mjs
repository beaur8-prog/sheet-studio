import {chromium} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import http from 'node:http';
import {createWorkbook,address} from '../src/engine.mjs';
import {applyValidationChange,planValidationChange} from '../src/data-validation.mjs';
import {readSpreadsheet} from '../src/io.mjs';
const root=path.resolve(import.meta.dirname,'..'),dist=path.join(root,'dist');
const mime={'.html':'text/html','.js':'text/javascript','.css':'text/css','.json':'application/json','.webmanifest':'application/manifest+json','.svg':'image/svg+xml'};
const server=http.createServer(async(req,res)=>{try{if(req.url==='/__test__/recovery-store.mjs'){res.setHeader('Content-Type','text/javascript');res.end(await fs.readFile(path.join(root,'src/recovery-store.mjs')));return;}const requested=decodeURIComponent(new URL(req.url,'http://local').pathname),file=path.resolve(dist,'.'+(requested==='/'?'/index.html':requested));if(!file.startsWith(dist+path.sep))throw Error();res.setHeader('Content-Type',mime[path.extname(file)]||'application/octet-stream');res.end(await fs.readFile(file));}catch{res.statusCode=404;res.end();}});
await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
const url='http://127.0.0.1:'+server.address().port;
const browser=await chromium.launch({channel:'chrome',headless:true});
const errors=[];let context,page;
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const boot=async(options={},setup)=>{
 await context?.close();context=await browser.newContext({viewport:{width:412,height:915},hasTouch:true,isMobile:true,userAgent:'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 Chrome/142.0.0.0 Mobile Safari/537.36',...options});
 await context.addInitScript(()=>{
  window.testClipboard='';Object.defineProperty(navigator,'clipboard',{value:{writeText:async t=>window.testClipboard=t,write:async items=>window.testClipboard=await(await items[0].getType('text/plain')).text(),readText:async()=>window.testClipboard},configurable:true});
  window.sharedFiles=[];navigator.canShare=()=>true;navigator.share=async({files})=>{window.sharedFiles.push({name:files[0].name,bytes:Array.from(new Uint8Array(await files[0].arrayBuffer()))});};
 });
 if(setup)await context.addInitScript(setup);
 page=await context.newPage();page.setDefaultTimeout(8000);page.on('pageerror',e=>errors.push(e.message));await page.goto(url);await page.waitForFunction(()=>!!window.sheetStudio,null,{timeout:10000});await settle();
};
const load=async(book,name='android-test.sheet.json')=>{await page.locator('#file-input').setInputFiles({name,mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,book.name);await settle();};
const command=async name=>{await page.evaluate(name=>{const el=document.createElement('button');el.dataset.action=name;document.body.append(el);el.click();el.remove();},name);await settle();};
const mobile=async name=>{if(!await page.locator('[data-mobile="'+name+'"]').filter({visible:true}).count()&&['share','mode','find','sheets','view','address','detail'].includes(name))await mobile('more');await page.locator('[data-mobile="'+name+'"]').filter({visible:true}).first().click();await settle();};
const close=async()=>{await page.locator('.backdrop [data-close]').click();await settle();};
const select=async text=>{await mobile('address');await page.locator('#android-address').fill(text);await page.getByRole('button',{name:'이동',exact:true}).click();await settle();};
const workbook=()=>page.evaluate(()=>sheetStudio.getWorkbook());
const snapshot=()=>page.evaluate(()=>({book:sheetStudio.getWorkbook(),history:sheetStudio.historyStats(),dirty:sheetStudio.isDirty()}));
try{
 await boot();assert.equal(await page.evaluate(()=>sheetStudio.getMode()),'edit');
 let b=createWorkbook();b.name='Android viewer';b.sheets[0].cells={A1:{raw:'제목'},A2:{raw:'10'},A3:{raw:'20'},B2:{raw:'=SUM(A2:A3)'},C1:{raw:'긴 메모',note:'확인 메모'},D1:{raw:'=FutureFunction()',cachedFormula:'=FutureFunction()',cachedValue:42}};
 const other=structuredClone(createWorkbook().sheets[0]);other.id='second';other.name='두 번째';other.cells.A1={raw:'other'};b.sheets.push(other);
 b.sheets[0].charts=[{id:'chart-1',type:'bar',range:'A1:A3',title:'보기 차트',x:200,y:100,width:360,height:230}];
 await load(b);await mobile('mode');await page.locator('[data-chart="chart-1"]').tap({position:{x:50,y:50}});await page.getByRole('dialog').waitFor();assert.equal(await page.locator('.android-chart-data table').count(),1);await close();const original=await snapshot();assert.equal(original.dirty,false);assert.equal(original.history.undo,0);assert.equal(await page.evaluate(async()=>{const db=await new Promise(resolve=>{const r=indexedDB.open('sheet-studio-android-recovery',1);r.onsuccess=()=>resolve(r.result);});return new Promise(resolve=>{const r=db.transaction('state').objectStore('state').getAll();r.onsuccess=()=>{db.close();resolve(r.result.length);};});}),0);
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
  await boot({userAgent,isMobile:false});assert.equal(await page.locator('.android-header').count(),0);assert.equal(await page.locator('link[rel=manifest]').count(),0);assert.equal(await page.locator('meta[name=viewport]').getAttribute('content').then(value=>value.includes('viewport-fit')),false);assert.equal(await page.evaluate(()=>document.body.classList.contains('sheet-android')),false);assert.equal(await page.evaluate(()=>sheetStudio.getMode()),'edit');
 }
 await boot({viewport:{width:1280,height:800},isMobile:false});await mobile('more');await mobile('ribbon');assert.equal(await page.locator('.tabs').isVisible(),true);
 await page.locator('#grid').focus();await page.keyboard.press('ArrowDown');await page.locator('#editor').dispatchEvent('compositionstart',{data:''});await page.locator('#editor').fill('한글');await page.locator('#editor').dispatchEvent('input');await page.locator('#editor').dispatchEvent('compositionend',{data:'한글'});await page.locator('#editor').press('Enter');await settle();assert.equal((await workbook()).sheets[0].cells.A2.raw,'한글');
 await page.screenshot({path:path.join(root,'test-results/android-ui/tablet.png')});
 console.log('PASS platform isolation: narrow Windows/iPad unchanged; Android tablet ribbon and native composition event path');
 // A phone tap enters the stable formula field. Enter keeps that exact field
 // focused while advancing, rather than blurring to the grid between cells.
 await boot();assert.equal(await page.evaluate(()=>document.activeElement.id),'grid');
 await page.locator('#android-cell-0-0').tap();await page.waitForFunction(()=>document.activeElement.id==='formula');
 await page.evaluate(()=>{window.entryInput=document.querySelector('#formula');window.entryBlurs=0;entryInput.addEventListener('blur',()=>entryBlurs++);});
 for(const [index,value] of ['12','34','56'].entries()){
  await page.locator('#formula').fill(value);await page.locator('#formula').press('Enter');await settle();
  assert.equal(await page.evaluate(()=>document.activeElement===entryInput),true);assert.equal(await page.evaluate(()=>entryBlurs),0);assert.equal(await page.evaluate(()=>sheetStudio.getSelection().r),index+1);
 }
 assert.deepEqual(await page.evaluate(()=>['A1','A2','A3'].map(key=>sheetStudio.getCell(key).raw)),['12','34','56']);
 await page.locator('.android-entry-tools [data-mobile=keyboard]').tap();assert.equal(await page.locator('#formula').getAttribute('inputmode'),'decimal');
 await page.locator('.android-entry-tools [data-mobile=keyboard]').tap();assert.equal(await page.locator('#formula').getAttribute('inputmode'),'text');assert.equal(await page.evaluate(()=>entryBlurs),0);
 await page.locator('#formula').dispatchEvent('compositionstart',{data:''});await page.locator('#formula').fill('한글');await page.locator('#formula').press('Enter');assert.equal(await page.evaluate(()=>sheetStudio.getSelection().r),3);
 await page.locator('#formula').dispatchEvent('compositionend',{data:'한글'});await page.locator('#formula').press('Enter');await settle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('A4').raw),'한글');assert.equal(await page.evaluate(()=>entryBlurs),0);
 await page.locator('.android-entry-tools [data-mobile=next-cell]').tap();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getSelection().r),5);assert.equal(await page.evaluate(()=>entryBlurs),0);
 await page.locator('[data-action=commit-edit]').click();await settle();await mobile('mode');await page.locator('#android-cell-0-0').tap();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getMode()),'view');assert.equal(await page.locator('.formula-bar').isVisible(),false);
 console.log('PASS direct entry: single tap, 3 numeric Enter commits, same focused input, text/numeric toggle, Korean composition, next button, read-only view');

 await boot();const aggregate=createWorkbook();aggregate.name='Aggregate shortcuts';aggregate.sheets[0].cells={A1:{raw:'1'},B1:{raw:'2'},A2:{raw:'3'},B2:{raw:'4'}};await load(aggregate);
 await select('A1:B2');await mobile('aggregate-sum');assert.equal(await page.locator('#android-aggregate-range').inputValue(),'A1:B2');assert.equal(await page.locator('#android-aggregate-target').inputValue(),'A3');
 await page.getByRole('button',{name:'수식 넣기',exact:true}).click();assert.equal(await page.locator('#formula').inputValue(),'=SUM(A1:B2)');await page.locator('#formula').press('Enter');await settle();assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('A3')),10);assert.equal(await page.evaluate(()=>document.activeElement.id),'formula');
 await mobile('aggregate-average');await page.locator('#android-aggregate-range').fill('A1:B2');await page.locator('#android-aggregate-target').fill('A3');await page.getByRole('button',{name:'수식 넣기',exact:true}).click();assert.match(await page.locator('#android-aggregate-error').innerText(),/내용/);assert.equal(await page.evaluate(()=>sheetStudio.getCell('A3').raw),'=SUM(A1:B2)');
 await page.locator('#android-aggregate-target').fill('C1');await page.getByRole('button',{name:'수식 넣기',exact:true}).click();await page.locator('#formula').press('Enter');await settle();assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('C1')),2.5);
 await mobile('function');await page.locator('#function-search').fill('SUM');await page.locator('#function-list button').filter({has:page.getByText('SUM',{exact:true})}).click();await page.getByRole('button',{name:'수식에 삽입',exact:true}).click();assert.match(await page.locator('#formula').inputValue(),/^=SUM\(/);await page.locator('[data-action=cancel-edit]').click();await settle();
 console.log('PASS aggregate shortcuts: entire rectangle sum/average, visible range and target preview, occupied-cell protection, all-functions search');

 await boot();const validation=createWorkbook();validation.name='Continuous validation';validation.sheets[0].cells.A1={raw:'5'};
 applyValidationChange(validation.sheets[0],planValidationChange(validation,validation.sheets[0],'A1:A3',{type:'whole',operator:'between',formulae:['1','10']}));await load(validation);
 await page.locator('#android-cell-0-0').tap();await page.locator('#formula').fill('99');await page.locator('#formula').press('Enter');await settle();assert.equal(await page.evaluate(()=>sheetStudio.getSelection().r),0);assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'5');assert.equal(await page.locator('#formula').inputValue(),'99');assert.equal(await page.evaluate(()=>document.activeElement.id),'formula');
 await page.locator('#formula').fill('7');await page.keyboard.down('Enter');await page.locator('#formula').evaluate(el=>el.dispatchEvent(new InputEvent('beforeinput',{inputType:'insertLineBreak',bubbles:true,cancelable:true})));assert.equal(await page.evaluate(()=>sheetStudio.getSelection().r),1);await page.keyboard.up('Enter');await settle();
 await page.locator('#formula').fill('8');await page.locator('#formula').evaluate(el=>el.dispatchEvent(new InputEvent('beforeinput',{inputType:'insertLineBreak',bubbles:true,cancelable:true})));await settle();assert.equal(await page.evaluate(()=>sheetStudio.getSelection().r),2);assert.equal(await page.evaluate(()=>sheetStudio.getCell('A2').raw),'8');
 await page.locator('#formula').evaluate(el=>el.dispatchEvent(new InputEvent('beforeinput',{inputType:'insertLineBreak',bubbles:true,cancelable:true})));await settle();assert.equal(await page.evaluate(()=>sheetStudio.getSelection().r),3);
 console.log('PASS continuous validation: rejected input stays in place with draft and focus; beforeinput/Enter does not advance twice');

 await boot({viewport:{width:915,height:412},isMobile:false});assert.equal(await page.locator('body').evaluate(el=>el.classList.contains('android-landscape')),true);
 const idleGrid=await page.locator('#grid').boundingBox();assert.ok(idleGrid.height>=300,JSON.stringify(idleGrid));
 await page.locator('#android-cell-0-0').tap();await settle();const entryGrid=await page.locator('#grid').boundingBox(),entryBar=await page.locator('.formula-bar').boundingBox();assert.ok(entryBar.height<=50,JSON.stringify(entryBar));assert.ok(entryGrid.height>=360,JSON.stringify(entryGrid));
 await page.setViewportSize({width:915,height:200});await page.waitForFunction(()=>innerHeight===200);await settle();assert.equal(await page.evaluate(()=>document.activeElement.id),'formula');const keyboardGrid=await page.locator('#grid').boundingBox();assert.ok(keyboardGrid.height>=148,JSON.stringify(keyboardGrid));
 await page.screenshot({path:path.join(root,'test-results/android-ui/landscape-entry.png')});
 await page.setViewportSize({width:412,height:300});await page.waitForFunction(()=>innerWidth===412);await settle();assert.equal(await page.locator('body').evaluate(el=>el.classList.contains('android-landscape')),false);assert.equal(await page.evaluate(()=>document.activeElement.id),'formula');
 await boot();await mobile('aggregate-sum');await page.locator('#android-aggregate-range').focus();await page.setViewportSize({width:412,height:300});await page.waitForFunction(()=>innerHeight===300);assert.equal(await page.locator('body').evaluate(el=>el.classList.contains('android-landscape')),false);await close();await page.locator('#android-cell-0-0').tap();await page.setViewportSize({width:412,height:300});await page.waitForFunction(()=>innerHeight===300);await settle();assert.equal(await page.locator('body').evaluate(el=>el.classList.contains('android-landscape')),false);assert.ok((await page.locator('.formula-bar').boundingBox()).height<=98);assert.equal(await page.evaluate(()=>document.activeElement.id),'formula');
 console.log('PASS compact layout: 96px landscape idle chrome, 48px entry row, 152px grid with 200px keyboard-reduced viewport, portrait stays portrait');

 await boot();assert.equal(await page.locator('link[rel=manifest]').count(),1);await mobile('fullscreen');await page.waitForFunction(()=>!!document.fullscreenElement);await mobile('fullscreen');await page.waitForFunction(()=>!document.fullscreenElement);
 await boot({},()=>{window.systemUiCalls=[];window.sheetStudioAndroid={protocolVersion:1,platform:'android',setFullscreen:async options=>{systemUiCalls.push(options);window.dispatchEvent(new CustomEvent('sheet-studio-android-insets',{detail:{systemBars:{top:40,bottom:44},displayCutout:{top:18},displayCutoutRects:[{left:195,top:6,right:217,bottom:18}]}}));return {enabled:options.enabled};}};});await page.waitForFunction(()=>systemUiCalls.length===1);assert.equal(await page.evaluate(()=>systemUiCalls[0].hideStatusBar),true);assert.equal(await page.evaluate(()=>systemUiCalls[0].hideNavigationBar),true);assert.equal(await page.evaluate(()=>systemUiCalls[0].edgeToEdge),true);assert.equal(await page.locator('.android-header').evaluate(el=>getComputedStyle(el).paddingTop),'0px');assert.ok((await page.locator('.android-header').boundingBox()).height<=50);await mobile('fullscreen');assert.equal(await page.evaluate(()=>systemUiCalls.at(-1).enabled),false);assert.equal(await page.locator('.android-header').evaluate(el=>getComputedStyle(el).paddingTop),'40px');
 console.log('PASS fullscreen: actual browser entry/exit and acknowledged native status-bar request (native host simulated)');

 // A hidden status bar must yield worksheet space, not a permanent top/bottom gutter.
 await boot({viewport:{width:412,height:842},isMobile:false});
 const frame=()=>page.evaluate(()=>{const rect=selector=>{const r=document.querySelector(selector).getBoundingClientRect();return {top:r.top,bottom:r.bottom,height:r.height}};return {app:rect('#app'),header:rect('.android-header'),nav:rect('.android-nav'),grid:rect('#grid'),input:rect('.formula-bar'),paddingTop:getComputedStyle(document.querySelector('.android-header')).paddingTop,paddingBottom:getComputedStyle(document.querySelector('.android-nav')).paddingBottom};});
 const insets=detail=>page.evaluate(detail=>window.dispatchEvent(new CustomEvent('sheet-studio-android-insets',{detail})),detail);
 await insets({top:40,bottom:44,left:0,right:0});await settle();const barsFrame=await frame();assert.equal(barsFrame.app.top,0);assert.equal(barsFrame.header.top,0);assert.equal(barsFrame.nav.bottom,842);assert.equal(barsFrame.paddingTop,'40px');assert.equal(barsFrame.paddingBottom,'44px');
 await mobile('fullscreen');await page.waitForFunction(()=>document.body.classList.contains('android-fullscreen'));await settle();const fullFrame=await frame();assert.equal(fullFrame.header.top,0);assert.equal(fullFrame.nav.bottom,842);assert.equal(fullFrame.paddingTop,'0px');assert.equal(fullFrame.paddingBottom,'0px');assert.equal(fullFrame.grid.height-barsFrame.grid.height,84);
 // A late legacy message cannot put the removed bars back into the document.
 await insets({top:40,bottom:44});await settle();assert.equal((await frame()).grid.height,fullFrame.grid.height);
 const edgeBook=createWorkbook();edgeBook.name='전체 화면';edgeBook.sheets[0].cells={A1:{raw:'16'},A2:{raw:'56'},A3:{raw:'68'},A4:{raw:'140'},B1:{raw:'10'},B2:{raw:'20'},B3:{raw:'30'},B4:{raw:'40'},B5:{raw:'100'}};await load(edgeBook);await select('B5');
 await page.locator('.toast').waitFor({state:'detached'});await page.screenshot({path:path.join(root,'test-results/android-ui/full-viewport.png')});
 await insets({systemBars:{top:40,bottom:44},displayCutout:{top:18,right:24}});await settle();assert.equal((await frame()).paddingTop,'18px');assert.equal(await page.locator('#grid-shell').evaluate(el=>getComputedStyle(el).marginRight),'24px');
 await insets({consumedByWebView:true,systemBars:{top:40,bottom:44},displayCutout:{top:18,right:24}});await settle();assert.equal((await frame()).paddingTop,'0px');assert.equal(await page.locator('#grid-shell').evaluate(el=>getComputedStyle(el).marginRight),'0px');
 await mobile('fullscreen');await page.waitForFunction(()=>!document.fullscreenElement);await insets({top:22});await settle();assert.equal((await frame()).paddingTop,'22px');assert.equal((await frame()).paddingBottom,'0px');
 await insets({});await page.locator('#android-cell-0-0').tap();await page.evaluate(()=>{for(const [name,value] of Object.entries({height:300,width:412,offsetTop:120,offsetLeft:0,scale:1}))Object.defineProperty(visualViewport,name,{value,configurable:true});visualViewport.dispatchEvent(new Event('scroll'));});await page.waitForFunction(()=>document.querySelector('#app').getBoundingClientRect().top===120);await settle();const panned=await frame();assert.equal(panned.app.height,300);assert.equal(panned.input.bottom,420);assert.equal(await page.evaluate(()=>document.activeElement.id),'formula');
 console.log('PASS full viewport: 84px reclaimed from hidden bars, no outer gutters, cutouts retained, consumed insets not doubled, stale edges cleared, keyboard pan follows visible viewport');


 await boot({},()=>{window.closeRequests=0;window.sheetStudioAndroid={protocolVersion:1,platform:'android',closeApp:async()=>{closeRequests++;}};});const large=createWorkbook();large.name='100k Android viewer';for(let i=0;i<100000;i++)large.sheets[0].cells[address(Math.floor(i/10),i%10)]={raw:String(i)};
 const start=performance.now();await load(large);const elapsed=performance.now()-start;await mobile('mode');await select('J10000');assert.equal(await page.evaluate(()=>sheetStudio.getCell('J10000').raw),'99999');assert.equal(await page.evaluate(()=>sheetStudio.isDirty()),false);
 await page.locator('#file-input').setInputFiles({name:'broken.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from('invalid zip')});await settle();assert.equal((await page.evaluate(()=>sheetStudio.getWorkbook().name)),'100k Android viewer');assert.equal(await page.evaluate(()=>sheetStudio.getCell('J10000').raw),'99999');
 await page.evaluate(()=>window.dispatchEvent(new CustomEvent('sheet-studio-android-back',{cancelable:true})));await page.getByRole('dialog',{name:'앱을 닫을까요?'}).waitFor();await page.getByRole('button',{name:'취소',exact:true}).click();assert.equal(await page.evaluate(()=>closeRequests),0);assert.equal((await workbook()).name,'100k Android viewer');await page.evaluate(()=>window.dispatchEvent(new CustomEvent('sheet-studio-android-back',{cancelable:true})));await page.getByRole('button',{name:'닫기',exact:true}).click();await page.waitForFunction(()=>closeRequests===1);assert.equal((await workbook()).name,'100k Android viewer');
 console.log('PASS large viewer/failure/back: 100,000 cells in '+Math.round(elapsed)+'ms on desktop Chrome; malformed import preserves the open document');


 await boot({},()=>{window.closeRequests=0;window.closeFails=true;window.sheetStudioAndroid={protocolVersion:1,platform:'android',closeApp:async()=>{closeRequests++;if(closeFails)throw Error('test host failed');}};});
 const preview=createWorkbook();preview.name='Live draft';preview.sheets[0].cells.A1={raw:'16'};await load(preview);const originalPreview=await snapshot();
 await page.locator('#android-cell-0-0').tap();await page.locator('#formula').fill('123');await page.waitForFunction(()=>document.getElementById('android-cell-0-0').textContent==='123');assert.deepEqual(await snapshot(),originalPreview);assert.equal(await page.evaluate(()=>document.activeElement.id),'formula');
 await page.locator('#formula').fill('<b>초안</b>');await page.waitForFunction(()=>document.getElementById('android-cell-0-0').textContent==='<b>초안</b>');assert.equal(await page.locator('#android-cell-0-0 b').count(),0);await page.locator('[data-action=cancel-edit]').click();await page.waitForFunction(()=>document.getElementById('android-cell-0-0').textContent==='16');assert.deepEqual(await snapshot(),originalPreview);
 await page.locator('#android-cell-0-0').tap();await page.locator('#formula').fill('123');await page.locator('#formula').press('Enter');await page.locator('[data-action=commit-edit]').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'123');const committedPreview=await snapshot();
 await page.evaluate(()=>window.dispatchEvent(new Event('sheet-studio-android-back',{cancelable:true})));await page.getByRole('dialog',{name:'앱을 닫을까요?'}).waitFor();await page.getByRole('button',{name:'닫기',exact:true}).click();await page.waitForFunction(()=>closeRequests===1);await page.waitForFunction(()=>!document.querySelector('[data-modal-button="0"]').disabled);assert.deepEqual(await snapshot(),committedPreview);assert.match(await page.locator('.toast').innerText(),/닫지 못했습니다/);
 await page.evaluate(()=>{closeFails=false;});await page.getByRole('button',{name:'닫기',exact:true}).click();await page.waitForFunction(()=>closeRequests===2);assert.deepEqual(await snapshot(),committedPreview);
 console.log('PASS live draft and exit: cell preview without model/history mutation, escaped Korean text, cancel restores, close failure retains document, retry calls native close');
 assert.deepEqual(errors,[]);
}catch(error){console.error('BROWSER ERRORS',errors);await fs.mkdir(path.join(root,'test-results/android-ui'),{recursive:true});await page?.screenshot({path:path.join(root,'test-results/android-ui/failure.png')});throw error;}finally{await context?.close();await browser.close();await new Promise(resolve=>server.close(resolve));}
