import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import path from 'node:path';
import fs from 'node:fs/promises';
import platform from '../scripts/platform.cjs';
import {createRequire} from 'node:module';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged');
const version=createRequire(import.meta.url)('../package.json').version;
const dataDir=path.join(root,'.package','cell-format-test-'+Date.now());
const env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(8000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const cells=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells);
const load=async b=>{await page.locator('#file-input').setInputFiles({name:'format.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(n=>sheetStudio.getWorkbook().name===n,b.name);await settle();};
const select=async range=>{await page.locator('#name-box').fill(range);await page.locator('#name-box').press('Enter');await settle();};
const cell=ref=>page.locator(`.cell[aria-label="${ref}"]`);
const popup=async(ref='A1')=>{await cell(ref).click({button:'right'});await page.locator('.menu').waitFor();};
const format=async()=>{await popup();await page.locator('.menu [data-action=cell-format]').click();await page.getByRole('dialog',{name:'셀 서식',exact:true}).waitFor();};
const apply=async()=>{await page.getByRole('button',{name:'적용',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('.backdrop'));await settle();};
const preset=key=>page.locator(`[data-border-preset=${key}]`).click();
try{
 await page.locator('.cell').first().waitFor();
 const b=createWorkbook();b.name='format regression';for(const key of ['A1','A2','B1','B2','C1'])b.sheets[0].cells[key]={raw:key};b.sheets[0].cells.B2.style={italic:true,color:'#7030a0'};await load(b);
 // Pointer capture must dismiss the popup before cell rendering removes the event target.
 await popup();await cell('D2').click();await settle();assert.equal(await page.locator('.menu').count(),0);
 await popup();await page.keyboard.press('ArrowRight');await settle();assert.equal(await page.locator('.menu').count(),0);
 await popup();await page.keyboard.press('Escape');assert.equal(await page.locator('.menu').count(),0);
 await popup();await page.locator('#grid').evaluate(el=>el.scrollTop=200);await settle();assert.equal(await page.locator('.menu').count(),0);await select('A1:B2');
 await format();assert.equal(await page.getByRole('dialog').getByText('A1:B2',{exact:true}).count(),1);
 await page.locator('#format-line').selectOption('medium');await page.locator('#format-border-color').fill('#ff0000');await preset('outline');
 await page.screenshot({path:'test-results/cell-format-dialog.png'});await apply();
 let s=await cells();assert.deepEqual(s.A1.style.borders,{top:{style:'medium',color:'#ff0000'},right:null,bottom:null,left:{style:'medium',color:'#ff0000'}});assert.deepEqual(s.B2.style.borders,{top:null,right:{style:'medium',color:'#ff0000'},bottom:{style:'medium',color:'#ff0000'},left:null});assert.equal(s.B2.style.italic,true);assert.equal(s.C1.style,undefined);assert.ok(await page.locator('.cell-border').evaluateAll(els=>els.some(el=>getComputedStyle(el).borderTopWidth==='2px'&&getComputedStyle(el).borderTopColor==='rgb(255, 0, 0)')));
 await page.locator('#undo').click();await settle();s=await cells();assert.equal(s.A1.style,undefined);assert.deepEqual(s.B2.style,{italic:true,color:'#7030a0'});await page.locator('#redo').click();await settle();
 const before=JSON.stringify(await cells());await format();await preset('none');await page.getByRole('button',{name:'취소',exact:true}).click();assert.equal(JSON.stringify(await cells()),before);
 // Add inner edges without removing the existing outline; change font/alignment in the same undo step.
 await page.keyboard.press(process.platform==='darwin'?'Meta+1':'Control+1');await page.getByRole('dialog',{name:'셀 서식',exact:true}).waitFor();assert.equal(await page.locator('[data-border-side=top]').getAttribute('aria-pressed'),'true');
 await page.locator('#format-line').selectOption('dashed');await page.locator('#format-border-color').fill('#0070c0');await preset('inside');
 await page.getByRole('tab',{name:'글꼴',exact:true}).click();await page.locator('#format-size').fill('0');await page.getByRole('button',{name:'적용',exact:true}).click();assert.match(await page.locator('#format-error').textContent(),/1~4096/);await page.locator('#format-size').fill('16');await page.locator('#format-bold').check();
 await page.getByRole('tab',{name:'맞춤',exact:true}).click();await page.locator('#format-vertical').selectOption('bottom');await apply();s=await cells();assert.deepEqual(s.A1.style.borders.top,{style:'medium',color:'#ff0000'});assert.deepEqual(s.A1.style.borders.bottom,{style:'dashed',color:'#0070c0'});for(const k of ['A1','A2','B1','B2']){assert.equal(s[k].style.fontSize,16);assert.equal(s[k].style.bold,true);assert.equal(s[k].style.verticalAlign,'bottom');}assert.equal(s.B2.style.italic,true);assert.equal(s.B2.style.color,'#7030a0');
 await page.locator('#undo').click();await settle();assert.equal(JSON.stringify(await cells()),before);await page.locator('#redo').click();await settle();
 // A narrow window scrolls the ribbon to reveal Format. Its queued scroll must not close the new menu.
 await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].setSize(1024,768));await settle();
 // Home routes preserve the selection; individual edges can be removed and all edges cleared.
 await page.locator('[data-action=cell-format-menu]').click();await settle();assert.equal(await page.locator('.menu').count(),1);await page.locator('#ribbon').evaluate(el=>el.scrollLeft=0);await settle();assert.equal(await page.locator('.menu').count(),0);await page.locator('[data-action=cell-format-menu]').click();await settle();await page.locator('.menu [data-action=cell-format]').click();await page.locator('[data-border-side=bottom]').click();await apply();assert.equal((await cells()).B2.style.borders.bottom,null);
 await page.locator('[data-action=borders]').click();await preset('none');await apply();for(const k of ['A1','A2','B1','B2'])assert.ok(Object.values((await cells())[k].style.borders).every(v=>v===null));
 await page.screenshot({path:'test-results/cell-format-applied.png'});
 // Esc closes without applying; cell-format honors merged cell extents and hidden rows.
 await page.locator('[data-action=borders]').click();await preset('all');await page.keyboard.press('Escape');assert.equal(await page.locator('.backdrop').count(),0);assert.equal((await cells()).A1.style.borders.top,null);
 const merged=createWorkbook();merged.name='merged format';merged.sheets[0].cells.A1={raw:'merged'};merged.sheets[0].merges=['A1:B2'];await load(merged);await select('A1');await format();await preset('outline');await apply();assert.ok(Object.values((await cells()).A1.style.borders).every(v=>v?.style==='thin'));
 const hidden=createWorkbook();hidden.name='hidden format';hidden.sheets[0].rowHeights[0]=0;hidden.sheets[0].cells.A1={raw:'hidden'};await load(hidden);await select('A1:B2');await page.locator('[data-action=borders]').click();await preset('outline');await apply();s=await cells();assert.equal(s.A1.style,undefined);assert.ok(s.A2.style.borders.top);assert.ok(s.B2.style.borders.bottom);
 // Native focus cannot be acquired on a locked macOS session. Border/style
 // checks still run there; keep the OS window-switch check explicit and opt out
 // only when the caller has established that limitation.
 if(process.env.SHEET_STUDIO_TEST_NO_NATIVE_FOCUS==='1')console.log('SKIP: native window-switch popup dismissal (OS session cannot grant focus)');
 else {await app.evaluate(({app,BrowserWindow})=>{app.focus({steal:true});BrowserWindow.getAllWindows()[0].focus();});await page.waitForFunction(()=>document.hasFocus());await select('A2');await popup('A2');await app.evaluate(({BrowserWindow})=>{const other=new BrowserWindow({width:320,height:200,show:true});other.focus();});await page.waitForFunction(()=>!document.querySelector('.menu'));}
 assert.deepEqual(errors,[]);console.log('PASS: pointer/key/scroll popup dismissal, cell format routes, outline/inside/individual/none borders, style apply/cancel, undo/redo, merged/hidden selection');
}finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
