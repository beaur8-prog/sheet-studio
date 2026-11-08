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
async function boot(android=true,phone=false){
 await context?.close();context=await browser.newContext({viewport:phone?{width:412,height:915}:{width:1280,height:800},hasTouch:android,isMobile:android,userAgent:android?'Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/142.0.0.0 Mobile Safari/537.36':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/142.0.0.0 Safari/537.36'});
 await context.addInitScript(()=>Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async()=>{},readText:async()=>''}}));
 page=await context.newPage();page.on('pageerror',e=>errors.push(e.message));await page.goto('http://127.0.0.1:'+server.address().port);await page.waitForFunction(()=>!!window.sheetStudio);await idle();
 if(android&&!phone){await page.evaluate(()=>window.dispatchEvent(new CustomEvent('sheet-studio-android-input',{detail:{hardwareKeyboard:true}})));await page.waitForFunction(()=>document.body.classList.contains('android-desktop-input'));}
 const book=createWorkbook();book.name='Average range checks';book.sheets[0].cells={A1:{raw:'10'},A2:{raw:'20'},A3:{raw:'30'},C1:{raw:'1'},D1:{raw:'2'},E1:{raw:'6'},D3:{raw:'5'},D4:{raw:'7'},D5:{raw:'9'},H1:{raw:'1'},H2:{raw:'3'},I1:{raw:'100'},I2:{raw:'메모'}};
 await page.locator('#file-input').setInputFiles({name:'average.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Average range checks'&&!document.querySelector('.workbook-busy'));await idle();
}
const select=async(r,c,r2=r,c2=c)=>{await page.evaluate(([r,c,r2,c2])=>sheetStudio.select(r,c,r2,c2),[r,c,r2,c2]);await idle();};
const snapshot=()=>page.evaluate(()=>({book:sheetStudio.getWorkbook(),history:sheetStudio.historyStats(),dirty:sheetStudio.isDirty()}));
const action=async name=>{await page.evaluate(name=>{const b=document.createElement('button');b.dataset.action=name;document.body.append(b);b.click();b.remove();},name);await idle();};
async function average(){await page.getByRole('button',{name:'수식',exact:true}).click();await page.locator('[data-action=func-average]').click();await idle();}
async function proposal(expected){assert.equal(await page.locator('#editor').inputValue(),expected);const range=expected.slice(9,-1);assert.equal(await page.locator('#editor').evaluate(el=>el.value.slice(el.selectionStart,el.selectionEnd)),range);assert.ok(await page.locator('.formula-reference-box').count()>0);assert.equal(await page.getByRole('dialog').count(),0);}
try{
 await fs.mkdir(path.join(root,'test-results/android-average-range'),{recursive:true});await boot();await select(3,0);const before=await snapshot();await average();
 const raw=await page.locator('#editor').inputValue();await fs.writeFile(path.join(root,'test-results/android-average-range',baseline?'baseline.json':'proposal.json'),JSON.stringify({raw,range:await page.locator('#editor').evaluate(el=>el.value.slice(el.selectionStart,el.selectionEnd))},null,2));
 if(baseline){assert.equal(raw,'=AVERAGE(');console.log('REPRODUCED tablet Average opens a bare function with no suggested range');}
 else{
  await proposal('=AVERAGE(A1:A3)');await page.keyboard.press('Escape');await idle();assert.deepEqual(await snapshot(),before);
  await average();await proposal('=AVERAGE(A1:A3)');await page.locator('#editor').fill('=AVERAGE(A2:A3)');await page.keyboard.press('Enter');await idle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('A4').raw),'=AVERAGE(A2:A3)');assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('A4')),25);await action('undo');assert.equal(await page.evaluate(()=>sheetStudio.getCell('A4')),undefined);await action('redo');assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('A4')),25);
  await select(0,5);await average();await proposal('=AVERAGE(C1:E1)');await page.keyboard.press('Enter');await idle();assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('F1')),3);
  await select(2,3,4,3);await average();await proposal('=AVERAGE(D3:D5)');await page.keyboard.press('Enter');await idle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('D6').raw),'=AVERAGE(D3:D5)');assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('D6')),7);
  await select(0,7,2,9);const matrixBefore=await snapshot();await average();await proposal('=AVERAGE(H1:H2)');await page.keyboard.press('Enter');await idle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('J3').raw),'=AVERAGE(H1:I2)');assert.ok(Math.abs(await page.evaluate(()=>sheetStudio.evaluateAsync('J3'))-104/3)<1e-10);await action('undo');assert.deepEqual((await snapshot()).book,matrixBefore.book);await action('redo');assert.equal(await page.evaluate(()=>sheetStudio.getCell('J3').raw),'=AVERAGE(H1:I2)');
  await page.screenshot({path:path.join(root,'test-results/android-average-range/tablet.png')});console.log('PASS tablet Average: above/left ranges, selected rows, replaceable range highlight, cancel, Enter, undo/redo, whole-source matrix average');
  // SUM keeps its original function prefix and range-selection positions.
  await select(6,3);await action('autosum');assert.equal(await page.locator('#editor').inputValue(),'=SUM(D3:D6)');assert.equal(await page.locator('#editor').evaluate(el=>el.value.slice(el.selectionStart,el.selectionEnd)),'D3:D6');await page.keyboard.press('Escape');await idle();
  await boot(true,true);await select(3,0);await page.locator('.android-tools [data-mobile=aggregate-average]').click();await page.getByRole('dialog').waitFor();assert.equal(await page.locator('#android-aggregate-range').inputValue(),'A1:A3');assert.equal(await page.locator('#android-aggregate-target').inputValue(),'A4');assert.equal(await page.locator('#android-aggregate-preview').textContent(),'A4 ← =AVERAGE(A1:A3)');await page.getByRole('button',{name:'수식 넣기',exact:true}).click();await idle();assert.equal(await page.locator('#editor').inputValue(),'=AVERAGE(A1:A3)');await page.locator('#formula').press('Enter');await idle();assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('A4')),20);console.log('PASS phone Average keeps the suggested range dialog and Enter confirmation');
  await boot(false);await select(3,0);await average();assert.equal(await page.locator('#editor').inputValue(),'=AVERAGE(');await page.keyboard.press('Escape');console.log('PASS desktop isolation and existing SUM proposal');
 }
 assert.deepEqual(errors,[]);
}catch(error){console.error(await page?.evaluate(()=>({body:document.body.innerText.slice(-900),editor:document.querySelector('#editor')?.value})));await page?.screenshot({path:path.join(root,'test-results/android-average-range/failure.png')});throw error;}finally{await context?.close();await browser.close();await new Promise(r=>server.close(r));}
