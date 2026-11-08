import {chromium} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import http from 'node:http';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),dist=path.join(root,'dist'),errors=[];
const server=http.createServer(async(req,res)=>{try{const name=new URL(req.url,'http://local').pathname,file=path.resolve(dist,'.'+(name==='/'?'/index.html':name));if(!file.startsWith(dist+path.sep))throw Error();res.setHeader('Content-Type',file.endsWith('.js')?'text/javascript':file.endsWith('.css')?'text/css':file.endsWith('.json')?'application/json':'text/html');res.end(await fs.readFile(file));}catch{res.statusCode=404;res.end();}});
await new Promise(r=>server.listen(0,'127.0.0.1',r));const url='http://127.0.0.1:'+server.address().port,browser=await chromium.launch({channel:'chrome',headless:true});let context,page;
const idle=()=>page.evaluate(()=>sheetStudio.whenIdle());
async function boot(android=false,viewport={width:1440,height:900}){
 await context?.close();context=await browser.newContext({viewport,hasTouch:android,isMobile:android,userAgent:android?'Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/142.0.0.0 Mobile Safari/537.36':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/142.0.0.0 Safari/537.36'});
 await context.addInitScript(()=>{const BaseWorker=window.Worker;window.__calcs=[];window.Worker=class extends BaseWorker{constructor(...args){super(...args);this.delays=new Set();}postMessage(message,transfer){if(message.type==='calculate'){window.__calcs.push(message.payload?.refs?.length||0);const timer=setTimeout(()=>{this.delays.delete(timer);super.postMessage(message,transfer);},180);this.delays.add(timer);}else super.postMessage(message,transfer);}terminate(){for(const timer of this.delays)clearTimeout(timer);super.terminate();}};window.testClip='';Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async text=>testClip=text,write:async items=>{testClip=await(await items[0].getType('text/plain')).text();},readText:async()=>testClip}});});
 page=await context.newPage();page.setDefaultTimeout(10000);page.on('pageerror',error=>errors.push(error.message));await page.goto(url);await page.waitForFunction(()=>!!window.sheetStudio);await idle();
}
const snap=()=>page.evaluate(()=>({book:sheetStudio.getWorkbook(),history:sheetStudio.historyStats(),dirty:sheetStudio.isDirty()}));
const action=async name=>{await page.evaluate(name=>{const b=document.createElement('button');b.dataset.action=name;document.body.append(b);b.click();b.remove();},name);await idle();};
const back=async()=>{await page.keyboard.press('Escape');await page.waitForFunction(()=>!sheetStudio.isViewer());await idle();};
async function load(charts=false){const book=createWorkbook();book.name='Viewer checks';if(charts)book.sheets[0].charts=[{id:'tablet-chart',type:'bar',range:'A1:B2',title:'Tablet chart',x:350,y:100,width:360,height:230}];book.sheets[0].cells={A1:{raw:'10'},B1:{raw:'20'},A2:{raw:'=SUM(A1:B1)'}};book.sheets.push({...structuredClone(book.sheets[0]),id:'second',name:'다른 시트',cells:{C3:{raw:'읽기'}}});await page.locator('#file-input').setInputFiles({name:'viewer.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Viewer checks'&&!document.querySelector('.workbook-busy'));await idle();}
async function fillScreen(){const box=await page.locator('#grid-shell').boundingBox(),height=await page.evaluate(()=>innerHeight);assert.equal(box.y,0);assert.ok(Math.abs(box.height-height)<=1,JSON.stringify({box,height}));}
try{
 await fs.mkdir(path.join(root,'test-results/android-smoothness'),{recursive:true});await boot(true,{width:412,height:915});
 const book=createWorkbook();book.name='Smooth display';for(let r=1;r<=1500;r++){book.sheets[0].cells['A'+r]={raw:String(r)};book.sheets[0].cells['B'+r]={raw:'=A'+r+'*2'};}
 await page.locator('#file-input').setInputFiles({name:'smooth.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Smooth display');await idle();
 await page.evaluate(()=>{window.__frames=[];window.__sample=true;function sample(){if(!__sample)return;const cells=[...document.querySelectorAll('#cells .cell')];__frames.push({calculating:cells.filter(c=>c.textContent.includes('계산 중')).length,b1:document.querySelector('#android-cell-0-1')?.textContent,a1:document.querySelector('#android-cell-0-0')?.textContent});requestAnimationFrame(sample);}requestAnimationFrame(sample);});
 await page.locator('#android-cell-0-0').tap();await page.locator('#formula').fill('10');await page.locator('#formula').press('Enter');await page.waitForTimeout(300);await idle();
 assert.equal(await page.locator('#android-cell-0-1').innerText(),'20');
 for(const top of [1250,5000,17500,27500,0]){await page.locator('#grid').evaluate((el,top)=>el.scrollTop=top,top);await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));}
 await idle();const frames=await page.evaluate(()=>{__sample=false;return __frames;});await fs.writeFile(path.join(root,'test-results/android-smoothness/frames.json'),JSON.stringify(frames,null,2));assert.ok(frames.length>5);assert.equal(frames.reduce((n,f)=>n+f.calculating,0),0,'No visible cell should become a calculation placeholder');assert.equal(frames.some(f=>f.b1!==undefined&&!['2','20'].includes(f.b1)),false,'Previously calculated B1 should retain its last result while recalculating');
 await page.screenshot({path:path.join(root,'test-results/android-smoothness/after-scroll.png')});assert.deepEqual(errors,[]);console.log('PASS delayed worker display: '+frames.length+' frames; numeric entry/Enter and scroll never show placeholders or erase known formula results');
}catch(error){console.error(error);await page?.screenshot({path:path.join(root,'test-results/android-smoothness/failure.png')});throw error;}finally{await context?.close();await browser.close();await new Promise(r=>server.close(r));}
