import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook,address} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const dataDir=path.join(root,'.package','legend-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];await page.setViewportSize({width:1280,height:900});page.on('pageerror',e=>errors.push(e.message));
const settle=()=>page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const load=async(names,width=500,type='bar')=>{const b=createWorkbook();b.name='legend '+names.length+' '+width+' '+type;const s=b.sheets[0];names.forEach((n,c)=>{s.cells[address(0,c)]={raw:n};s.cells[address(1,c)]={raw:String(10+c)};s.cells[address(2,c)]={raw:String(20+c)};});s.charts=[{id:'legend',type,range:`A1:${address(2,names.length-1)}`,firstRowAsHeader:true,firstColumnAsLabels:false,title:'범례 가운데 정렬',x:60,y:40,width,height:300,legend:true}];await page.locator('#file-input').setInputFiles({name:'legend.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(n=>sheetStudio.getWorkbook().name===n,b.name);await settle();};
const check=async()=>{const rows=await page.locator('.chart-card>svg').evaluate(svg=>{const w=svg.viewBox.baseVal.width;return [...svg.querySelectorAll('.chart-legend-row')].map(row=>{const box=row.getBBox(),items=[...row.querySelectorAll('.chart-legend-item')].map(item=>{const r=item.getBBox(),x=item.transform.baseVal.consolidate()?.matrix.e||0;return {left:r.x+x,right:r.x+x+r.width};});return {center:box.x+box.width/2,width:box.width,expected:w/2,available:w,items};});});assert.ok(rows.length);for(const row of rows){assert.ok(Math.abs(row.center-row.expected)<.5,JSON.stringify(row));assert.ok(row.width<=row.available+.5);for(let i=1;i<row.items.length;i++)assert.ok(row.items[i].left-row.items[i-1].right>=15,JSON.stringify(row));}};
try{
 await page.locator('.cell').first().waitFor();await fs.mkdir('test-results',{recursive:true});
 await load(['매출','비용']);
 // This also measures the old ungrouped legend, so the original offset is reproducible.
 const before=await page.locator('.chart-card>svg').evaluate(svg=>{const labels=[...svg.querySelectorAll('text')].filter(el=>['매출','비용'].includes(el.textContent));const boxes=labels.map(el=>el.getBoundingClientRect()),scale=svg.getScreenCTM().a,rect=svg.getBoundingClientRect();const left=Math.min(...boxes.map(b=>b.x))-14*scale,right=Math.max(...boxes.map(b=>b.x+b.width));return {center:(left+right)/2,expected:rect.x+rect.width/2};});
 assert.ok(Math.abs(before.center-before.expected)<.5,`legend is at ${before.center}; chart center is ${before.expected}`);await check();
 await page.screenshot({path:'test-results/chart-legend-centered.png'});
 for(const width of [280,420,800])for(const type of ['bar','bar-horizontal','line','pie','doughnut']){await load(['매출','Operating expenses','아주 긴 이름을 가진 한글 계열','세금','이익','기타','마지막 계열'],width,type);await check();}
 await page.locator('#zoom').evaluate(el=>{el.value=150;el.dispatchEvent(new Event('input',{bubbles:true}));});await settle();await check();
 await page.reload();await page.locator('.chart-card').waitFor();await settle();await check();assert.deepEqual(errors,[]);
 console.log('PASS: actual legend bounds centered, Korean/Latin labels, multiple rows, narrow/wide charts, chart types, zoom and reload');
}catch(e){await page.screenshot({path:'test-results/chart-legend-failure.png'});throw e;}finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
