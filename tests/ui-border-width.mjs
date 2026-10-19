import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import path from 'node:path';
import fs from 'node:fs/promises';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';
// Playwright's bundled PNG decoder avoids adding a runtime dependency for screenshots.
const require=createRequire(import.meta.url),{PNG}=require(path.join(path.dirname(require.resolve('playwright-core/package.json')),'lib/utilsBundle.js'));
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=require('../package.json').version;
const dataDir=path.join(root,'.package','border-width-test-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(8000);
const settle=()=>page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const select=async range=>{await page.locator('#name-box').fill(range);await page.locator('#name-box').press('Enter');await settle();};
const load=async b=>{await page.locator('#file-input').setInputFiles({name:'borders.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(n=>sheetStudio.getWorkbook().name===n,b.name);await select('F10');};
const edge=(style='thin',color='#ff0000')=>({style,color});
const book=(style='thin')=>{const b=createWorkbook();b.name='border pixels '+style;for(const ref of ['B2','C2','B3','C3'])b.sheets[0].cells[ref]={raw:'',style:{borders:Object.fromEntries(['top','bottom','left','right'].map(side=>[side,edge(style)]))}};return b;};
const rect=ref=>page.locator(`.cell[aria-label="${ref}"]`).evaluate(el=>{const {x,y,width:w,height:h}=el.getBoundingClientRect();return {x,y,w,h};});
const screenshot=async name=>{await settle();const buffer=await page.screenshot({path:`test-results/borders-${name}.png`,scale:'css'});return PNG.sync.read(buffer);};
const red=(png,x,y)=>{const i=(Math.round(y)*png.width+Math.round(x))*4;return png.data[i]>200&&png.data[i+1]<70&&png.data[i+2]<70;};
const thickness=(png,axis,pos,along)=>{let count=0;for(let d=-5;d<=5;d++)if(axis==='h'?red(png,along,pos+d):red(png,pos+d,along))count++;return count;};
const check=async(style,width,{reload=true,name=style}={})=>{if(reload)await load(book(style));const b=await rect('B2'),c=await rect('C3'),png=await screenshot(name),x=b.x+b.w/2,y=b.y+b.h/2;
 const actual={outerTop:thickness(png,'h',b.y,x),insideH:thickness(png,'h',b.y+b.h,x),outerBottom:thickness(png,'h',c.y+c.h,x),outerLeft:thickness(png,'v',b.x,y),insideV:thickness(png,'v',b.x+b.w,y),outerRight:thickness(png,'v',c.x+c.w,y)};
 console.log(name,actual);for(const [side,n] of Object.entries(actual))assert.equal(n,width,`${style} ${side} has ${n}px instead of ${width}px`);
};
try{
 await page.locator('.cell').first().waitFor();await fs.mkdir('test-results',{recursive:true});
 await check('thin',1);await check('medium',2);await check('thick',3);
 // Exercise the real all-borders preset and its preview, rather than only fixtures.
 for(const [style,width] of [['thin',1],['medium',2],['thick',3]]){
  await load(book(style));await select('B2:C3');await page.keyboard.press(process.platform==='darwin'?'Meta+1':'Control+1');
  await page.locator('[data-border-preset=all]').click();
  const q=await page.locator('#format-border-preview').evaluate(el=>{const {x,y,width:w,height:h}=el.getBoundingClientRect();return{x,y,w,h};});
  const png=await screenshot('preview-'+style);
  for(const [axis,pos,along] of [['h',q.y,q.x+q.w/4],['h',q.y+q.h/2,q.x+q.w/4],['h',q.y+q.h,q.x+q.w/4],['v',q.x,q.y+q.h/4],['v',q.x+q.w/2,q.y+q.h/4],['v',q.x+q.w,q.y+q.h/4]])assert.equal(thickness(png,axis,pos,along),width,style+' preview '+axis);
  await page.getByRole('button',{name:'적용',exact:true}).click();await select('F10');await check(style,width,{reload:false,name:'applied-'+style});
 }
 for(const zoom of [70,150]){
  await page.locator('#zoom').evaluate((el,value)=>{el.value=value;el.dispatchEvent(new Event('input',{bubbles:true}));},zoom);await settle();
  await check('thin',1,{name:'zoom-'+zoom});
 }
 await page.locator('#zoom').evaluate(el=>{el.value=100;el.dispatchEvent(new Event('input',{bubbles:true}));});await settle();
 // Different merged and ordinary neighbours must share a single stroke.
 const merged=book('medium');merged.name='merged border pixels';merged.sheets[0].merges=['B2:C2','B3:C3'];
 for(const ref of ['D2','D3'])merged.sheets[0].cells[ref]={raw:'',style:{borders:Object.fromEntries(['top','bottom','left','right'].map(side=>[side,edge('medium')]))}};
 await load(merged);let q=await rect('B2'),png=await screenshot('merged');
 assert.equal(thickness(png,'h',q.y+q.h,q.x+q.w/2),2);assert.equal(thickness(png,'v',q.x+q.w,q.y+q.h/2),2);
 // Hidden row/column boundaries also collapse to their visible neighbours.
 const hidden=book('medium');hidden.name='hidden border pixels';hidden.sheets[0].rowHeights[2]=0;hidden.sheets[0].columnWidths[2]=0;
 for(const ref of ['B4','D2','D4'])hidden.sheets[0].cells[ref]={raw:'',style:{borders:Object.fromEntries(['top','bottom','left','right'].map(side=>[side,edge('medium')]))}};
 await load(hidden);q=await rect('B2');png=await screenshot('hidden');assert.equal(thickness(png,'h',q.y+q.h,q.x+q.w/2),2);assert.equal(thickness(png,'v',q.x+q.w,q.y+q.h/2),2);
 // Hiding gridlines must leave user borders visible.
 await page.locator('[data-tab="보기"]').click();await page.locator('[data-action=toggle-grid]').click();await load(book());png=await screenshot('no-grid');q=await rect('B2');assert.equal(thickness(png,'h',q.y+q.h,q.x+q.w/2),1);
 assert.ok(await page.locator('.cell-border').evaluateAll(els=>els.every(el=>getComputedStyle(el).pointerEvents==='none')));
 // Bottom/right-only borders remain intact while scrolling their owner past a header.
 await page.locator('[data-action=toggle-grid]').click();
 const one=createWorkbook();one.name='offscreen one-sided borders';one.sheets[0].cells.B2={raw:'',style:{borders:{bottom:edge('medium'),right:edge('medium')}}};await load(one);
 await page.locator('#grid').evaluate(el=>{el.scrollTop=45;el.scrollLeft=0;});await settle();q=await rect('B2');png=await screenshot('scroll-horizontal');assert.equal(thickness(png,'h',q.y+q.h,q.x+q.w/2),2);
 await page.locator('#grid').evaluate(el=>{el.scrollTop=0;el.scrollLeft=175;});await settle();q=await rect('B2');png=await screenshot('scroll-vertical');assert.equal(thickness(png,'v',q.x+q.w,q.y+q.h/2),2);
 // Frozen styled cells remain visible after scrolling, without duplicate strokes.
 const frozen=book('medium');frozen.name='frozen border pixels';frozen.sheets[0].freeze={rows:3,cols:3};await load(frozen);
 await page.locator('#grid').evaluate(el=>{el.scrollTop=200;el.scrollLeft=200;});await settle();await check('medium',2,{reload:false,name:'frozen'});
 assert.deepEqual(errors,[]);console.log('PASS: border pixel widths, presets/preview, zoom, merged/hidden cells, grid visibility, scroll and freeze');
}finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
