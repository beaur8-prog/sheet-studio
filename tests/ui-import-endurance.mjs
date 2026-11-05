// Optional repeated-use check. This runs only against generated files and an
// isolated profile; no OS clipboard, user document or user recovery is touched.
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {desktopLaunch} from './electron-launch.mjs';
import {createWorkbook,address} from '../src/engine.mjs';
import {writeSpreadsheet} from '../src/io.mjs';
import {assertBlankStart,recoverPrevious} from './recovery-ui.mjs';

const root=path.resolve(import.meta.dirname,'..'),million=process.argv.includes('--million'),gcProbe=process.argv.includes('--gc-probe'),count=million?1000000:100000,cols=100,rows=count/cols,cycles=20,profile=path.join(root,'.package','import-endurance-'+Date.now()),kind=(million?'million':'100k')+(gcProbe?'-gc':''),output=path.join(root,'test-results/import-endurance-'+kind+'.json');
await fs.mkdir(profile,{recursive:true});await fs.mkdir(path.dirname(output),{recursive:true});
const input=path.join(profile,'mixed-input.xlsx'),saved=path.join(profile,'edited.xlsx'),sha=bytes=>createHash('sha256').update(bytes).digest('hex'),records=[],errors=[];
let fixture=createWorkbook();fixture.name='Repeated mixed input';let sheet=fixture.sheets[0];
for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){
 const i=r*cols+c,raw=r===0?'Field '+(c+1):c===0?String(r+1):c===1?'=A'+(r+1)+'*2':i%19===0?'한글 항목 '+r:i%23===0?"'00012":String((r+c)%997);
 sheet.cells[address(r,c)]={raw,...(i%97===0?{style:{color:'#224466',numberFormat:'0.00'}}:{})};
}
const lastRaw=sheet.cells[address(rows-1,cols-1)].raw;await fs.writeFile(input,(await writeSpreadsheet(fixture)).data);fixture=null;sheet=null;globalThis.gc?.();const sourceHash=sha(await fs.readFile(input));
let app,page;
async function launch(){const config=desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'});if(gcProbe)config.args.push('--js-flags=--expose-gc');app=await _electron.launch(config);page=await app.firstWindow();page.setDefaultTimeout(180000);page.on('pageerror',e=>errors.push(e.message));await page.setViewportSize({width:1280,height:900});await page.locator('.cell').first().waitFor();await page.evaluate(()=>sheetStudio.whenIdle());}
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const stats=()=>page.evaluate(()=>sheetStudio.historyStats());
async function sample(){return app.evaluate(({app})=>app.getAppMetrics().map(p=>({type:p.type,workingSetMiB:(p.memory?.workingSetSize||0)/1024,privateMiB:(p.memory?.privateBytes||0)/1024})));}
async function open(file){await page.locator('#file-input').setInputFiles(file);await page.locator('.workbook-busy').waitFor({state:'visible'});await settle();assert.equal((await stats()).cells,count);const limit=Date.now()+5000;while(page.workers().length>1&&Date.now()<limit)await page.waitForTimeout(20);assert.equal(page.workers().length,1,'only the live document worker remains');}
async function blank(){await page.keyboard.press('ControlOrMeta+n');await page.getByRole('dialog').getByRole('button',{name:'새로 만들기',exact:true}).click();await settle();assert.equal((await stats()).cells,0);}
const flush=()=>page.evaluate(()=>sheetStudio.flushRecovery());
const record=async row=>{records.push(row);await fs.writeFile(output,JSON.stringify({at:new Date().toISOString(),platform:process.platform,arch:process.arch,count,cycles,sourceHash,gcProbe,packaged:process.argv.includes('--packaged'),records,errors,notes:'Same app/profile across 20 open/edit/undo/redo/save/new/reopen/new cycles. Memory is summed process working-set observations at stage boundaries, may double-count shared pages and miss transient peaks; no memory pass threshold or universal performance claim. Optional --gc-probe exposes GC only in this isolated diagnostic app and collects renderer/document-worker/main heaps every five cycles; that run is separate from ordinary performance observations.'},null,2));console.log(JSON.stringify(row));};
async function collect(cycle){
 const session=await page.context().newCDPSession(page);
 try{
  const before={processes:await sample(),heap:await session.send('Runtime.getHeapUsage')},workers=[];
  for(const worker of page.workers())workers.push(await worker.evaluate(()=>{const available=typeof globalThis.gc==='function';if(available){gc();gc();}return {gcAvailable:available};}));
  await page.requestGC();const main=await app.evaluate(()=>{const available=typeof globalThis.gc==='function';if(available){gc();gc();}return {gcAvailable:available,heap:process.memoryUsage()};});await page.waitForTimeout(200);await record({stage:'gc-diagnostic',cycle,documentRetained:cycle===cycles,before,workers,main,after:{processes:await sample(),heap:await session.send('Runtime.getHeapUsage')}});
 }finally{await session.detach();}
}
try{
 await launch();await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},saved);const runtime=await app.evaluate(({app})=>({version:app.getVersion(),appPath:app.getAppPath()}));await record({stage:'blank-start',runtime:{...runtime,...(runtime.appPath.endsWith('.asar')?{asarSHA256:sha(await fs.readFile(runtime.appPath))}:{})},processes:await sample()});
 for(let i=0;i<cycles;i++){
  await page.evaluate(()=>{globalThis.__enduranceGaps=[];let last=performance.now();globalThis.__enduranceTimer=setInterval(()=>{const now=performance.now();__enduranceGaps.push(now-last);last=now;},20);});
  const start=performance.now();await open(input);assert.equal(await page.evaluate(()=>sheetStudio.getCell('A2').raw),'2');await page.locator('#name-box').fill('A2');await page.locator('#name-box').press('Enter');await page.keyboard.press('F2');await page.locator('#editor').fill(String(100+i));await page.locator('#editor').press('Enter');await settle();assert.equal(await page.evaluate(()=>sheetStudio.evaluate('B2')),(100+i)*2);
  await page.locator('#undo').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.evaluate('B2')),4);await page.locator('#redo').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.evaluate('B2')),(100+i)*2);await flush();
  await page.locator('#save').click();await page.locator('.workbook-busy').waitFor({state:'visible'});await settle();assert.ok((await fs.stat(saved)).size>1000);const imported=await sample();await blank();await open(saved);assert.equal(await page.evaluate(()=>sheetStudio.getCell('A2').raw),String(100+i));assert.equal(await page.evaluate(()=>sheetStudio.evaluate('B2')),(100+i)*2);assert.equal(await page.evaluate(ref=>sheetStudio.getCell(ref).raw,address(rows-1,cols-1)),lastRaw);await flush();if(i<cycles-1)await blank();assert.ok((await stats()).bytes<=128*1024*1024,'undo retention stays within its documented budget');
  const gaps=await page.evaluate(()=>{clearInterval(__enduranceTimer);return {samples:__enduranceGaps.length,maxGapMs:Math.max(0,...__enduranceGaps)};});await record({cycle:i+1,ms:performance.now()-start,...gaps,workers:page.workers().length,history:await stats(),documentRetained:i===cycles-1,afterImport:imported,afterCycle:await sample()});
  if(gcProbe&&(i+1)%5===0)await collect(i+1);
 }
 assert.equal(sha(await fs.readFile(input)),sourceHash,'input file stays unchanged');assert.deepEqual(errors,[]);await app.close();app=null;await launch();await assertBlankStart(page);await recoverPrevious(page);await settle();assert.equal((await stats()).cells,count);assert.equal(await page.evaluate(()=>sheetStudio.getCell('A2').raw),'119');assert.equal(await page.evaluate(()=>sheetStudio.evaluate('B2')),238);await page.screenshot({path:path.join(root,'test-results/import-endurance-'+kind+'.png')});await record({stage:'cold-restart-explicit-recovery',cells:(await stats()).cells,value:await page.evaluate(()=>sheetStudio.evaluate('B2')),processes:await sample()});console.log('PASS: 20 repeated mixed-document cycles, atomic undo/redo, saved values/formula/last cell, no orphan parsers, original SHA, blank startup and explicit last-document recovery.');
}catch(error){await page?.screenshot({path:path.join(root,'test-results/import-endurance-'+kind+'-failure.png')}).catch(()=>{});await record({stage:'failure',error:String(error)});throw error;}
finally{if(app)await app.close();await fs.rm(profile,{recursive:true,force:true});}
