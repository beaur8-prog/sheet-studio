import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createHash} from 'node:crypto';
const root=path.resolve(import.meta.dirname,'..'),baseline=process.argv.find(v=>v.startsWith('--compare-root='))?.slice('--compare-root='.length);
if(!baseline)throw Error('Pass --compare-root=<baseline workspace> to run alternating, isolated imports.');
const fixture=path.join(root,'.package','import-benchmark-'+Date.now()),output=path.join(root,'test-results/import-wire-benchmark.json');await fs.mkdir(fixture,{recursive:true});await fs.mkdir(path.dirname(output),{recursive:true});
const records=[],sources={};for(const [label,appRoot] of [['object',baseline],['bytes',root]]){const files=(await fs.readdir(path.join(appRoot,'dist/assets'))).filter(v=>v.startsWith('workbook-')&&v.endsWith('.js'));sources[label]=Object.fromEntries(await Promise.all(files.map(async name=>[name,createHash('sha256').update(await fs.readFile(path.join(appRoot,'dist/assets',name))).digest('hex')])));}
const median=a=>[...a].sort((a,b)=>a-b)[Math.floor(a.length/2)];
try{
 for(const kind of ['numbers','text']){
  const header=Array.from({length:100},(_,i)=>'Field '+i).join(','),row=Array(100).fill(kind==='text'?'한글 항목':'1').join(','),buffer=Buffer.from(header+'\n'+(row+'\n').repeat(9999)),sha=createHash('sha256').update(buffer).digest('hex');
  for(let pair=0;pair<5;pair++)for(const label of pair%2?['bytes','object']:['object','bytes']){
   const appRoot=label==='object'?baseline:root,profile=path.join(fixture,kind+'-'+pair+'-'+label),app=await _electron.launch(desktopLaunch(appRoot,{...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'})),page=await app.firstWindow();page.setDefaultTimeout(180000);
   let peak=0;const sample=async()=>{const size=await app.evaluate(({app})=>app.getAppMetrics().reduce((n,p)=>n+(p.memory?.workingSetSize||0),0));peak=Math.max(peak,size/1024);};
   try{
    await page.locator('.cell').first().waitFor();await page.evaluate(()=>sheetStudio.whenIdle());await sample();await page.evaluate(()=>{globalThis.__wireGaps=[];let last=performance.now();globalThis.__wireTimer=setInterval(()=>{const now=performance.now();__wireGaps.push(now-last);last=now;},20);});
    const start=performance.now();await page.locator('#file-input').setInputFiles({name:'million.csv',mimeType:'text/csv',buffer});await page.waitForFunction(()=>sheetStudio.historyStats().cells===1000000);await page.locator('.workbook-busy').waitFor({state:'detached'});await sample();await page.evaluate(()=>sheetStudio.flushRecovery());await sample();
    const ms=performance.now()-start,gaps=await page.evaluate(()=>{clearInterval(__wireTimer);return {samples:__wireGaps.length,maxGapMs:Math.max(0,...__wireGaps)};});assert.equal(await page.evaluate(()=>sheetStudio.getCell('CV10000').raw),kind==='text'?'한글 항목':'1');records.push({kind,pair,label,ms,peakObservedWorkingSetMiB:peak,inputSHA256:sha,...gaps});console.log(JSON.stringify(records.at(-1)));
   }finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
  }
 }
 const summary={};for(const kind of ['numbers','text'])for(const label of ['object','bytes']){const selected=records.filter(r=>r.kind===kind&&r.label===label);summary[kind+'-'+label]={observations:selected.length,medianMs:median(selected.map(r=>r.ms)),maxMs:Math.max(...selected.map(r=>r.ms)),medianMaxGapMs:median(selected.map(r=>r.maxGapMs)),peakObservedWorkingSetMiB:Math.max(...selected.map(r=>r.peakObservedWorkingSetMiB))};}
 await fs.writeFile(output,JSON.stringify({platform:process.platform,arch:process.arch,sources,records,summary,notes:'Five alternating cold-profile observations per transport and data kind; source Electron on the same machine; working sets sampled at stage boundaries, summed across Electron processes, may count shared memory and may miss transient peaks. Not Excel performance or a universal capacity guarantee.'},null,2));console.log(JSON.stringify(summary));
}finally{await fs.rm(fixture,{recursive:true,force:true});}
