import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {desktopLaunch} from './electron-launch.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','find-capacity-'+Date.now());
const cycles=Number(process.argv.find(a=>a.startsWith('--cycles='))?.split('=')[1]||10),longText=process.argv.includes('--text');
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:dataDir,STUDIO_DISABLE_UPDATES:'1'}));
const page=await app.firstWindow(),errors=[],findings=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(180000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const record=data=>{findings.push(data);console.log(JSON.stringify(data));};
const memory=()=>app.evaluate(({app})=>app.getAppMetrics().reduce((sum,item)=>sum+(item.memory?.workingSetSize||0),0)/1024);
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  const row=Array(100).fill(longText?'매출자료를다루는반복문자열':'1').join(','),csv='marker-0,'+row.slice(row.indexOf(',')+1)+'\n'+(row+'\n').repeat(9999);
  const started=Date.now();await page.locator('#file-input').setInputFiles({name:'find-million.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});await page.waitForFunction(()=>sheetStudio.historyStats().cells===1000000);await settle();await page.evaluate(()=>sheetStudio.flushRecovery());record({case:'million-cell import',kind:longText?'mixed Korean text':'numbers',ms:Date.now()-started,rssMiB:await memory()});
  const baselineHistory=await page.evaluate(()=>sheetStudio.historyStats().bytes);await page.locator('.search-command').click();await page.locator('#find-whole').check();
  for(let i=0;i<cycles;i++){
    await page.locator('#find-text').fill('marker-'+i);await page.locator('#replace-text').fill('marker-'+(i+1));
    await page.evaluate(()=>{globalThis.findGaps=[];let last=performance.now();globalThis.findTimer=setInterval(()=>{const now=performance.now();findGaps.push(now-last);last=now;},20);});
    const start=Date.now();await page.getByRole('dialog').getByRole('button',{name:'모두 찾기 / 미리 보기',exact:true}).click();await settle();assert.match(await page.locator('#find-status').innerText(),/^1개 발견/);const searched=Date.now();
    await page.getByRole('dialog').getByRole('button',{name:'모두 바꾸기',exact:true}).click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'marker-'+(i+1));
    const stats=await page.evaluate(()=>sheetStudio.historyStats()),gaps=await page.evaluate(()=>{clearInterval(findTimer);return {maxGapMs:Math.max(0,...findGaps),samples:findGaps.length};});
    assert.ok(stats.undo>0);assert.ok(stats.bytes-baselineHistory<1_000_000,JSON.stringify({baselineHistory,...stats}));record({case:'replace one cell in one million',iteration:i+1,searchMs:searched-start,replaceMs:Date.now()-searched,undo:stats.undo,historyBytes:stats.bytes,addedHistoryBytes:stats.bytes-baselineHistory,rssMiB:await memory(),...gaps});
  }
  await page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();await page.keyboard.press('ControlOrMeta+z');await settle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'marker-'+(cycles-1));await page.keyboard.press('ControlOrMeta+Shift+z');await settle();assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'marker-'+cycles);
  const recovery=await page.evaluate(()=>sheetStudio.flushRecovery());assert.ok(recovery.chunksWritten<=2);record({case:'incremental recovery and undo/redo',...recovery});
  assert.deepEqual(errors,[]);await fs.writeFile(path.join(root,'test-results',longText?'find-capacity-text.json':'find-capacity-numbers.json'),JSON.stringify(findings,null,2));
  console.log('PASS: repeated worker searches/replacements on one million cells retain small undo deltas and incremental recovery.');
}catch(error){await fs.writeFile(path.join(root,'test-results/find-capacity-failure.json'),JSON.stringify({findings,error:String(error),errors},null,2));await page.screenshot({path:path.join(root,'test-results/find-capacity-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
