import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {createWorkbook,address} from '../src/engine.mjs';
import {writeSpreadsheet} from '../src/io.mjs';
const root=path.resolve(import.meta.dirname,'..'),fixture=path.join(root,'.package','import-cancel-large-'+Date.now()),profile=path.join(fixture,'profile');await fs.mkdir(fixture,{recursive:true});await fs.mkdir(path.join(root,'test-results'),{recursive:true});
const input=createWorkbook();input.name='100k mixed input';for(let r=0;r<10000;r++)for(let c=0;c<10;c++)input.sheets[0].cells[address(r,c)]={raw:c===0?'상품 '+r:c===9?'='+address(r,1)+'*2':String(r+c),style:c===9?{format:'number',numberFormat:'0.000'}:{}};
const data=(await writeSpreadsheet(input)).data,file=path.join(fixture,'large.xlsx');await fs.writeFile(file,data);const sha=createHash('sha256').update(data).digest('hex');
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile})),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(90000);
async function parsersClosed(){const start=performance.now();while(page.workers().length>1&&performance.now()-start<5000)await new Promise(resolve=>setTimeout(resolve,20));assert.equal(page.workers().length,1,JSON.stringify(page.workers().map(w=>w.url())));return performance.now()-start;}
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
try{
  await page.locator('.cell').first().waitFor();const keeper=createWorkbook();keeper.name='keep during real parse';keeper.sheets[0].cells={A1:{raw:'이전 자료'},B1:{raw:'=21*2'}};
  await page.locator('#file-input').setInputFiles({name:'keeper.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(keeper))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='keep during real parse');await settle();await page.evaluate(()=>sheetStudio.flushRecovery());
  const original=await page.evaluate(()=>JSON.stringify(sheetStudio.getWorkbook()));await page.evaluate(()=>{
    globalThis.__parseGaps=[];let last=performance.now();globalThis.__parseTimer=setInterval(()=>{const now=performance.now();__parseGaps.push(now-last);last=now;},20);
    // A fast CI machine can finish this phase between Playwright locator polls.
    // Observe the real UI progress update and activate its real cancel button;
    // do not hold/throttle/mock the parser or change production timing.
    globalThis.__largeCancel=null;const observer=new MutationObserver(()=>{
      const overlay=document.querySelector('.workbook-busy'),button=overlay?.querySelector('[data-busy-cancel]');
      if(!__largeCancel&&overlay?.textContent.includes('XLSX 셀 데이터를 해석하는 중')&&button){__largeCancel={at:performance.now(),enabled:!button.disabled,text:overlay.textContent};button.click();}
      if(__largeCancel&&!overlay){__largeCancel.finished=performance.now();observer.disconnect();}
    });observer.observe(document.body,{childList:true,subtree:true,characterData:true});
  });
  await page.locator('#file-input').setInputFiles(file);await page.waitForFunction(()=>globalThis.__largeCancel?.finished);const cancellation=await page.evaluate(()=>__largeCancel);assert.equal(cancellation.enabled,true);const cancelAt=cancellation.at,canceledAt=cancellation.finished;await settle();
  const gaps=await page.evaluate(()=>{clearInterval(__parseTimer);return {samples:__parseGaps.length,maxGapMs:Math.max(0,...__parseGaps)};});assert.equal(await page.evaluate(()=>JSON.stringify(sheetStudio.getWorkbook())),original);const cleanupMs=await parsersClosed();assert.equal(createHash('sha256').update(await fs.readFile(file)).digest('hex'),sha);assert.ok(canceledAt-cancelAt<1000);
  await page.locator('#file-input').setInputFiles(file);await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='large');await settle();await page.evaluate(()=>sheetStudio.flushRecovery());assert.equal(await page.evaluate(()=>sheetStudio.historyStats().cells),100000);assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('J10000')),20000);await parsersClosed();
  const metrics={inputCells:100000,inputBytes:data.length,inputSHA256:sha,cancelDuring:'decode',cancelMs:canceledAt-cancelAt,cleanupMs,...gaps,afterCancelOriginalUnchanged:true,subsequentImport:true,pageErrors:errors};assert.deepEqual(errors,[]);await fs.writeFile(path.join(root,'test-results/import-cancel-large.json'),JSON.stringify(metrics,null,2));console.log('PASS: cancel during real 100k-cell XLSX decode, no source/current-document changes, parser released and subsequent import succeeds',JSON.stringify(metrics));
}catch(error){await page.screenshot({path:path.join(root,'test-results/import-cancel-large-failure.png')});throw error;}
finally{await app.close();await fs.rm(fixture,{recursive:true,force:true});}
