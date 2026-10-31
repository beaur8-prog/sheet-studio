import {assertBlankStart,recoverPrevious} from './recovery-ui.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const tall=process.argv.includes('--tall'),columns=tall?1:100,rows=1_000_000/columns,last=tall?'A1000000':'CV10000',sumFormula=tall?'=SUM(A2:A999999)':'=SUM(A2:CU10000)',initialSum=tall?999998:989901,over=tall?'B1':'CW1',resultName=tall?'capacity-tall':'capacity-million';
const dataDir=path.join(root,'.package','capacity-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const executablePath=process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix()));
let app,page;const errors=[],findings=[];
const record=finding=>{findings.push(finding);console.log(JSON.stringify(finding));};
async function launch(){app=await _electron.launch({executablePath,args:packaged?[]:[root],env});page=await app.firstWindow();page.setDefaultTimeout(240000);await page.setViewportSize({width:1280,height:900});page.on('pageerror',e=>errors.push(e.message));await page.locator('.cell').first().waitFor();}
const select=async range=>{await page.locator('#name-box').fill(range);await page.locator('#name-box').press('Enter');};
const raw=ref=>page.evaluate(ref=>sheetStudio.getCell(ref)?.raw,ref);
async function heartbeat(){await page.evaluate(()=>{window.capacityGaps=[];let last=performance.now();window.capacityTimer=setInterval(()=>{const now=performance.now();capacityGaps.push(now-last);last=now;},20);});}
async function gaps(){return page.evaluate(()=>{clearInterval(capacityTimer);return {samples:capacityGaps.length,maxGapMs:Math.round(Math.max(0,...capacityGaps))};});}
try{
 await fs.mkdir('test-results',{recursive:true});await launch();
 const header=Array.from({length:columns},(_,i)=>'Field '+(i+1)).join(',');
 const csv=header+'\n'+(Array(columns).fill('1').join(',')+'\n').repeat(rows-1);
 await heartbeat();const started=Date.now();
 await page.locator('#file-input').setInputFiles({name:'million.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
 await page.waitForFunction(()=>sheetStudio.historyStats().cells===1_000_000);
 await page.evaluate(()=>sheetStudio.flushRecovery());
 record({case:'one million CSV cells imported and durably saved',ms:Date.now()-started,...await gaps()});
 await select('A1');await heartbeat();const navigating=Date.now();await page.keyboard.press('ControlOrMeta+ArrowDown');await page.waitForFunction(row=>sheetStudio.getSelection().r===row,rows-1);await page.evaluate(()=>sheetStudio.whenIdle());record({case:'data-edge navigation to the final populated row',ms:Date.now()-navigating,...await gaps()});await page.keyboard.press('ControlOrMeta+ArrowUp');await page.waitForFunction(()=>sheetStudio.getSelection().r===0);
 if(process.argv.includes('--structure')){
  await select('A2:A101');await page.locator('[data-tab="홈"]').click();await heartbeat();const changing=Date.now();await page.locator('[data-action="insert-row"]').click();await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());assert.equal(await raw('A102'),'1');assert.equal(await raw('A2'),undefined);assert.equal(await page.evaluate(()=>sheetStudio.historyStats().cells),1_000_000);
  const stats=await page.evaluate(()=>sheetStudio.historyStats());record({case:'100 rows inserted in a million-cell document',ms:Date.now()-changing,undoRetained:stats.undo>0,...await gaps()});
  if(stats.undo>0){await page.locator('#undo').click();await page.evaluate(()=>sheetStudio.whenIdle());assert.equal(await raw('A2'),'1');await page.locator('#redo').click();await page.evaluate(()=>sheetStudio.whenIdle());assert.equal(await raw('A102'),'1');}
  await select('A2:A101');await page.locator('[data-action="delete-menu"]').click();await page.locator('[data-action="delete-row"]').click();await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());assert.equal(await raw('A2'),'1');assert.equal(await raw(last),'1');
 }
 await select(last);await page.keyboard.type(sumFormula);await page.keyboard.press('Enter');
 await select(last);await page.waitForFunction(({last,value})=>sheetStudio.evaluate(last)===value,{last,value:initialSum});
 record({case:'background sum over a large range',value:initialSum});
 await select(over);await page.keyboard.type('over limit');await page.keyboard.press('Enter');
 await page.waitForFunction(()=>document.querySelector('.toast')?.textContent.includes('변경을 취소'));
 assert.equal(await page.evaluate(()=>sheetStudio.historyStats().cells),1_000_000);
 assert.equal(await raw(over),undefined);
 await select('A2');await heartbeat();
 for(let i=0;i<65;i++){await page.keyboard.type(String(i));await page.keyboard.press('Enter');await select('A2');}
 const stats=await page.evaluate(()=>sheetStudio.historyStats());assert.equal(stats.undo,60);assert.ok(stats.bytes<100_000,JSON.stringify(stats));
 record({case:'65 single-cell edits keep only small undo deltas',...stats,...await gaps()});
 await page.keyboard.press('ControlOrMeta+z');await page.keyboard.press('ControlOrMeta+Shift+z');
 const delta=await page.evaluate(()=>sheetStudio.flushRecovery());assert.ok(delta.chunksWritten<=2,JSON.stringify(delta));
 record({case:'incremental recovery writes at most two affected chunks',...delta});
 await select('A3');await page.keyboard.type('128');await page.keyboard.press('Enter');
 await select('A2:'+(tall?'A999999':last));await page.getByRole('button',{name:'데이터',exact:true}).click();await page.locator('[data-action="sort-desc"]').click();
 await page.locator('.workbook-busy').waitFor({state:'detached'});assert.equal(await raw('A2'),'128');assert.equal(await raw('A3'),'64');
 record({case:'background sort moves rows in a million-cell document',firstValue:'128',secondValue:'64'});
 await select('A1:'+last);await page.getByRole('button',{name:'데이터',exact:true}).click();await page.getByRole('button',{name:'필터',exact:true}).click();
 await page.locator('.workbook-busy').waitFor({state:'detached'});await page.locator('[data-filter-column="0"]').click();
 await page.locator('.filter-condition-menu summary').click();await page.locator('[data-condition="greaterThan"]').click();await page.locator('#filter-value-0').fill('10');await page.locator('#filter-ok').click();
 await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());
 assert.equal(await page.getByRole('gridcell',{name:'A2',exact:true}).textContent(),'128');assert.equal(await page.getByRole('gridcell',{name:'A4',exact:true}).count(),0);
 record({case:'background filter covers the full million-cell table',firstVisibleValue:'128'});
 await app.close();await launch();await assertBlankStart(page);await recoverPrevious(page);
 assert.equal(await page.evaluate(()=>sheetStudio.historyStats().cells),1_000_000);
 assert.equal(await raw('A2'),'128');assert.equal(await raw(last),sumFormula);
 assert.equal(await page.locator('[data-filter-column="0"].is-filtered').count(),1);
 await select(last);await page.waitForFunction(({last,value})=>sheetStudio.evaluate(last)===value,{last,value:initialSum+190});
 record({case:'explicit recovery after a blank restart restores all cells, edits and formula',value:initialSum+190});
 const output=path.join(dataDir,'million.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);
 await heartbeat();const saving=Date.now();await page.locator('#save').click();
 await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');
 const bytes=(await fs.stat(output)).size;
 record({case:'one million cells exported as XLSX in the worker',ms:Date.now()-saving,bytes,...await gaps()});
 await heartbeat();const reading=Date.now();await page.locator('#file-input').setInputFiles(output);
 await page.waitForFunction(()=>document.querySelector('.toast')?.textContent.startsWith('파일을 열었습니다.'));
 await page.waitForFunction(()=>sheetStudio.historyStats().cells===1_000_000);
 await page.evaluate(()=>sheetStudio.flushRecovery());
 assert.equal(await raw('A2'),'128');assert.equal(await raw(last),sumFormula);
 record({case:'actual exported XLSX reopened without losing the first or last cell',ms:Date.now()-reading,...await gaps()});
 await select('A1');await page.screenshot({path:'test-results/'+resultName+'.png'});
 assert.deepEqual(errors,[]);await fs.writeFile('test-results/'+resultName+'.json',JSON.stringify(findings,null,2));console.log(JSON.stringify(findings,null,2));
}catch(error){if(page)await page.screenshot({path:'test-results/capacity-failure.png'}).catch(()=>{});await fs.writeFile('test-results/capacity-partial.json',JSON.stringify({findings,error:String(error),errors},null,2));throw error;}finally{if(app)await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
