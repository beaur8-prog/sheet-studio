import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','data-region-'+Date.now());
const env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch(desktopLaunch(root,env));
const page=await app.firstWindow(),errors=[];page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(15000);
await page.setViewportSize({width:1280,height:900});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
let sequence=0;
async function load(book){book.name='recalculation-'+ ++sequence;await page.locator('#file-input').setInputFiles({name:book.name+'.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,book.name);await settle();}

async function save(){const file=path.join(dataDir,'result-'+sequence+'.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');return file;}
const click=async action=>{await page.locator(`[data-action="${action}"]`).click();await settle();};
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  const book=createWorkbook();book.sheets[0].cells=Object.fromEntries(Object.entries({A1:'이름',B1:'매출',A2:'서울',B2:'20',A3:'부산',B3:'10',D10:'상품',E10:'수량',D11:'사과',E11:'2',D12:'배',E12:'1'}).map(([ref,raw])=>[ref,{raw}]));await load(book);
  await page.locator('[data-tab="데이터"]').click();await go('E11');await click('sort-asc');
  assert.equal(await value('D11'),'배');assert.equal(await value('A2'),'서울');assert.equal(await value('D2'),'');
  await page.locator('#undo').click();await settle();assert.equal(await value('D11'),'사과');
  await page.locator('#redo').click();await settle();assert.equal(await value('D11'),'배');
  await go('B2:B3');await click('sort-asc');await page.getByRole('dialog',{name:'정렬 범위 확인'}).waitFor();
  assert.equal(await page.locator('#sort-expand').isChecked(),true);assert.equal(await page.locator('#sort-headers').isChecked(),false);assert.equal(await page.locator('#sort-range').innerText(),'A2:B3');
  await page.getByRole('dialog').getByRole('button',{name:'정렬',exact:true}).click();await settle();assert.equal(await value('A2'),'부산');assert.equal(await value('B2'),10);assert.equal(await value('D11'),'배');
  const file=await save();await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('A2'),'부산');assert.equal(await value('D11'),'배');
  await go('E11');await page.locator('[data-tab="삽입"]').click();await click('chart-insert');assert.equal(await page.locator('#chart-range').inputValue(),'D10:E12');
  await page.getByRole('dialog').getByRole('button',{name:'삽입',exact:true}).click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].charts[0].range),'D10:E12');
  await page.screenshot({path:path.join(root,'test-results/data-region-chart.png')});
  const numbers=createWorkbook();numbers.sheets[0].cells={A1:{raw:'3'},A2:{raw:'2'},A3:{raw:'1'}};await load(numbers);await go('A2');await page.locator('[data-tab="데이터"]').click();await click('sort-asc');
  assert.deepEqual(await Promise.all(['A1','A2','A3'].map(value)),[1,2,3]);
  await go('A2');await page.locator('[data-tab="삽입"]').click();await click('chart-bar');assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].charts[0].range),'A1:A3');
  assert.deepEqual(errors,[]);console.log('PASS: separated-table sort, undo/redo, partial-column expansion, XLSX save/reopen, numeric first row, chart current-region inference through both insertion paths');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/data-region-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
