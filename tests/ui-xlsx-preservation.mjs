import {desktopLaunch} from './electron-launch.mjs';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {preservationFixture} from './xlsx-preservation-fixture.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','xlsx-preservation-'+Date.now());
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
async function edit(ref,raw){await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  const source=await JSZip.loadAsync(await preservationFixture());let xmlText=await source.file('xl/worksheets/sheet1.xml').async('string');xmlText=xmlText.replace(/<sheetProtection[^>]*\/>/,'');source.file('xl/worksheets/sheet1.xml',xmlText);
  await page.locator('#file-input').setInputFiles({name:'preservation-ui.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(await source.generateAsync({type:'uint8array'}))});
  await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='preservation-ui');await settle();
  assert.equal(await value('A1'),'중요 메모');await go('A1');await page.locator('[data-tab="홈"]').click();await page.locator('[data-action="bold"]').click();await settle();
  await edit('A2','도움말');const file=await save(),out=new ExcelJS.Workbook();await out.xlsx.load(await fs.readFile(file));const cells=out.worksheets[0];
  assert.equal(cells.getCell('A1').alignment.textRotation,45);assert.equal(cells.getCell('A1').font.strike,true);assert.ok(cells.getCell('A1').value.richText.every(run=>run.font.bold));
  assert.equal(cells.getCell('A2').value.hyperlink,'https://example.org/manual?q=a&x=b');assert.equal(cells.getCell('A2').value.text,'도움말');assert.equal(cells.getCell('A3').fill.type,'gradient');assert.equal(cells.getCell('B1').value.shareType,'array');assert.equal(cells.getCell('D1').dataValidation.type,'list');assert.equal(out.definedNames.getRanges('Revenue').ranges[0],'Details!$A$3');
  await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('A2'),'도움말');assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.A2.xlsxValue.value.tooltip),'원본 도움말');
  await page.locator('#undo').click();await settle();await page.locator('#redo').click();await settle();
  await page.screenshot({path:path.join(root,'test-results/xlsx-preservation.png')});
  assert.deepEqual(errors,[]);console.log('PASS: original XLSX advanced formatting, rich text, link editing, array formulas, validation and defined names survive actual Electron edit/save/reopen');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/xlsx-preservation-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
