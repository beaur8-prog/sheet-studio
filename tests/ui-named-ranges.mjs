import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','named-ranges-'+Date.now());
const env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch(desktopLaunch(root,env));
const page=await app.firstWindow(),errors=[];page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(15000);
await page.setViewportSize({width:1280,height:900});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
let sequence=0;
async function load(book){book.name='recalculation-'+ ++sequence;await page.locator('#file-input').setInputFiles({name:book.name+'.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,book.name);await settle();}
async function edit(ref,raw){await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();}
async function save(){const file=path.join(dataDir,'result-'+sequence+'.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');return file;}
const closeDialog=async()=>{await page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();await settle();};
const manager=async()=>{await page.locator('[data-tab="수식"]').click();await page.locator('[data-action="name-manager"]').click();await page.getByRole('dialog',{name:'이름 관리자'}).waitFor();};
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  const book=createWorkbook();book.sheets[0].name='Input';book.sheets[0].cells={A1:{raw:'2'},A2:{raw:'3'},A3:{raw:'4'},B1:{raw:'=SUM(Revenue)',cachedFormula:'=SUM(Revenue)',cachedValue:99}};await load(book);assert.equal(await value('B1'),'#NAME?');
  await go('A1:A3');await manager();await page.getByRole('button',{name:'새 이름',exact:true}).click();await page.locator('#name-label').fill('Revenue');assert.equal(await page.locator('#name-formula').inputValue(),"'Input'!$A$1:$A$3");await page.getByRole('dialog').getByRole('button',{name:'확인',exact:true}).click();await closeDialog();assert.equal(await value('B1'),9);
  await manager();await page.locator('[data-name-edit="0"]').click();await page.locator('#name-formula').fill("'Input'!$A$2:$A$3");await page.getByRole('dialog').getByRole('button',{name:'확인',exact:true}).click();await closeDialog();assert.equal(await value('B1'),7);
  await page.locator('#undo').click();await settle();assert.equal(await value('B1'),9);await page.locator('#redo').click();await settle();assert.equal(await value('B1'),7);
  await go('Revenue');assert.equal(await page.locator('#name-box').inputValue(),'A2:A3');
  const file=await save(),out=new ExcelJS.Workbook();await out.xlsx.load(await fs.readFile(file));assert.equal(out.worksheets[0].getCell('B1').result,7);assert.deepEqual(out.definedNames.getRanges('Revenue').ranges,['Input!$A$2:$A$3']);
  await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('B1'),7);await manager();await page.screenshot({path:path.join(root,'test-results/named-ranges.png')});await page.locator('[data-name-delete="0"]').click();await closeDialog();assert.equal(await value('B1'),'#NAME?');
  const deleted=await save(),check=new ExcelJS.Workbook();await check.xlsx.load(await fs.readFile(deleted));assert.equal(check.worksheets[0].getCell('B1').result,undefined);
  assert.deepEqual(errors,[]);console.log('PASS: create/edit/delete defined name through the manager, calculate dependent formulas, undo/redo, name-box range navigation and XLSX save/reopen without resurrecting deleted-name caches');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/named-ranges-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
