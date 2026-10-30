import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','table-editing-'+Date.now());
const env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix()),args:[root],env});
const page=await app.firstWindow(),errors=[];page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(15000);
await page.setViewportSize({width:1280,height:900});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
let sequence=0;
async function load(book){book.name='recalculation-'+ ++sequence;await page.locator('#file-input').setInputFiles({name:book.name+'.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,book.name);await settle();}
async function edit(ref,raw){await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();}
async function save(){const file=path.join(dataDir,'result-'+sequence+'.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');return file;}
const act=async action=>{await page.locator(`[data-action="${action}"]`).first().click();await settle();};
async function applyDialog(name='적용'){await page.getByRole('dialog').getByRole('button',{name,exact:true}).click();await page.getByRole('dialog').waitFor({state:'detached'});await settle();}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  const book=createWorkbook();book.sheets[0].name='Input';book.sheets[0].cells={A1:{raw:'Qty'},B1:{raw:'Price'},C1:{raw:'Amount'},A2:{raw:'2'},B2:{raw:'3'},A3:{raw:'4'},B3:{raw:'5'}};await load(book);const originalSize=await page.getByRole('gridcell',{name:'A1',exact:true}).evaluate(el=>({width:el.offsetWidth,height:el.offsetHeight}));await go('A1:C3');await page.locator('[data-tab="삽입"]').click();await act('table-style');await page.locator('#table-name').fill('Sales');await applyDialog('표 만들기');
  assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].tables[0].name),'Sales');assert.equal(await page.getByRole('gridcell',{name:'A1',exact:true}).evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(16, 124, 65)');assert.equal(await page.getByRole('gridcell',{name:'A1',exact:true}).evaluate(el=>getComputedStyle(el).color),'rgb(255, 255, 255)');
  await go('C2');await act('table-calc');await page.locator('#table-calc-formula').fill('=[@Qty]*[@Price]');await applyDialog('전체 행에 적용');assert.equal(await value('C2'),6);assert.equal(await value('C3'),20);
  await edit('E1','=SUM(Sales[Amount])');assert.equal(await value('E1'),26);await go('B2');await act('table-style');await page.locator('#table-name').fill('Revenue');await applyDialog();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.E1.raw),'=SUM(Revenue[Amount])');assert.equal(await value('E1'),26);
  await edit('A1','Count');assert.match(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.C2.raw),/Count/);assert.equal(await value('C2'),6);await page.locator('#undo').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].tables[0].columns[0].name),'Qty');await page.locator('#redo').click();await settle();assert.equal(await value('C3'),20);
  await go('B2');await act('table-style');await page.locator('#table-totals').check();assert.equal(await page.locator('#table-range').inputValue(),'A1:C4');await applyDialog();assert.equal(await value('A4'),'합계');assert.equal(await value('C4'),26);
  await go('A1');await page.locator('[data-filter-column="0"]').click();await page.locator('#filter-all').waitFor();await page.locator('#filter-all').uncheck();await page.getByRole('checkbox',{name:'2',exact:true}).check();await page.locator('#filter-ok').click();await settle();assert.equal(await value('C4'),6);assert.equal(await value('E1'),26);
  const filteredFile=await save();await page.locator('#file-input').setInputFiles(filteredFile);await settle();assert.equal(await value('C4'),6);await go('A1');await page.locator('[data-tab="데이터"]').click();await act('clear-filter');assert.equal(await value('C4'),26);
  await page.locator('[data-filter-column="0"]').click();await page.locator('.filter-color-menu summary').click();await page.locator('[data-color-target="background"][value="#e2efda"]').check();await page.locator('#filter-ok').click();await settle();assert.equal(await value('C4'),6);await act('clear-filter');assert.equal(await value('C4'),26);
  let file=await save(),out=new ExcelJS.Workbook();await out.xlsx.load(await fs.readFile(file));assert.equal(out.worksheets[0].getTable('Revenue').table.totalsRow,true);assert.equal(out.worksheets[0].getCell('C4').result,26);await page.locator('#file-input').setInputFiles(file);await settle();await go('A1');assert.deepEqual(await page.getByRole('gridcell',{name:'A1',exact:true}).evaluate(el=>({width:el.offsetWidth,height:el.offsetHeight})),originalSize);assert.equal(await page.getByRole('gridcell',{name:'A1',exact:true}).evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(16, 124, 65)');
  await go('B2');await act('table-style');await page.locator('#table-totals').uncheck();await applyDialog();await go('B2');await act('table-style');await page.locator('#table-range').fill('A1:C2');await applyDialog();assert.equal(await value('C3'),20);assert.equal(await value('E1'),6);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.C3.raw),'=A3*B3');
  await page.screenshot({path:path.join(root,'test-results/table-editing.png')});assert.deepEqual(errors,[]);console.log('PASS: create styled Excel table, calculated column, rename references, edit header/undo/redo, totals on/off, value/color filters, XLSX style/totals/default-size roundtrip and shrink without breaking excluded formulas');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/table-editing-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
