import JSZip from 'jszip';
import {tableFixture} from './tables-fixture.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {applyFilters} from '../src/filters.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','excel-tables-'+Date.now());
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
async function loadBytes(bytes,name){await page.locator('#file-input').setInputFiles({name:name+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(bytes)});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,name);await settle();}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  await loadBytes(await tableFixture(),'tables-ui');assert.equal(await value('C2'),6);assert.equal(await value('C3'),20);assert.equal(await value('C4'),26);assert.equal(await value('E1'),26);
  await edit('A3','6');assert.equal(await value('C3'),30);assert.equal(await value('C4'),36);assert.equal(await value('E1'),36);await page.locator('#undo').click();await settle();assert.equal(await value('E1'),26);await page.locator('#redo').click();await settle();assert.equal(await value('E1'),36);
  let file=await save(),out=new ExcelJS.Workbook();await out.xlsx.load(await fs.readFile(file));assert.equal(out.worksheets[0].getTable('Sales').table.tableRef,'A1:C4');assert.equal(out.worksheets[0].getCell('C4').result,36);
  await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('C4'),36);await go('B2');await page.locator('[data-tab="데이터"]').click();await page.locator('[data-action="sort-desc"]').click();await settle();assert.equal(await value('A2'),6);assert.equal(await value('A3'),2);assert.equal(await value('A4'),'합계');assert.equal(await value('C4'),36);
  await go('B2');await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="chart-insert"]').click();assert.equal(await page.locator('#chart-range').inputValue(),'A1:C3');await page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();
  const filtered=await readSpreadsheet(await tableFixture(),'filtered.xlsx'),sheet=filtered.sheets[0],table=sheet.tables[0];table.filter.columns[0]={type:'custom',kind:'number',join:'and',conditions:[{operator:'greaterThan',value:'2'}]};applyFilters(filtered,{...sheet,autoFilter:table.filter});await loadBytes((await writeSpreadsheet(filtered)).data,'filtered-tables-ui');assert.equal(await value('C4'),20);assert.equal(await value('E1'),26);assert.equal(await page.getByRole('gridcell',{name:'A2',exact:true}).count(),0);assert.equal(await page.getByRole('gridcell',{name:'A3',exact:true}).count(),1);
  file=await save();const zip=await JSZip.loadAsync(await fs.readFile(file));assert.match(await zip.file('xl/tables/table1.xml').async('string'),/<customFilter[^>]*operator="greaterThan"[^>]*val="2"/);await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('C4'),20);
  await page.screenshot({path:path.join(root,'test-results/excel-tables.png')});assert.deepEqual(errors,[]);console.log('PASS: real Excel table import, current-row formulas/totals, input/edit/undo/redo/save/reopen, sorting without moving totals, chart selection without totals and filtered-table import/visibility/SUBTOTAL roundtrip');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/excel-tables-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
