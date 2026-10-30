import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','recalculation-'+Date.now());
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
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  let book=createWorkbook();book.sheets[0].cells={A1:{raw:'=RAND()'},B1:{raw:'=A1*2'}};await load(book);
  const initial=await value('A1');assert.equal(await value('B1'),initial*2);
  await page.locator('[data-tab="수식"]').click();await page.locator('[data-action="recalculate"]').click();await settle();
  const next=await value('A1');assert.notEqual(next,initial);assert.equal(await value('B1'),next*2);
  let file=await save(),saved=new ExcelJS.Workbook();await saved.xlsx.load(await fs.readFile(file));
  assert.equal(saved.worksheets[0].getCell('A1').result,next);assert.equal(saved.worksheets[0].getCell('B1').result,next*2);
  await page.locator('#grid').focus();await page.keyboard.press('F9');await settle();const keyboardValue=await value('A1');assert.notEqual(keyboardValue,next);assert.equal(await value('B1'),keyboardValue*2);

  book=createWorkbook();for(let i=1;i<=10000;i++)book.sheets[0].cells['A'+i]={raw:i===1?'1':`=A${i-1}+1`};await load(book);await go('A10000');
  assert.equal(await value('A10000'),10000);assert.match(await page.getByRole('gridcell',{name:'A10000',exact:true}).innerText(),/10,?000/);
  await edit('A1','2');await go('A10000');assert.equal(await value('A10000'),10001);
  await page.locator('#undo').click();await settle();assert.equal(await value('A10000'),10000);
  await page.locator('#redo').click();await settle();assert.equal(await value('A10000'),10001);
  file=await save();await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('A10000'),10001);
  await page.screenshot({path:path.join(root,'test-results/recalculation-chain.png')});

  book=createWorkbook();book.sheets[0].cells=Object.fromEntries(Object.entries({A1:'=DATE(2011,3,11)',A2:'=DATE(2010,1,1)',A3:'=DATE(2010,12,31)',A4:'=DATE(2010,6,30)',B1:'10',B2:'20',B3:'30',B4:'40',C1:'0',C3:'2'}).map(([ref,raw])=>[ref,{raw}]));await load(book);
  for(const [ref,formula,expected] of [['E1','=COUNTIF(A1:A4,">3/1/2010")',3],['E2','=SUMIF(A1:A4,">2010-03-01",B1:B4)',80],['E3','=AVERAGEIFS(B1:B3,C1:C3,0)',15],['E4','=AVERAGEIF(B1:B3,">0",B1)',20]]){await edit(ref,formula);assert.equal(await value(ref),expected,formula);}
  file=await save();await page.locator('#file-input').setInputFiles(file);await settle();
  for(const [ref,expected] of [['E1',3],['E2',80],['E3',15],['E4',20]])assert.equal(await value(ref),expected);
  const source=new ExcelJS.Workbook();source.properties.date1904=true;const dateSheet=source.addWorksheet('Dates');
  dateSheet.getCell('A1').value='날짜';dateSheet.getCell('B1').value='날짜 차이';
  dateSheet.getCell('A2').value=new Date('2011-07-05T00:00:00Z');dateSheet.getCell('A2').numFmt='yyyy-mm-dd';
  dateSheet.getCell('A3').value=new Date('2011-07-06T00:00:00Z');dateSheet.getCell('A3').numFmt='yyyy-mm-dd';
  dateSheet.getCell('B2').value={formula:'DATE(2011,7,5)-39267',result:0};dateSheet.getCell('B3').value=2;
  await page.locator('#file-input').setInputFiles({name:'epoch-1904.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(await source.xlsx.writeBuffer())});
  await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='epoch-1904');await settle();await go('A2');
  assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().date1904),true);
  assert.equal(await page.getByRole('gridcell',{name:'A2',exact:true}).innerText(),'2011-07-05');assert.equal(await value('B2'),0);
  await edit('B2','=DATE(2011,7,6)-A2');assert.equal(await value('B2'),1);
  await page.locator('[data-tab="데이터"]').click();await go('A1:B3');await page.locator('[data-action="filter"]').click();await settle();
  await page.locator('[data-filter-column="0"]').click();await page.getByRole('checkbox',{name:'2011-07-05',exact:true}).waitFor({state:'visible'});
  await page.locator('.filter-condition-menu summary').click();await page.locator('[data-condition="greaterThan"]').click();await page.locator('#filter-value-0').fill('2011-07-05');await page.locator('#filter-ok').click();await settle();
  assert.deepEqual(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].autoFilter.hiddenRows),[1]);
  file=await save();saved=new ExcelJS.Workbook();await saved.xlsx.load(await fs.readFile(file));assert.equal(saved.properties.date1904,true);assert.equal(saved.worksheets[0].getCell('B2').result,1);
  await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('B2'),1);
  assert.deepEqual(errors,[]);console.log('PASS: Calculate Now and F9 refresh RAND and dependents; displayed values match XLSX caches; 10,000-cell chain navigation/edit/undo/redo/save/reopen; date, blank and AVERAGEIF criteria entered through the UI; 1904 XLSX date display, formulas, date filter and roundtrip');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/recalculation-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
