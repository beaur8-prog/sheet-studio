import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','cell-constraints-'+Date.now());
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
const act=async action=>{await page.locator(`[data-action="${action}"]`).click();await settle();};
async function loadXlsx(source,name){await page.locator('#file-input').setInputFiles({name:name+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(await source.xlsx.writeBuffer())});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,name);await settle();}
async function typeDraft(ref,raw){await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();}
async function pasteText(text){await page.locator('#grid').focus();await page.evaluate(text=>{const data=new DataTransfer();data.setData('text/plain',text);document.querySelector('#grid').dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));},text);await settle();}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  const source=new ExcelJS.Workbook(),input=source.addWorksheet('Input');
  const base={showErrorMessage:true,errorStyle:'stop',errorTitle:'조건 확인',error:'허용된 값을 입력하세요.'};
  input.getCell('A1').value='예';input.getCell('A1').dataValidation={...base,type:'list',formulae:['"예,아니오"']};
  input.getCell('B2').value=5;input.getCell('B2').dataValidation={...base,type:'whole',operator:'between',formulae:[1,10]};
  input.getCell('C2').value=6;input.getCell('C2').dataValidation={...base,type:'custom',formulae:['C2>B2']};
  input.getCell('E2').value='abc';input.getCell('E2').dataValidation={...base,type:'textLength',operator:'lessThanOrEqual',formulae:[3],errorStyle:'warning'};
  input.getCell('G1').value={formula:'SUM(B2:C2)',result:11,shareType:'array',ref:'G1:G3'};input.getCell('G2').value=11;input.getCell('G3').value=11;
  await loadXlsx(source,'constraints-ui');await go('A1');await page.locator('#validation-list-button').click();await page.getByRole('option',{name:'아니오',exact:true}).click();await settle();assert.equal(await value('A1'),'아니오');
  await typeDraft('B2','11');assert.equal(await value('B2'),5);assert.match(await page.locator('.toast').innerText(),/허용된 값/);assert.equal(await page.locator('#editor').inputValue(),'11');
  await page.locator('#editor').fill('8');await page.locator('#editor').press('Enter');await settle();assert.equal(await value('B2'),8);
  await typeDraft('C2','7');assert.equal(await value('C2'),6);await page.locator('#editor').fill('9');await page.locator('#editor').press('Enter');await settle();assert.equal(await value('C2'),9);
  await typeDraft('E2','abcd');await page.getByRole('dialog',{name:'조건 확인'}).waitFor();assert.equal(await value('E2'),'abc');await page.getByRole('dialog').getByRole('button',{name:'계속 입력'}).click();await settle();assert.equal(await value('E2'),'abcd');
  await go('G2');await page.keyboard.press('F2');assert.equal(await page.locator('#editor').evaluate(el=>el.classList.contains('entry-ready')),true);await pasteText('999');assert.equal(await value('G2'),17);assert.match(await page.locator('.toast').innerText(),/배열/);
  await page.locator('[data-tab="홈"]').click();await act('insert-row');assert.equal(await value('B2'),8);assert.match(await page.locator('.toast').innerText(),/배열/);
  await go('A1');await act('insert-row');assert.equal(await value('B3'),8);await typeDraft('B3','99');assert.equal(await value('B3'),8);await page.keyboard.press('Escape');await settle();
  await page.screenshot({path:path.join(root,'test-results/cell-constraints.png')});
  const protectedSource=new ExcelJS.Workbook(),s=protectedSource.addWorksheet('Protected');s.getCell('A1').value='locked';s.getCell('B1').value=1;s.getCell('B1').protection={locked:false};s.getCell('C1').value={formula:'B1*2',result:2};s.getCell('C1').protection={hidden:true};s.getColumn(4).protection={locked:false};await s.protect('secret',{spinCount:1000});const hash=s.sheetProtection.hashValue;
  await loadXlsx(protectedSource,'protected-ui');await go('A1');await page.keyboard.press('F2');assert.equal(await page.locator('#editor').evaluate(el=>el.classList.contains('entry-ready')),true);await page.keyboard.press('Delete');await settle();assert.equal(await value('A1'),'locked');
  await go('A1:B1');await pasteText('bad\t5');assert.equal(await value('A1'),'locked');assert.equal(await value('B1'),1);
  await typeDraft('B1','4');assert.equal(await value('B1'),4);await go('C1');assert.equal(await page.locator('#formula').inputValue(),'');assert.equal(await value('C1'),8);
  await go('D5');await typeDraft('D5','새 입력');assert.equal(await value('D5'),'새 입력');
  await go('A1');await page.locator('[data-tab="홈"]').click();await act('bold');assert.match(await page.locator('.toast').innerText(),/보호/);await act('insert-row');assert.match(await page.locator('.toast').innerText(),/보호/);
  const file=await save(),out=new ExcelJS.Workbook();await out.xlsx.load(await fs.readFile(file));assert.equal(out.worksheets[0].sheetProtection.hashValue,hash);assert.equal(out.worksheets[0].getCell('B1').value,4);assert.equal(out.worksheets[0].getCell('C1').result,8);assert.equal(out.worksheets[0].getCell('D5').protection.locked,false);
  await page.locator('#file-input').setInputFiles(file);await settle();await go('C1');assert.equal(await page.locator('#formula').inputValue(),'');assert.equal(await value('D5'),'새 입력');
  assert.deepEqual(errors,[]);console.log('PASS: validation dropdown, rejected input/correction, dependent custom rule, warning override, array edit/paste/insertion guards, shifted validation, atomic protected paste, unlocked input/column, hidden formulas and password-preserving XLSX roundtrip');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/cell-constraints-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
