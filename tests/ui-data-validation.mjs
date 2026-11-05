import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import {createWorkbook} from '../src/engine.mjs';
import {validationRules,validateCellUpdates} from '../src/cell-constraints.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','validation-authoring-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:dataDir}));
const page=await app.firstWindow(),errors=[];page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(15000);
await page.setViewportSize({width:1280,height:900});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
const snapshot=()=>page.evaluate(()=>sheetStudio.getWorkbook());
const open=async range=>{await go(range);await page.locator('[data-tab="데이터"]').click();await page.locator('[data-action="validation"]').click();await page.getByRole('dialog',{name:'데이터 유효성 검사'}).waitFor();};
const close=()=>page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();
async function edit(ref,raw){await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();}
async function apply(){await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await settle();}
async function undo(reverse=false){await page.locator('#grid').focus();await page.keyboard.press((process.platform==='darwin'?'Meta':'Control')+(reverse?'+Shift':'')+'+z');await settle();}
let name='';async function load(book){name='validation-authoring-'+Date.now();book.name=name;await page.locator('#file-input').setInputFiles({name:name+'.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,name);await settle();}
async function paste(text){await page.locator('#grid').focus();await page.evaluate(text=>{const data=new DataTransfer();data.setData('text/plain',text);document.querySelector('#grid').dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));},text);await settle();}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();await settle();
  if(process.argv.includes('--baseline')){
    await page.locator('[data-tab="데이터"]').click();assert.equal(await page.locator('[data-action="validation"]').count(),0);await page.screenshot({path:path.join(root,'test-results/data-validation-before.png')});console.log('BASELINE CONFIRMED: 0.7.38 has no data-validation authoring command');
  }else{
    const book=createWorkbook();book.sheets[0].cells={A1:{raw:'기존 값'},H1:{raw:'5'}};await load(book);
    await open('A2:A8');await page.locator('#dv-type').selectOption('list');await page.locator('#dv-source').fill('예,아니오');await page.locator('summary').filter({hasText:'입력 안내'}).click();await page.locator('#dv-show-input').check();await page.locator('#dv-prompt-title').fill('승인 여부');await page.locator('#dv-prompt').fill('예 또는 아니오를 선택하세요.');await page.locator('#dv-title').fill('조건 확인');await page.locator('#dv-message').fill('목록의 항목을 입력하세요.');await apply();
    await go('A2');await page.locator('#validation-prompt').waitFor({state:'visible'});assert.match(await page.locator('#validation-prompt').innerText(),/승인 여부/);await page.keyboard.press('Escape');await page.locator('#validation-prompt').waitFor({state:'hidden'});await go('B1');await go('A2');await page.locator('#validation-prompt').waitFor({state:'visible'});
    await page.locator('#validation-list-button').click();await page.getByRole('option',{name:'아니오',exact:true}).click();await settle();assert.equal(await value('A2'),'아니오');
    await edit('A3','보류');assert.equal(await value('A3'),'');assert.match(await page.locator('.toast').innerText(),/목록의 항목/);await page.locator('#editor').fill('예');await page.locator('#editor').press('Enter');await settle();assert.equal(await value('A3'),'예');
    await open('B2:B8');await page.locator('#dv-type').selectOption('whole');await page.locator('#dv-value').fill('1');await page.locator('#dv-value2').fill('10');await apply();await edit('B2','4');await go('B3:B4');await paste('5\n99');assert.equal(await value('B3'),5);assert.equal(await value('B4'),99);await open('B2:B8');await page.getByRole('dialog').getByRole('button',{name:'현재 값 검사'}).click();await settle();assert.match(await page.locator('#dv-review').innerText(),/조건 불일치 1개/);assert.match(await page.locator('#dv-review').innerText(),/B4/);await close();await undo();assert.equal(await value('B3'),'');assert.equal(await value('B4'),'');
    await open('B4:B5');assert.equal(await page.locator('#dv-value').inputValue(),'1');await page.locator('#dv-value').fill('20');await apply();assert.match(await page.locator('#dv-error').innerText(),/최솟값/);await page.locator('#dv-value').fill('1');await page.locator('#dv-value2').fill('3');await close();assert.equal(validationRules((await snapshot()).sheets[0]).filter(r=>r.type==='whole').length,1);
    await open('B4:B5');await page.getByRole('dialog').getByRole('button',{name:'규칙 지우기',exact:true}).click();await settle();await edit('B4','99');assert.equal(await value('B4'),99);await edit('B6','99');assert.equal(await value('B6'),'');await page.keyboard.press('Escape');
    await undo();assert.equal(await value('B4'),'');await undo();let current=await snapshot();assert.equal(validateCellUpdates(current,current.sheets[0],[{key:'B4',cell:{raw:'99'}}]).severity,'stop');await undo(true);current=await snapshot();assert.equal(validateCellUpdates(current,current.sheets[0],[{key:'B4',cell:{raw:'99'}}]),null);
    await open('C2:C8');await page.locator('#dv-type').selectOption('custom');await page.locator('#dv-value').fill('=C2>$H$1');await page.locator('#dv-severity').selectOption('warning');await page.locator('#dv-title').fill('금액 확인');await apply();
    await edit('C3','2');await page.getByRole('dialog',{name:'금액 확인'}).waitFor();assert.equal(await value('C3'),'');await close();assert.equal(await page.locator('#editor').inputValue(),'2');await page.locator('#editor').press('Enter');await page.getByRole('dialog').getByRole('button',{name:'계속 입력'}).click();await settle();assert.equal(await value('C3'),2);
    await open('C5:C6');assert.equal(await page.locator('#dv-value').inputValue(),'C5>$H$1');await close();
    await open('D2:D8');await page.locator('#dv-type').selectOption('date');await page.locator('#dv-value').fill('2026-01-01');await page.locator('#dv-value2').fill('2026-12-31');await apply();
    await open('E2:E8');await page.locator('#dv-type').selectOption('time');await page.locator('#dv-value').fill('09:00');await page.locator('#dv-value2').fill('18:00');await apply();
    await open('F2:F8');await page.locator('#dv-type').selectOption('textLength');await page.locator('#dv-operator').selectOption('lessThanOrEqual');await page.locator('#dv-value').fill('3');await apply();
    await open('G2:G8');await page.locator('#dv-type').selectOption('decimal');await page.locator('#dv-operator').selectOption('greaterThan');await page.locator('#dv-value').fill('0');await apply();
    await go('A1');await page.locator('[data-tab="홈"]').click();await page.locator('[data-action="insert-row"]').click();await settle();assert.equal((await snapshot()).sheets[0].cells.A2.raw,'기존 값');
    const target=path.join(dataDir,'validation-authored.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},target);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');
    const independent=new ExcelJS.Workbook();await independent.xlsx.readFile(target);assert.equal(independent.worksheets[0].getCell('A3').dataValidation.promptTitle,'승인 여부');assert.equal(independent.worksheets[0].getCell('C4').dataValidation.errorStyle,'warning');assert.equal(independent.worksheets[0].getCell('D3').dataValidation.type,'date');
    await page.locator('#file-input').setInputFiles(target);await settle();await go('A3');await page.locator('#validation-prompt').waitFor({state:'visible'});assert.equal(await value('A3'),'아니오');await open('C6:C7');assert.equal(await page.locator('#dv-value').inputValue(),'C6>$H$2');await close();
    await page.setViewportSize({width:1000,height:680});await open('A3:A9');await page.screenshot({path:path.join(root,'test-results/data-validation.png')});assert.ok(await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).isVisible());await close();
    const protectedBook=createWorkbook();protectedBook.sheets[0].xlsxFeatures={sheetProtection:'<sheetProtection xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" sheet="1"/>'};await load(protectedBook);await go('A1');await page.locator('[data-tab="데이터"]').click();await page.locator('[data-action="validation"]').click();assert.equal(await page.getByRole('dialog').count(),0);assert.match(await page.locator('.toast').innerText(),/보호/);
    assert.deepEqual(errors,[]);await fs.writeFile(path.join(root,'test-results/data-validation.json'),JSON.stringify({version:JSON.parse(await fs.readFile(path.join(root,'package.json'))).version,scenarios:['seven rule types','dropdown','input prompt and Escape','stop retry','warning cancel/accept','Excel-style paste bypass with undo','partial clear','undo/redo','relative prefill','row insert','XLSX reopen','1000x680 layout','protected sheet'],pageErrors:errors},null,2));
    console.log('PASS: data-validation authoring, input guidance, restrictions, paste behavior, partial clear, undo/redo, relative rules, row insertion and XLSX reopen');
  }
}catch(error){await page.screenshot({path:path.join(root,'test-results/data-validation-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
