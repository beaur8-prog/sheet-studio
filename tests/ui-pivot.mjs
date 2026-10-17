import {chromium,_electron} from 'playwright';
import path from 'node:path';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
const root=path.resolve(import.meta.dirname,'..'),desktop=process.argv.includes('--packaged');
let application,browser,page;
if(desktop){const env={...process.env,SHEET_STUDIO_DATA:path.join(root,'.package','pivot-smoke-'+Date.now())};delete env.ELECTRON_RUN_AS_NODE;application=await _electron.launch({executablePath:path.join(root,'release/win-unpacked/Sheet Studio.exe'),env});page=await application.firstWindow();}
else{browser=await chromium.launch({headless:true,channel:'msedge'});page=await browser.newPage({viewport:{width:1480,height:960}});}
const errors=[];
page.on('pageerror',e=>errors.push(e.message));
const go=async address=>{await page.locator('#name-box').fill(address);await page.locator('#name-box').press('Enter');};
const value=addr=>page.evaluate(a=>sheetStudio.evaluate(a),addr);
const enter=async(addr,raw)=>{await go(addr);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');};
try{
 if(!desktop)await page.goto('http://127.0.0.1:5192');await page.locator('#cells .cell').first().waitFor();
 await page.locator('#file-input').setInputFiles({name:'매출.csv',mimeType:'text/csv',buffer:Buffer.from('지역,상품,매출,수량\n동부,A,10,1\n동부,A,20,2\n동부,B,30,3\n서부,A,100,4\n서부,B,,5')});
 await page.waitForFunction(()=>sheetStudio.getWorkbook().sheets[0].cells.A1?.raw==='지역');
 await go('A1:D6');await page.getByRole('button',{name:'삽입',exact:true}).click();await page.getByRole('button',{name:'피벗테이블',exact:true}).click();
 const dialog=page.getByRole('dialog');await dialog.getByRole('button',{name:'+ 열 추가',exact:true}).click();
 await dialog.getByLabel('열 필드 1',{exact:true}).selectOption('상품');
 await dialog.getByLabel('값 필드 1',{exact:true}).selectOption('매출');
 await fs.mkdir('test-results',{recursive:true});await page.screenshot({path:'test-results/pivot-fields.png'});
 await dialog.getByRole('button',{name:'만들기',exact:true}).click();await page.locator('#pivot-bar').waitFor();
 assert.equal(await value('D4'),160);assert.equal(await page.locator('#formula').getAttribute('readonly'),'');
 await go('B2');await page.keyboard.press('Delete');assert.equal(await value('B2'),30);
 await page.locator('#pivot-bar').getByRole('button',{name:'원본 보기',exact:true}).click();await enter('C2','40');
 await page.locator('[data-sheet]').filter({hasText:'피벗 1'}).click();assert.equal(await value('D4'),160);
 await page.locator('#pivot-bar').getByRole('button',{name:'새로 고침',exact:true}).click();assert.equal(await value('D4'),190);
 await page.locator('#pivot-bar').getByRole('button',{name:'필드 편집',exact:true}).click();
 await dialog.getByLabel('집계 방식 1',{exact:true}).selectOption('average');await dialog.getByRole('button',{name:'적용',exact:true}).click();assert.equal(await value('D4'),47.5);
 await page.screenshot({path:'test-results/pivot-report.png'});
 await page.reload();await page.locator('#pivot-bar').waitFor();assert.equal(await value('D4'),47.5);
 let saved;
 if(desktop){const file=path.join(root,'.package','pivot-ui-export.xlsx');await application.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');saved=await fs.readFile(file);}
 else{const downloadPromise=page.waitForEvent('download');await page.locator('#save').click();const download=await downloadPromise;saved=await fs.readFile(await download.path());}
 await page.locator('#file-input').setInputFiles({name:'피벗.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:saved});
 await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='피벗');assert.equal(await value('D4'),47.5);
 await page.locator('#pivot-bar').getByRole('button',{name:'필드 편집',exact:true}).click();await dialog.getByRole('button',{name:'+ 필터 추가',exact:true}).click();await dialog.getByLabel('필터 필드 1',{exact:true}).selectOption('지역');
 await dialog.getByLabel('필터 항목 1',{exact:true}).selectOption({label:'동부'});await dialog.getByRole('button',{name:'적용',exact:true}).click();assert.equal(await value('D3'),30);
 await page.locator('#pivot-bar').getByRole('button',{name:'일반 범위로 변환',exact:true}).click();await page.getByRole('dialog').getByRole('button',{name:'변환',exact:true}).click();assert.equal(await page.locator('#pivot-bar').isVisible(),false);await enter('B2','999');assert.equal(await value('B2'),999);
 await page.keyboard.press('Control+z');await page.keyboard.press('Control+z');await page.locator('#pivot-bar').waitFor();assert.equal(await value('B2'),30);
 assert.deepEqual(errors,[]);console.log('Pivot UI: create, edit, preview, protection, refresh, reload, XLSX roundtrip, filter, conversion and undo passed.');
}finally{if(application)await application.close();if(browser)await browser.close();}
