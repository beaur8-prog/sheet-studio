import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {desktopLaunch} from './electron-launch.mjs';
import {fixture,spec} from './fixtures/pivot.mjs';
import {applyPivot} from '../src/pivot.mjs';
import {readSpreadsheet} from '../src/io.mjs';

const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','pivot-format-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'}));
const page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(20000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const report=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets[1]);
const refresh=async()=>{await page.locator('#pivot-bar').getByRole('button',{name:'새로 고침',exact:true}).click();await settle();};
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
 const book=fixture(),target={id:'pivot',name:'피벗',charts:[]},p={...spec(),rows:['상품'],columns:[],filters:[{field:'지역',values:['동부']}],preserveFormatting:true};applyPivot(book,target,p);book.sheets.push(target);book.activeSheetId=target.id;book.name='Pivot formatting';
 await page.locator('#file-input').setInputFiles({name:'pivot-format.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Pivot formatting');await settle();await go('B2');await page.keyboard.press('ControlOrMeta+1');
 await page.getByRole('tab',{name:'표시 형식',exact:true}).click();await page.locator('#format-format').selectOption('custom');await page.locator('#format-numberFormat').fill('0.000 "원"');await page.getByRole('tab',{name:'글꼴',exact:true}).click();await page.locator('#format-color').fill('#aa2255');await page.getByRole('button',{name:'적용',exact:true}).click();await settle();
 const handle=await page.locator('[data-resize-col="1"]').boundingBox();await page.mouse.move(handle.x+handle.width/2,handle.y+handle.height/2);await page.mouse.down();await page.mouse.move(handle.x+handle.width/2+40,handle.y+handle.height/2,{steps:5});await page.mouse.up();await settle();const before=await report();
 assert.equal(before.cells.B2.style.numberFormat,'0.000 "원"');assert.equal(before.columnWidths[1],230);
 await page.locator('#pivot-bar').getByRole('button',{name:'원본 보기',exact:true}).click();await go('C2');await page.keyboard.press('F2');await page.locator('#editor').fill('40');await page.locator('#editor').press('Enter');await settle();await page.locator('[data-sheet]').filter({hasText:'피벗'}).click();await refresh();assert.equal((await report()).cells.B2.raw,'60');assert.deepEqual(JSON.parse(JSON.stringify((await report()).cells.B2.style)),JSON.parse(JSON.stringify(before.cells.B2.style)));assert.equal((await report()).columnWidths[1],230);
 await page.locator('#undo').click();await settle();assert.equal((await report()).cells.B2.raw,'30');await page.locator('#redo').click();await settle();assert.equal((await report()).cells.B2.raw,'60');
 await page.locator('#pivot-bar').getByRole('button',{name:'필드 편집',exact:true}).click();await page.locator('#pivot-preserve-format').uncheck();await page.getByRole('dialog').getByRole('button',{name:'적용',exact:true}).click();await page.getByRole('dialog').waitFor({state:'detached'});await settle();assert.equal((await report()).columnWidths[1],190);assert.notEqual((await report()).cells.B2.style.color,'#aa2255');await page.locator('#undo').click();await settle();assert.equal((await report()).columnWidths[1],230);
 const file=path.join(profile,'styled.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');const read=await readSpreadsheet(await fs.readFile(file),'styled.xlsx');assert.ok(read.sheets[1].pivot);assert.equal(read.sheets[1].cells.B2.style.numberFormat,'0.000 "원"');assert.equal(read.sheets[1].cells.B2.style.color.toLowerCase(),'#aa2255');assert.equal(Math.round(read.sheets[1].columnWidths[1]),230);
 await page.locator('#file-input').setInputFiles(file);await settle();assert.equal((await report()).cells.B2.raw,'60');assert.equal((await report()).cells.B2.style.numberFormat,'0.000 "원"');await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('#pivot-bar').waitFor();assert.equal((await report()).cells.B2.style.numberFormat,'0.000 "원"');await page.screenshot({path:path.join(root,'test-results/pivot-format.png')});assert.deepEqual(errors,[]);
 console.log('PASS: user-set pivot format/color/width survive source edits, refresh, undo/redo, opt-out/reset undo, filtered XLSX save/reopen and recovery.');
}catch(error){await page.screenshot({path:path.join(root,'test-results/pivot-format-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
