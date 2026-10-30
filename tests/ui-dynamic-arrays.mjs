import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','dynamic-arrays-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch(desktopLaunch(root,env)),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(15000);
await app.evaluate(({ipcMain})=>{ipcMain.removeHandler('sheet:write-clipboard');ipcMain.handle('sheet:write-clipboard',(_event,payload)=>{globalThis.arrayTestClipboard=payload;});});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();},value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
async function edit(ref,raw){await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();}
async function copy(cut=false){await page.locator('[data-tab="홈"]').click();await page.locator('[data-action="'+(cut?'cut':'copy')+'"]').click();await settle();return app.evaluate(()=>globalThis.arrayTestClipboard.text);}
async function paste(text){await page.locator('#grid').focus();await page.evaluate(text=>{const data=new DataTransfer();data.setData('text/plain',text);document.querySelector('#grid').dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));},text);await settle();}
async function save(){const file=path.join(dataDir,'dynamic.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.evaluate(()=>document.querySelector('.toast')?.remove());await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');return file;}
try{
 await page.locator('.cell').first().waitFor();const b=createWorkbook();b.name='dynamic-arrays';b.sheets[0].cells={A1:{raw:'=SEQUENCE($D$1,2)'},D1:{raw:'3'},F1:{raw:'=SUM(A1#)'},H1:{raw:'=SORT(UNIQUE({3;1;2;1}))'}};
 await page.locator('#file-input').setInputFiles({name:'dynamic-arrays.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='dynamic-arrays');await settle();
 assert.equal(await value('B3'),6);assert.equal(await value('F1'),21);assert.equal(await value('H3'),3);assert.equal(await page.getByRole('gridcell',{name:'B3',exact:true}).textContent(),'6');
 await go('B2');assert.equal(await page.locator('#formula').inputValue(),'=SEQUENCE($D$1,2)');assert.equal(await page.locator('#formula').evaluate(el=>el.readOnly),true);assert.equal(await page.locator('.array-outline').count(),1);
 await page.keyboard.press('Delete');await settle();assert.equal(await value('B2'),4);assert.match(await page.locator('.toast').textContent(),/배열/);
 await go('A1:B3');assert.match(await page.locator('#stats').textContent(),/개수: 6/);assert.match(await page.locator('#stats').textContent(),/합계: 21/);
 await edit('D1','2');assert.equal(await value('B2'),4);assert.equal(await value('B3'),'');assert.equal(await value('F1'),10);
 await page.locator('#undo').click();await settle();assert.equal(await value('B3'),6);await page.locator('#redo').click();await settle();assert.equal(await value('B3'),'');
 await go('A1:B2');const whole=await copy();await go('J1');await paste(whole);assert.equal(await value('K2'),4);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.J2.raw),'');
 await go('B2');const partial=await copy();await go('M1');await paste(partial);assert.equal(await value('M1'),4);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.M1.raw),'4');
 assert.equal(await value('K3'),'');
 await edit('A1','plain');assert.equal(await value('B2'),'');assert.equal(await value('F1'),'#REF!');await page.locator('#undo').click();await settle();assert.equal(await value('B2'),4);
 const file=await save(),zip=await JSZip.loadAsync(await fs.readFile(file)),xml=await zip.file('xl/worksheets/sheet1.xml').async('string');assert.match(xml,/<c r="A1"[^>]*cm="\d+"/);assert.match(xml,/_xlfn\._xlws\.SORT/);assert.ok(zip.file('xl/metadata.xml'));
 const out=new ExcelJS.Workbook();await out.xlsx.load(await fs.readFile(file));assert.equal(out.worksheets[0].getCell('B2').value,4);assert.equal(out.worksheets[0].getCell('K2').value,4);
 await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await value('B2'),4);assert.equal(await value('K2'),4);await go('K2');assert.equal(await page.locator('#formula').evaluate(el=>el.readOnly),true);
 await edit('D1','3');assert.equal(await value('B3'),6);assert.equal(await value('F1'),21);
 // Clearing only the anchor removes all derived followers without persisting them.
 await go('A1');await page.keyboard.press('Delete');await settle();assert.equal(await value('A2'),'');assert.equal(await value('B3'),'');await page.locator('#undo').click();await settle();assert.equal(await value('B3'),6);
 await edit('T1','=SUM(H1#)');await go('H1');const moved=await copy(true);await go('R1');await paste(moved);assert.equal(await value('H1'),'');assert.equal(await value('R3'),3);assert.equal(await value('T1'),6);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.T1.raw),'=SUM(R1#)');await page.locator('#undo').click();await settle();assert.equal(await value('H3'),3);assert.equal(await value('R3'),'');await page.locator('#redo').click();await settle();assert.equal(await value('R3'),3);
 // An occupied follower blocks the complete output and explains the cause.
 await edit('P2','blocked');await edit('P1','=SEQUENCE(3)');assert.equal(await value('P1'),'#SPILL!');assert.match(await page.getByRole('gridcell',{name:'P1',exact:true}).getAttribute('title'),/P2.*内容|P2.*내용/);await go('P2');await page.keyboard.press('Delete');await settle();assert.equal(await value('P3'),3);
 // Offscreen source values are planned in the worker, not from a partial
 // visible-cell cache. The reference overlay covers the complete spill.
 await edit('V1','=SEQUENCE(1000)');await go('V1001');await page.locator('[data-tab="수식"]').click();await page.locator('[data-action="autosum"]').click();await page.locator('.workbook-busy').waitFor({state:'detached'});await page.waitForFunction(()=>document.querySelector('#editor').value==='=SUM(V1:V1000)');assert.equal(await page.locator('#editor').inputValue(),'=SUM(V1:V1000)');await page.locator('#editor').press('Enter');await settle();assert.equal(await value('V1001'),500500);
 await edit('X1','=SUM(V1#)');await go('X1');await page.keyboard.press('F2');await page.locator('#editor').press('End');await page.locator('#editor').press('ArrowLeft');await page.keyboard.press('F4');assert.equal(await page.locator('#editor').inputValue(),'=SUM($V$1#)');assert.equal(await page.locator('.formula-reference-box[data-reference="$V$1#"]').count(),1);await page.keyboard.press('Escape');await settle();
 // Unsupported imported dynamic formulas keep labelled caches; an edit makes
 // those values visibly stale while preserving their original file payload.
 const future=await JSZip.loadAsync(await fs.readFile(path.join(root,'tests/fixtures/xlsxwriter-dynamic.xlsx'))),part='xl/worksheets/sheet1.xml';future.file(part,(await future.file(part).async('string')).replace('_xlfn.SEQUENCE(3,2)','_xlfn.UNKNOWN_DYNAMIC(3,2)'));
 await page.locator('#file-input').setInputFiles({name:'future-array.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(await future.generateAsync({type:'uint8array'}))});await settle();await go('B3');assert.equal(await page.getByRole('gridcell',{name:'B3',exact:true}).textContent(),'6');assert.match(await page.getByRole('gridcell',{name:'B3',exact:true}).getAttribute('title'),/저장된 계산값/);await edit('J1','changed');await go('B3');assert.equal(await page.getByRole('gridcell',{name:'B3',exact:true}).textContent(),'#NAME?');assert.equal(await value('D1'),'#NAME?');
 // Real-world search/type/error masks must inspect every source element.
 const predicates=createWorkbook();predicates.name='Array predicates';predicates.sheets[0].cells={A2:{raw:'apple'},B2:{raw:'2'},A3:{raw:'pear'},B3:{raw:"'2"},A4:{raw:'APPLE'},B4:{raw:'#N/A'},A5:{raw:'banana'},B5:{raw:'TRUE'},A6:{raw:"'"},B6:{raw:'',style:{background:'#abcdef'}},D1:{raw:'=FILTER(A2:B6,ISNUMBER(SEARCH("apple",A2:A6)),"none")'},F1:{raw:'=FILTER(A2:B6,ISNUMBER(B2:B6),"none")'},J1:{raw:'=ISLOGICAL(B2:B6)'},L1:{raw:'=NOT(ISERROR(B2:B6))'},N1:{raw:'=FILTER(A2:B6,NOT(ISERROR(B2:B6)))'},R1:{raw:'=ISBLANK(B2:B6)'},T1:{raw:'=TYPE(B2:B6)'}};
 await page.locator('#file-input').setInputFiles({name:'array-predicates.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(predicates))});await settle();
 assert.equal(await value('D2'),'APPLE');assert.equal(await value('E2'),'#N/A');assert.equal(await value('F1'),'apple');assert.equal(await value('F2'),'');assert.equal(await value('J4'),true);assert.equal(await value('L3'),false);assert.equal(await value('N3'),'banana');assert.equal(await value('O3'),true);assert.equal(await value('R5'),true);assert.equal(await value('T1'),64);
 await go('D2');assert.equal(await page.locator('#formula').evaluate(el=>el.readOnly),true);
 await edit('A4','plum');assert.equal(await value('D2'),'');await page.locator('#undo').click();await settle();assert.equal(await value('D2'),'APPLE');await page.locator('#redo').click();await settle();assert.equal(await value('D2'),'');await page.locator('#undo').click();await settle();
 await edit('B3','3');assert.equal(await value('F2'),'pear');assert.equal(await value('G2'),3);
 const predicatesFile=await save();await page.locator('#file-input').setInputFiles(predicatesFile);await settle();assert.equal(await value('F2'),'pear');assert.equal(await value('G2'),3);assert.equal(await value('D2'),'APPLE');assert.equal(await value('E2'),'#N/A');assert.equal(await value('J4'),true);assert.equal(await value('R5'),true);
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});await go('A1:B3');await page.screenshot({path:path.join(root,'test-results/dynamic-arrays.png')});assert.deepEqual(errors,[]);
 console.log('PASS: Electron spill display/readonly ghost formula/outline/stats; resize, undo/redo, partial-edit protection, isolated copies, anchor replacement/removal, typed/search/error FILTER masks, fresh XLSX caches/metadata and reopen');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/dynamic-arrays-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
