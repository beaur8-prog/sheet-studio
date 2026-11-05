import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import {desktopLaunch} from './electron-launch.mjs';
import {fixture,spec} from './fixtures/pivot.mjs';
import {applyPivot} from '../src/pivot.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';
import {nativePivotIndex} from '../src/native-pivots.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','native-pivots-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:dataDir,STUDIO_DISABLE_UPDATES:'1'}));
const page=await app.firstWindow(),errors=[];page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(20000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const value=ref=>page.evaluate(ref=>sheetStudio.evaluateAsync(ref),ref);
async function save(){const file=path.join(dataDir,'native-saved.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');return file;}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  const book=fixture(),target={id:'pivot',name:'피벗',charts:[]};applyPivot(book,target,spec());book.sheets.push(target);book.activeSheetId=target.id;
  const out=await writeSpreadsheet(book),zip=await JSZip.loadAsync(out.data),table='xl/pivotTables/pivotTable1.xml';
  zip.file(table,(await zip.file(table).async('string')).replace('subtotal="sum"','subtotal="average"'));
  zip.file('xl/worksheets/sheet2.xml',(await zip.file('xl/worksheets/sheet2.xml').async('string')).replace(/(<c r="D4"[^>]*>[\s\S]*?<v>)160(<\/v>)/,(_,a,b)=>a+'40'+b));
  await page.locator('#file-input').setInputFiles({name:'native-edited.xlsx',mimeType:out.mime,buffer:Buffer.from(await zip.generateAsync({type:'uint8array'}))});
  await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='native-edited');await settle();
  await page.locator('#pivot-bar').getByText('Excel 피벗 원본 보존',{exact:true}).waitFor();assert.equal(await value('D4'),40);
  await go('D4');await page.keyboard.press('Delete');await settle();assert.equal(await value('D4'),40);
  await page.keyboard.press('F2');assert.equal(await page.locator('#editor').evaluate(el=>el.classList.contains('entry-ready')),true);await page.keyboard.press('Escape');
  await page.locator('[data-tab="홈"]').click();await page.locator('[data-action="insert-row"]').click();await settle();assert.match(await page.locator('.toast').innerText(),/원본 피벗/);
  const file=await save(),saved=await fs.readFile(file),loaded=await readSpreadsheet(saved,'checked.xlsx'),savedZip=await JSZip.loadAsync(saved),entries=await nativePivotIndex(savedZip,loaded);
  assert.equal(entries.length,1);assert.match(await savedZip.file(entries[0].path).async('string'),/subtotal="average"/);assert.equal(loaded.sheets[1].cells.D4.raw,'40');
  await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('#pivot-bar').getByText('Excel 피벗 원본 보존',{exact:true}).waitFor();assert.equal(await value('D4'),40);
  await page.locator('#pivot-bar').getByRole('button',{name:'일반 범위로 변환',exact:true}).click();await page.getByRole('dialog').getByRole('button',{name:'변환',exact:true}).click();await settle();assert.equal(await page.locator('#pivot-bar').isVisible(),false);
  await go('D4');await page.keyboard.press('F2');await page.locator('#editor').fill('55');await page.locator('#editor').press('Enter');await settle();assert.equal(await value('D4'),55);
  await page.keyboard.press('ControlOrMeta+z');await settle();assert.equal(await value('D4'),40);await page.keyboard.press('ControlOrMeta+z');await settle();await page.locator('#pivot-bar').getByText('Excel 피벗 원본 보존',{exact:true}).waitFor();
  assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().xlsxPivotArchive.sheets.length),1);
  await page.screenshot({path:path.join(root,'test-results/native-pivots-preserved.png')});assert.deepEqual(errors,[]);
  console.log('PASS: external pivot layout/results preserved; native cache XLSX saved; edit/structure guarded; reload, explicit conversion and undo retain archive.');
}catch(error){await page.screenshot({path:path.join(root,'test-results/native-pivots-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
