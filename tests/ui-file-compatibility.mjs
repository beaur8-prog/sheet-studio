import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createHash} from 'node:crypto';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {desktopLaunch} from './electron-launch.mjs';
import {parseWorkbook} from '../src/engine.mjs';
import {readSpreadsheet} from '../src/io.mjs';

const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','compatibility-'+Date.now());await fs.mkdir(dataDir,{recursive:true});
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:dataDir,STUDIO_DISABLE_UPDATES:'1'}));
const page=await app.firstWindow(),errors=[];page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(20000);
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const digest=file=>fs.readFile(file).then(b=>createHash('sha256').update(b).digest('hex'));
const saveTarget=file=>app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async(_win,options)=>{globalThis.lastCopySaveOptions=options;return {canceled:false,filePath:file};};},file);
const chooseCopy=async()=>{await page.locator('#save').click();await page.getByRole('dialog',{name:'XLSX 사본 저장 확인',exact:true}).getByRole('button',{name:'기능 제외 XLSX 사본',exact:true}).click();await settle();};
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  const original=path.join(dataDir,'중요 원본.xlsx'),copy=path.join(dataDir,'안전한 사본.xlsx'),project=path.join(dataDir,'작업.sheet.json');
  const source=new ExcelJS.Workbook();source.addWorksheet('문서').getCell('A1').value='원본 개체 유지';const zip=await JSZip.loadAsync(await source.xlsx.writeBuffer());zip.file('xl/embeddings/source.bin',new Uint8Array([0,255,8,99]));
  await fs.writeFile(original,await zip.generateAsync({type:'nodebuffer'}));const before=await digest(original);
  await app.evaluate(({dialog},file)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[file]});},original);await page.keyboard.press('ControlOrMeta+o');await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='중요 원본');await settle();
  await page.locator('#compatibility-status').click();await page.getByRole('dialog',{name:'파일 호환성',exact:true}).getByText('첨부·삽입 개체',{exact:true}).waitFor();await page.screenshot({path:path.join(root,'test-results/file-compatibility-report.png')});await page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();
  await page.locator('#save').click();await page.getByRole('dialog',{name:'XLSX 사본 저장 확인',exact:true}).getByRole('button',{name:'취소',exact:true}).click();assert.equal(await digest(original),before);
  await saveTarget(original);await chooseCopy();await page.locator('.toast').filter({hasText:'원본 파일에는 덮어쓸 수 없습니다'}).waitFor();assert.equal(await digest(original),before);
  await saveTarget(copy);await chooseCopy();await page.locator('.toast').filter({hasText:'기능을 제외한 XLSX 사본을 저장했습니다.'}).waitFor();assert.equal(await digest(original),before);
  const options=await app.evaluate(()=>globalThis.lastCopySaveOptions);assert.match(options.title,/사본/);assert.match(path.basename(options.defaultPath),/-사본\.xlsx$/);
  const loaded=await readSpreadsheet(await fs.readFile(copy),'copy.xlsx');assert.equal(loaded.sheets[0].cells.A1.raw,'원본 개체 유지');assert.equal(loaded.xlsxCompatibility,undefined);
  const out=await JSZip.loadAsync(await fs.readFile(copy));assert.equal(out.file('xl/embeddings/source.bin'),null);assert.ok((await JSZip.loadAsync(await fs.readFile(original))).file('xl/embeddings/source.bin'));
  await saveTarget(project);await page.locator('#save').click();await page.getByRole('dialog',{name:'XLSX 사본 저장 확인',exact:true}).getByRole('button',{name:'앱 문서 저장',exact:true}).click();await page.locator('.toast').filter({hasText:'파일을 저장했습니다.'}).waitFor();assert.equal(parseWorkbook(await fs.readFile(project,'utf8')).xlsxCompatibility.risks[0].kind,'ole');
  assert.equal(await digest(original),before);assert.deepEqual(errors,[]);
  console.log('PASS: compatibility report; warning/cancel; original overwrite blocked; explicit feature-loss XLSX copy; project save retains review; original binary/hash preserved.');
}catch(error){await page.screenshot({path:path.join(root,'test-results/file-compatibility-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
