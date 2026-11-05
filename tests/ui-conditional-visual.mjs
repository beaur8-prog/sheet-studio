import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','conditional-visual-'+Date.now());
const app=await _electron.launch(desktopLaunch(root,{...process.env,SHEET_STUDIO_DATA:profile})),page=await app.firstWindow(),errors=[];page.on('pageerror',error=>errors.push(error.message));page.setDefaultTimeout(15000);await page.setViewportSize({width:1280,height:900});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const cell=ref=>page.getByRole('gridcell',{name:ref,exact:true});
const style=(ref,property)=>cell(ref).evaluate((el,key)=>getComputedStyle(el)[key],property);
const manage=async()=>{await page.locator('[data-tab="홈"]').click();await page.locator('[data-action="conditional"]').click();await page.getByRole('dialog',{name:'조건부 서식 규칙 관리'}).waitFor();};
const newRule=async(range,type)=>{await go(range);await manage();await page.getByRole('button',{name:'새 규칙',exact:true}).click();await page.locator('#cf-type').selectOption(type);};
const close=()=>page.getByRole('dialog').getByRole('button',{name:'닫기',exact:true}).click();
const confirm=async()=>{await page.getByRole('dialog').getByRole('button',{name:'확인',exact:true}).click();await page.getByRole('dialog',{name:'조건부 서식 규칙 관리'}).waitFor();await close();await settle();};
const edit=async(ref,raw)=>{await go(ref);await page.keyboard.press('F2');await page.locator('#editor').fill(raw);await page.locator('#editor').press('Enter');await settle();};
let sequence=0;async function save(){const file=path.join(profile,'conditional-'+ ++sequence+'.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');return file;}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();const book=createWorkbook(),sheet=book.sheets[0];book.name='Visual rules';for(const col of ['A','B'])[-20,0,20,10,'','=5+5'].forEach((raw,i)=>sheet.cells[col+(i+1)]={raw:String(raw)});sheet.cells.C1={raw:'=SEQUENCE(3)'};
  await page.locator('#file-input').setInputFiles({name:'visual.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Visual rules');await settle();
  await newRule('A1:A6','colorScale');await page.locator('#cf-scale-count').selectOption('3');for(const [i,color] of [[0,'#ff0000'],[1,'#ffffff'],[2,'#00ff00']])await page.locator('#cf-point-color-'+i).fill(color);await page.locator('#cf-point-type-1').selectOption('percent');await page.locator('#cf-point-value-1').fill('50');await confirm();
  assert.equal(await style('A1','backgroundColor'),'rgb(255, 0, 0)');assert.equal(await style('A2','backgroundColor'),'rgb(255, 255, 255)');assert.equal(await style('A3','backgroundColor'),'rgb(0, 255, 0)');
  await newRule('B1:B6','dataBar');await page.locator('#cf-bar-color').fill('#638ec6');await confirm();assert.match(await style('B1','backgroundSize'),/^10%/);assert.match(await style('B2','backgroundSize'),/^50%/);assert.match(await style('B3','backgroundSize'),/^90%/);assert.match(await style('B1','backgroundImage'),/linear-gradient/);assert.equal(await style('B5','backgroundImage'),'none');
  await newRule('C1:C3','equal');await page.locator('#cf-value').fill('2');await page.locator('#cf-fill').fill('#0000ff');await confirm();assert.equal(await style('C2','backgroundColor'),'rgb(0, 0, 255)');
  await edit('B3','40');assert.match(await style('B2','backgroundSize'),/^36\.666/);await page.locator('#undo').click();await settle();assert.match(await style('B2','backgroundSize'),/^50%/);
  await manage();await page.getByRole('button',{name:'모두 지우기',exact:true}).click();await close();await settle();assert.equal(await style('B2','backgroundImage'),'none');await page.locator('#undo').click();await settle();assert.match(await style('B2','backgroundImage'),/linear-gradient/);
  const file=await save();await page.locator('#file-input').setInputFiles(file);await settle();assert.equal(await style('A3','backgroundColor'),'rgb(0, 255, 0)');assert.match(await style('B2','backgroundSize'),/^50%/);assert.equal(await style('C2','backgroundColor'),'rgb(0, 0, 255)');
  await page.setViewportSize({width:1000,height:680});await go('A1:A6');await manage();await page.locator('[data-rule-edit="2"]').click();await page.screenshot({path:path.join(root,'test-results/conditional-visual.png')});await close();
  const source=new ExcelJS.Workbook(),s=source.addWorksheet('Advanced');s.getCell('A1').value=-10;s.getCell('A2').value=20;s.addConditionalFormatting({ref:'A1:A2',rules:[{type:'dataBar',priority:1,cfvo:[{type:'min'},{type:'max'}],color:{argb:'FF638EC6'},gradient:false,negativeFillColor:{argb:'FFFF0000'},axisColor:{argb:'FF000000'}}]});
  await page.locator('#file-input').setInputFiles({name:'advanced.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(await source.xlsx.writeBuffer())});await settle();await go('A1');await page.locator('[data-tab="홈"]').click();await page.locator('[data-action="insert-row"]').click();await settle();
  for(let i=0;i<3;i++){const file=await save(),zip=await JSZip.loadAsync(await fs.readFile(file)),text=await zip.file('xl/worksheets/sheet1.xml').async('string');assert.match(text,/<x14:dataBar/);assert.match(text,/FFFF0000/);assert.match(text,/<xm:sqref>A2:A3<\/xm:sqref>/);await page.locator('#file-input').setInputFiles(file);await settle();}
  assert.deepEqual(errors,[]);await fs.writeFile(path.join(root,'test-results/conditional-visual.json'),JSON.stringify({scenarios:['three-color scale','base gradient bars','negative/zero/blank','spill conditional result','recalculation and undo','clear and undo','small layout','XLSX reopen','advanced x14 preservation, shifted range, three saves'],pageErrors:errors},null,2));console.log('PASS: visual rule authoring, grid colors and bars, spill formatting, undo, XLSX reopen and advanced bar definitions through shifted three-round saves');
}catch(error){await page.screenshot({path:path.join(root,'test-results/conditional-visual-failure.png')});throw error;}
finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
