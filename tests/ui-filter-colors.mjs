import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {colorFixture} from './fixtures/filter-colors.mjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const dataDir=path.join(root,'.package','filter-colors-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[],mod=process.platform==='darwin'?'Meta':'Control';page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(10000);await page.setViewportSize({width:1280,height:900});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const model=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());return page.evaluate(()=>sheetStudio.getWorkbook().sheets[0]);};
const visible=async()=>{const s=await model();return Array.from({length:7},(_,i)=>i+1).filter(r=>!s.autoFilter.hiddenRows.includes(r));};
const arrow=c=>page.locator(`[data-filter-column="${c}"]`),ok=async()=>{await page.locator('#filter-ok').click();await settle();};
const radio=(target,color)=>page.locator(`[data-color-target="${target}"][value="${color||''}"]`);
const choose=async(target,color)=>{const details=page.locator('.filter-color-menu');if(!await details.evaluate(el=>el.open))await details.locator('summary').click();await radio(target,color).check();};
try{
 await page.locator('.cell').first().waitFor();await fs.mkdir('test-results',{recursive:true});await page.locator('#file-input').setInputFiles({name:'colors.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(colorFixture()))});await arrow(0).waitFor();await page.locator('[data-tab="데이터"]').click();await settle();
 assert.equal(await page.locator('.cell[aria-label="A8"]').evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(255, 0, 0)');
 await arrow(0).click();await choose('background','#ff0000');assert.deepEqual(await visible(),[1,2,3,4,5,6,7]);await page.screenshot({path:'test-results/filter-colors-menu.png'});await ok();assert.deepEqual(await visible(),[1,3,6,7]);assert.equal(await arrow(0).getAttribute('class'),'filter-arrow is-filtered');
 await arrow(0).click();assert.equal(await radio('background','#ff0000').isChecked(),true);await ok();assert.deepEqual(await visible(),[1,3,6,7]);
 await arrow(0).click();await choose('color','#ff0000');await page.locator('#filter-cancel').click();assert.deepEqual(await visible(),[1,3,6,7]);
 await arrow(1).click();assert.equal(await page.getByRole('checkbox',{name:'20',exact:true}).count(),0);await page.locator('.filter-condition-menu summary').click();await page.locator('[data-condition="greaterThan"]').click();await page.locator('#filter-value-0').fill('20');await ok();assert.deepEqual(await visible(),[3,6,7]);
 await page.locator('#grid').focus();await page.keyboard.press(mod+'+z');await settle();assert.deepEqual(await visible(),[1,3,6,7]);await page.keyboard.press(mod+'+Shift+z');await settle();assert.deepEqual(await visible(),[3,6,7]);
 await arrow(0).click();assert.equal(await radio('background','#0000ff').count(),0);await choose('background',null);await ok();assert.deepEqual(await visible(),[4]);
 await arrow(1).click();await page.locator('#filter-clear').click();await arrow(0).click();await choose('color','#008000');await ok();assert.deepEqual(await visible(),[1,6,7]);
 await arrow(0).click();await choose('color',null);await ok();assert.deepEqual(await visible(),[3]);await arrow(0).click();await choose('background','#ffffff');await ok();assert.deepEqual(await visible(),[5]);
 await arrow(0).click();await choose('background','#ff0000');await page.keyboard.press('Escape');assert.deepEqual(await visible(),[5]);await arrow(0).click();await choose('background','#ff0000');await ok();assert.deepEqual(await visible(),[1,3,6,7]);
 const output=path.join(dataDir,'colors.xlsx');await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},output);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');await page.locator('#file-input').setInputFiles(output);await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='colors');await settle();assert.deepEqual(await visible(),[1,3,6,7]);assert.deepEqual((await model()).autoFilter.columns[0],{type:'color',target:'background',color:'#ff0000'});
 await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await arrow(0).waitFor();assert.deepEqual(await visible(),[1,3,6,7]);await arrow(0).click();assert.equal(await radio('background','#ff0000').isChecked(),true);await ok();assert.deepEqual(await visible(),[1,3,6,7]);
 // Change a conditional-format dependency, then explicitly reapply the color filter.
 await page.locator('#name-box').fill('B8');await page.locator('#name-box').press('Enter');await page.keyboard.press('F2');await page.locator('#editor').fill('1');await page.locator('#editor').press('Enter');assert.deepEqual(await visible(),[1,3,6,7]);await page.locator('[data-tab="데이터"]').click();await page.getByRole('button',{name:'다시 적용',exact:true}).click();assert.deepEqual(await visible(),[1,3,6]);
 await arrow(0).click();await page.locator('#filter-search').fill('사과');await ok();assert.equal((await model()).autoFilter.columns[0].type,'values');assert.deepEqual(await visible(),[1,2]);
 await page.getByRole('button',{name:'지우기',exact:true}).click();assert.deepEqual(await visible(),[1,2,3,4,5,6,7]);assert.equal(await page.locator('.filter-arrow').count(),3);assert.deepEqual(errors,[]);
 console.log('PASS: fill/font/no-fill/automatic swatches, conditional colors and styled blanks, confirm/cancel/Esc, other-column conditions, undo/redo, value-filter replacement, clear/reapply, native XLSX and reload');
}catch(error){await page.screenshot({path:'test-results/filter-colors-failure.png'});throw error;}finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
