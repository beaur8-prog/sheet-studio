import {_electron} from 'playwright';
import assert from 'node:assert/strict';import fs from 'node:fs/promises';import path from 'node:path';
import platform from '../scripts/platform.cjs';import {createRequire} from 'node:module';import {createWorkbook} from '../src/engine.mjs';
import ExcelJS from 'exceljs';
const root=path.resolve(import.meta.dirname,'..'),dataDir=path.join(root,'.package','studio-clipboard-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});const page=await app.firstWindow();page.setDefaultTimeout(8000);const errors=[];page.on('pageerror',e=>errors.push(e.message));
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const select=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
let saved;
try{
 await page.locator('.cell').first().waitFor();saved=await app.evaluate(({clipboard})=>({text:clipboard.readText(),html:clipboard.readHTML(),rtf:clipboard.readRTF(),image:clipboard.readImage().toDataURL()}));
 const b=createWorkbook();b.name='clipboard test';b.sheets[0].cells={A1:{raw:'보고서',style:{bold:true,background:'#fff2cc',borders:{bottom:{style:'double',color:'#ff0000'}}}},A2:{raw:'=1+2'},B2:{raw:'줄1\n줄2'}};b.sheets[0].merges=['A1:B1'];await page.locator('#file-input').setInputFiles({name:'clipboard.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='clipboard test');await select('A1:B2');
 await page.locator('[data-action=copy]').click();await page.getByText('선택한 셀을 복사했습니다.',{exact:true}).waitFor();
 const rich=await app.evaluate(({clipboard})=>({text:clipboard.readText(),html:clipboard.readHTML()}));assert.ok(rich.html.includes('colspan="2"'));assert.ok(rich.html.includes('double'));assert.ok(!rich.html.includes('=1+2'));assert.ok(rich.text.includes('"줄1\n줄2"'));
 await select('D1');await page.locator('[data-action=paste]').click();await page.waitForFunction(()=>sheetStudio.getWorkbook().sheets[0].cells.D2?.raw==='=1+2');assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.E2.raw),'줄1\n줄2');
 const fixture=JSON.parse(await fs.readFile(path.join(root,'tests/fixtures/studio-table.json'),'utf8'));await app.evaluate(({clipboard},fixture)=>clipboard.write(fixture),fixture);await select('G1');await page.locator('[data-action=paste]').click();await page.waitForFunction(()=>sheetStudio.getWorkbook().sheets[0].cells.G1?.raw==='보고서');const imported=await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0]);assert.ok(imported.merges.includes('G1:H1'));assert.equal(imported.cells.G2.raw,'줄1\n줄2');assert.equal(imported.cells.H2.raw,"'=literal");assert.equal(imported.cells.G1.style.borders.bottom.color,'#ff0000');
 await page.locator('#undo').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.G1),undefined);
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});
 const cases=[{name:'short-frame',points:[32,32,32,32]},{name:'unequal',points:[24,40,28,36],font:14},{name:'merged',points:[32,32,32,32],merged:true},{name:'taller-target',points:[32,32,32,32],target:80}];
 for(const scenario of cases){
  const heightBook=createWorkbook();heightBook.name='표 높이 '+scenario.name;heightBook.sheets[0].columnWidths={0:300};
  if(scenario.target)scenario.points.forEach((_,r)=>heightBook.sheets[0].rowHeights[r]=scenario.target);
  await page.locator('#file-input').setInputFiles({name:'height.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(heightBook))});await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,heightBook.name);await settle();
  const html='<table data-studio-table="1" style="height:72pt"><tbody>'+scenario.points.map((h,r)=>`<tr style="height:${h}pt">${scenario.merged&&r===1?'':`<td${scenario.merged&&r===0?' rowspan="2"':''} style="font-size:${scenario.font||18}pt">한글 ${r+1}행</td>`}</tr>`).join('')+'</tbody></table>';
  await app.evaluate(({clipboard},html)=>clipboard.write({text:'한글 표',html}),html);await select('A1');await page.locator('[data-action=paste]').click();await page.waitForFunction(()=>sheetStudio.getWorkbook().sheets[0].cells.A1?.raw==='한글 1행');await settle();
  const expected=scenario.points.map(h=>Math.max(h*4/3,scenario.target||25));
  const metrics=await page.evaluate(()=>{const s=sheetStudio.getWorkbook().sheets[0];return {rowHeights:s.rowHeights,merges:s.merges,cells:[...document.querySelectorAll('.cell[data-c="0"]')].filter(el=>Number(el.dataset.r)<4).map(el=>{const range=document.createRange();range.selectNodeContents(el);const text=range.getBoundingClientRect(),cell=el.getBoundingClientRect();return {r:Number(el.dataset.r),height:cell.height,overflow:Math.max(0,cell.top-text.top,text.bottom-cell.bottom)};})};});
  await fs.writeFile(path.join(root,`test-results/studio-table-height-${scenario.name}.json`),JSON.stringify(metrics,null,2));
  await page.screenshot({path:`test-results/studio-table-height-${scenario.name}.png`});
  expected.forEach((h,r)=>assert.ok(Math.abs(metrics.rowHeights[r]-h)<1e-7,`${scenario.name}: row ${r+1} height ${metrics.rowHeights[r]} != ${h}`));
  for(const cell of metrics.cells){const span=scenario.merged&&cell.r===0?2:1;assert.ok(Math.abs(cell.height-expected.slice(cell.r,cell.r+span).reduce((a,b)=>a+b,0))<0.1);assert.ok(cell.overflow<=1,`${scenario.name}: text clipped by ${cell.overflow}px`);}
  if(scenario.merged)assert.deepEqual(metrics.merges,['A1:A2']);
  await page.locator('#undo').click();await settle();assert.deepEqual(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].rowHeights),heightBook.sheets[0].rowHeights);assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.A1),undefined);
  await page.locator('#redo').click();await settle();await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('.cell').first().waitFor();await settle();
  const recovered=await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0]);expected.forEach((h,r)=>assert.ok(Math.abs(recovered.rowHeights[r]-h)<1e-7));
  if(scenario.name==='unequal'){
   const filename=path.join(dataDir,'height-roundtrip.xlsx');await app.evaluate(({dialog},filename)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:filename});},filename);
   await page.locator('#save').click();await page.getByText('파일을 저장했습니다.',{exact:true}).waitFor();assert.ok((await fs.stat(filename)).size>1000);
   await page.locator('#file-input').setInputFiles(filename);await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='height-roundtrip');await settle();
   const reopened=await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0]);expected.forEach((h,r)=>assert.ok(Math.abs(reopened.rowHeights[r]-h)<1e-7));
  }
 }
 const native=new ExcelJS.Workbook(),nativeSheet=native.addWorksheet('기본 행 높이');nativeSheet.properties.defaultRowHeight=32;nativeSheet.getColumn(1).width=40;
 for(let r=1;r<=4;r++){nativeSheet.getCell(`A${r}`).value=`한글 ${r}행`;nativeSheet.getCell(`A${r}`).font={size:18};}
 nativeSheet.getRow(2).height=40;nativeSheet.getRow(3).height=64;nativeSheet.getRow(3).hidden=true;
 await page.locator('#file-input').setInputFiles({name:'default-row-height.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(await native.xlsx.writeBuffer())});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='default-row-height');await settle();
 const defaultMetrics=await page.locator('.cell[data-r="0"][data-c="0"]').evaluate(el=>{const r=document.createRange();r.selectNodeContents(el);const cell=el.getBoundingClientRect(),text=r.getBoundingClientRect();return {height:cell.height,overflow:Math.max(0,cell.top-text.top,text.bottom-cell.bottom)};});
 assert.ok(Math.abs(defaultMetrics.height-32*4/3)<0.1);assert.ok(defaultMetrics.overflow<=1);assert.equal(await page.locator('.cell[data-r="2"][data-c="0"]').count(),0);
 await select('A1:A4');await page.locator('[data-action=copy]').click();await page.getByText('선택한 셀을 복사했습니다.',{exact:true}).waitFor();
 const defaultCopy=await app.evaluate(({clipboard})=>clipboard.readHTML());
 const copiedHeights=await page.evaluate(html=>{const template=document.createElement('template');template.innerHTML=html;return [...template.content.querySelectorAll('tr')].map(row=>parseFloat(row.style.height));},defaultCopy);
 await fs.writeFile(path.join(root,'test-results/studio-table-height-default-copy.json'),JSON.stringify({copiedHeights,expected:[32,40,32].map(pt=>pt*4/3)},null,2));
 assert.equal(copiedHeights.length,3);[32,40,32].forEach((pt,r)=>assert.ok(Math.abs(copiedHeights[r]-pt*4/3)<0.01,`${copiedHeights[r]} != ${pt*4/3}`));
 await select('A3');await page.locator('[data-action=cell-format-menu]').click();await page.locator('[data-action=unhide-row]').click();await settle();
 assert.ok(Math.abs(await page.locator('.cell[data-r="2"][data-c="0"]').evaluate(el=>el.getBoundingClientRect().height)-64*4/3)<0.1);
 await page.locator('#undo').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].rowHeights[2]),0);
 await page.screenshot({path:'test-results/studio-table-copy.png'});assert.deepEqual(errors,[]);console.log('PASS: OS clipboard HTML+TSV, own formulas/merged cells, external table import, row-height/merged text geometry, taller target preservation, atomic undo, recovery reload and XLSX roundtrip');
}finally{if(saved)await app.evaluate(({clipboard,nativeImage},data)=>clipboard.write({text:data.text,html:data.html,rtf:data.rtf,image:nativeImage.createFromDataURL(data.image)}),saved);await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
