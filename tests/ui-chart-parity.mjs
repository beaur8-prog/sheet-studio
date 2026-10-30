import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import platform from '../scripts/platform.cjs';
import {createWorkbook,address} from '../src/engine.mjs';
import {writeSpreadsheet} from '../src/io.mjs';

const root=path.resolve(import.meta.dirname,'..');
const dataDir=path.join(root,'.package','chart-parity-'+Date.now());
const env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch(desktopLaunch(root,env));
const page=await app.firstWindow(),errors=[];
page.setDefaultTimeout(10000);page.on('pageerror',error=>errors.push(error.message));
await page.setViewportSize({width:1280,height:900});
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));};
const go=async range=>{await page.locator('#name-box').fill(range);await page.locator('#name-box').press('Enter');await settle();};
const workbook=rows=>{const b=createWorkbook();rows.forEach((row,r)=>row.forEach((value,c)=>{if(value!==null)b.sheets[0].cells[address(r,c)]={raw:String(value)};}));return b;};
let sequence=0;
async function load(book,xlsx=false){
  const name='chart-parity-'+ ++sequence;
  const buffer=xlsx?Buffer.from((await writeSpreadsheet(book)).data):Buffer.from(JSON.stringify({...book,name}));
  await page.locator('#file-input').setInputFiles({name:name+(xlsx?'.xlsx':'.sheet.json'),mimeType:xlsx?'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':'application/json',buffer});
  await page.waitForFunction(name=>sheetStudio.getWorkbook().name===name,name);await settle();
}
async function insert(range,type,grouping='clustered',configure=async()=>{}){
  await go(range);await page.locator('[data-tab="삽입"]').click();await page.locator('[data-action="chart-insert"]').click();
  await page.locator('#chart-type').selectOption(type);await page.locator('#chart-grouping').selectOption(grouping);await configure();
  await page.getByRole('dialog').getByRole('button',{name:'삽입',exact:true}).click();await settle();
}
const points=()=>page.locator('.chart-card>svg .chart-point');
const expectPoints=async n=>{await page.waitForFunction(n=>document.querySelectorAll('.chart-card>svg .chart-point').length===n,n);};
async function roundtrip(){
  const filename=path.join(dataDir,'roundtrip-'+sequence+'.xlsx');
  await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},filename);
  await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('.toast')?.textContent==='파일을 저장했습니다.');
  const bytes=await fs.readFile(filename);await page.locator('#file-input').setInputFiles(filename);await settle();
  return JSZip.loadAsync(bytes);
}
try{
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.locator('.cell').first().waitFor();
  const formulas=workbook([['항목','매출'],['가','=40'],['나',10],['다',30]]);
  await load(formulas,true);await insert('A1:B4','bar');await expectPoints(3);
  await go('B2');await page.keyboard.press('F2');await page.locator('#editor').fill('=1/0');await page.locator('#editor').press('Enter');await settle();await expectPoints(2);
  assert.doesNotMatch((await points().locator('title').allTextContents()).join('\n'),/: 40/);
  await page.locator('#undo').click();await settle();await expectPoints(3);await page.locator('#redo').click();await settle();await expectPoints(2);

  await load(workbook([['항목','매출'],['가',10],['나',999],['다',30]]));await insert('A1:B4','bar');await expectPoints(3);
  await page.locator('[data-tab="데이터"]').click();await go('A1:B4');await page.locator('[data-action="filter"]').click();await settle();
  await page.locator('[data-filter-column="0"]').click();await page.getByRole('checkbox',{name:'나',exact:true}).uncheck();await page.locator('#filter-ok').click();await settle();await expectPoints(2);
  assert.doesNotMatch((await points().locator('title').allTextContents()).join('\n'),/999/);
  await page.locator('[data-edit-chart]').click();await page.locator('#chart-hidden-data').check();await page.getByRole('button',{name:'적용',exact:true}).click();await settle();await expectPoints(3);
  let zip=await roundtrip();assert.match(await zip.file(/Chart1\.xml$/)[0].async('string'),/plotVisOnly val="0"/);await expectPoints(3);

  await load(workbook([['항목','매출','비율'],['가',100,10],['나',50,5]]));await insert('A1:C3','combo','stacked');await expectPoints(4);
  const contained=await page.locator('.chart-card>svg').evaluate(svg=>[...svg.querySelectorAll('circle.chart-point')].every(point=>{const y=+point.getAttribute('cy');return y>=0&&y<=svg.viewBox.baseVal.height;}));
  assert.equal(contained,true);await page.screenshot({path:path.join(root,'test-results/chart-parity-combo.png')});

  await load(workbook([['X','Y1','크기1','Y2','크기2'],[1,10,25,20,100],[2,15,25,25,100]]));await insert('A1:E3','bubble');await expectPoints(4);
  const radii=await points().evaluateAll(nodes=>nodes.map(node=>+node.getAttribute('r')));assert.ok(Math.abs(radii[2]/radii[0]-2)<1e-9);

  await load(workbook([[1,3]]));await insert('A1:B1','pie','clustered',()=>page.locator('#chart-labels').check());await expectPoints(2);
  assert.match(await page.locator('.chart-card>svg').textContent(),/25%.*75%/);
  zip=await roundtrip();const pieXML=await zip.file(/Chart1\.xml$/)[0].async('string');assert.match(pieXML,/showPercent val="1"/);assert.match(pieXML,/showVal val="0"/);
  await expectPoints(2);assert.match(await page.locator('.chart-card>svg').textContent(),/25%.*75%/);

  const dates=workbook([['날짜','비율'],[46032,.5],[46022,.25],[46033,.75]]);
  for(const ref of ['A2','A3','A4'])dates.sheets[0].cells[ref].style={format:'date'};
  for(const ref of ['B2','B3','B4'])dates.sheets[0].cells[ref].style={format:'percent'};
  await load(dates);await insert('A1:B4','line','clustered',async()=>{await page.locator('#chart-first-column').check();await page.locator('#chart-labels').check();});await expectPoints(3);
  const dateSVG=await page.locator('.chart-card>svg').textContent();assert.match(dateSVG,/25\.00%/);assert.doesNotMatch(dateSVG,/46022/);
  assert.equal(await page.locator('.chart-legend-item').count(),1);
  const xs=await points().evaluateAll(nodes=>nodes.map(node=>+node.getAttribute('cx')));assert.ok(Math.abs((xs[1]-xs[0])/(xs[2]-xs[1])-10)<1e-9);
  const axisLabels=await page.locator('.chart-axis-category').evaluateAll(nodes=>nodes.map(node=>{const r=node.getBoundingClientRect();return {left:r.left,right:r.right};}));
  assert.ok(axisLabels.every((r,i)=>!i||r.left>=axisLabels[i-1].right));
  await page.screenshot({path:path.join(root,'test-results/chart-parity-dates.png')});
  assert.deepEqual(errors,[]);
  console.log('PASS: chart formula error/edit/undo/redo, filter visibility and inclusion option, XLSX reload, stacked combination, shared bubble size, pie percentage labels, formatted date axis intervals and single-series legend');
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/chart-parity-failure.png')});throw error;}
finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
