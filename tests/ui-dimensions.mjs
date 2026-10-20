import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),version=createRequire(import.meta.url)('../package.json').version,packaged=process.argv.includes('--packaged');
const env={...process.env,SHEET_STUDIO_DATA:path.join(root,'.package/dimension-ui-'+Date.now())};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix()),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(10000);
const frame=()=>page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
const model=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets[0]);
const go=async addr=>{await page.locator('#name-box').fill(addr);await page.locator('#name-box').press('Enter');await frame();};
const undo=async()=>{await page.locator('#undo').click();await frame();};
const redo=async()=>{await page.locator('#redo').click();await frame();};
async function start(axis,index){
  await frame();const rect=await page.locator(`[data-resize-${axis}="${index}"]`).boundingBox();assert.ok(rect);
  const point={x:rect.x+rect.width/2,y:rect.y+rect.height/2};
  // Both sides of the visible boundary must hit the resize handle, not its neighbor's header.
  for(const delta of [-3,3]){
    const p={x:point.x+(axis==='col'?delta:0),y:point.y+(axis==='row'?delta:0)};
    await page.mouse.move(p.x,p.y);
    const hit=await page.evaluate(({x,y})=>{const el=document.elementFromPoint(x,y);return {col:el?.dataset.resizeCol,row:el?.dataset.resizeRow,cursor:getComputedStyle(el).cursor};},p);
    assert.equal(hit[axis],String(index));assert.equal(hit.cursor,axis+'-resize');
  }
  await page.mouse.move(point.x,point.y);await page.mouse.down();return point;
}
async function move(point,dx,dy){await page.mouse.move(point.x+dx,point.y+dy,{steps:7});await frame();}
try{
  await page.locator('.cell').first().waitFor();
  const book=createWorkbook();book.name='Cell size verification';const sheet=book.sheets[0];
  sheet.cells.A1={raw:'열 너비를 마우스로 조절합니다',style:{bold:true}};sheet.cells.B2={raw:'두 줄\n내용',style:{wrap:true}};
  sheet.columnWidths={0:120,1:90,2:0,3:140};sheet.rowHeights={0:30,1:32,3:0};
  await page.locator('#file-input').setInputFiles({name:'dimensions.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});
  await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Cell size verification');await go('B2');
  let point=await start('col',0);await move(point,70,0);
  assert.equal((await model()).columnWidths[0],190);assert.equal(await page.locator('.dimension-guide-label').textContent(),'열 너비: 190px');
  assert.equal(await page.locator('#grid-shell').evaluate(el=>el.hasPointerCapture(1)),true);
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/dimension-column-preview.png')});
  await page.mouse.up();await frame();assert.equal(await page.locator('#name-box').inputValue(),'B2');
  assert.equal(Math.round((await page.getByRole('gridcell',{name:'A1',exact:true}).boundingBox()).width),190);
  await undo();assert.equal((await model()).columnWidths[0],120);await redo();assert.equal((await model()).columnWidths[0],190);
  point=await start('row',1);await move(point,0,43);assert.equal((await model()).rowHeights[1],75);
  await page.screenshot({path:path.join(root,'test-results/dimension-row-preview.png')});await page.mouse.up();await frame();
  await undo();assert.equal((await model()).rowHeights[1],32);
  point=await start('row',1);await move(point,0,50);await page.keyboard.press('Escape');await page.mouse.up();await frame();
  assert.equal((await model()).rowHeights[1],32);await redo();assert.equal((await model()).rowHeights[1],75);
  point=await start('col',1);await page.mouse.up();await frame();await undo();assert.equal((await model()).rowHeights[1],32);await redo();
  point=await start('col',1);await move(point,40,0);await move(point,0,0);await page.mouse.up();await frame();await undo();assert.equal((await model()).rowHeights[1],32);await redo();
  // The second click must still reach the handle after a no-op pointer gesture.
  await page.locator('[data-resize-col="0"]').dblclick();await frame();assert.ok((await model()).columnWidths[0]>190);
  await undo();assert.equal((await model()).columnWidths[0],190);
  await page.locator('[data-resize-row="1"]').dblclick();await frame();assert.notEqual((await model()).rowHeights[1],75);await undo();
  await page.locator('#zoom').evaluate(el=>{el.value='150';el.dispatchEvent(new Event('input',{bubbles:true}));});await frame();
  point=await start('col',1);await move(point,60,0);await page.mouse.up();await frame();assert.equal((await model()).columnWidths[1],130);
  await page.getByRole('button',{name:'보기',exact:true}).click();await page.locator('[data-action="freeze-menu"]').click();await page.locator('[data-action="freeze-selection"]').click();
  await page.locator('#grid').evaluate(el=>{el.scrollLeft=150;el.scrollTop=120;});await frame();
  point=await start('col',0);await move(point,30,0);await page.mouse.up();await frame();assert.equal((await model()).columnWidths[0],210);
  point=await start('col',3);await move(point,45,0);await page.mouse.up();await frame();assert.equal((await model()).columnWidths[3],170);
  point=await start('row',0);await move(point,0,30);await page.mouse.up();await frame();assert.equal((await model()).rowHeights[0],50);
  assert.equal((await model()).columnWidths[2],0);assert.equal((await model()).rowHeights[3],0);
  point=await start('col',0);await move(point,60,0);await page.evaluate(()=>window.dispatchEvent(new Event('blur')));await page.mouse.up();await frame();assert.equal((await model()).columnWidths[0],210);
  assert.equal(await page.locator('.dimension-guide').isVisible(),false);assert.equal(await page.locator('body').getAttribute('data-dimension-resize'),null);
  await go('B2');await page.keyboard.press('F2');await page.locator('#editor').fill('편집 내용 유지');
  point=await start('col',0);await move(point,30,0);await page.mouse.up();await frame();
  assert.equal((await model()).cells.B2.raw,'편집 내용 유지');assert.equal((await model()).columnWidths[0],230);
  await undo();assert.equal((await model()).columnWidths[0],210);assert.equal((await model()).cells.B2.raw,'편집 내용 유지');
  const saved=await model();await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('.cell').first().waitFor();assert.deepEqual((await model()).columnWidths,saved.columnWidths);assert.deepEqual((await model()).rowHeights,saved.rowHeights);
  assert.deepEqual(errors,[]);console.log('Dimension UI passed: boundary hover/hit area, live column/row drag, capture across redraws, guide, undo/redo/no-op, Escape/blur cancel, double-click AutoFit, zoom, frozen/scrolled headers, hidden dimensions and reload.');
}finally{await app.close();}
