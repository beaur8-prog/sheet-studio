import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),version=createRequire(import.meta.url)('../package.json').version,packaged=process.argv.includes('--packaged');
const env={...process.env,SHEET_STUDIO_DATA:path.join(root,'.package/cell-editing-'+Date.now())};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(10000);
const frame=()=>page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
const cell=a=>page.getByRole('gridcell',{name:a,exact:true});
const raw=a=>page.evaluate(a=>sheetStudio.getWorkbook().sheets[0].cells[a]?.raw||'',a);
const go=async a=>{await page.locator('#name-box').fill(a);await page.locator('#name-box').press('Enter');await frame();};
async function doubleClick(a,{x=45,y}={}){
  await frame();const box=await cell(a).boundingBox();await page.mouse.move(box.x+x,box.y+(y??box.height/2));
  // Keep each button press long enough for the virtualized cells to be replaced.
  for(const count of [1,2]){await page.mouse.down({clickCount:count});await page.waitForTimeout(70);await page.mouse.up({clickCount:count});if(count===1)await page.waitForTimeout(100);}
  await frame();assert.equal(await page.locator('#editor').isVisible(),true);assert.equal(await page.evaluate(()=>document.activeElement.id),'editor');
}
try{
  await page.locator('.cell').first().waitFor();const book=createWorkbook(),s=book.sheets[0];book.name='Inline cell editing verification';
  s.cells={A1:{raw:'Hello world',style:{fontSize:20,bold:true,color:'#ffffff',background:'#174d32'}},B2:{raw:'12345',style:{align:'right'}},C3:{raw:'=SUM(10,20)'},D4:{raw:'병합된 셀 직접 편집',style:{italic:true,fontSize:16}},A8:{raw:'한글 편집'}};
  s.columnWidths={0:250,1:140,2:160};s.rowHeights={0:48,3:35,4:35};s.merges=['D4:E5'];
  await page.locator('#file-input').setInputFiles({name:'inline.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Inline cell editing verification');await frame();
  await doubleClick('A1');const offset=await page.locator('#editor').evaluate(el=>el.selectionStart);assert.ok(offset>0&&offset<'Hello world'.length);
  const styles=await page.locator('#editor').evaluate(el=>{const s=getComputedStyle(el),c=getComputedStyle(document.querySelector('.cell[aria-label="A1"]'));return {cursor:s.cursor,caret:s.caretColor,color:s.color,background:s.backgroundColor,font:s.fontSize,cellFont:c.fontSize,weight:s.fontWeight};});
  assert.equal(styles.cursor,'text');assert.equal(styles.caret,styles.color);assert.equal(styles.color,'rgb(255, 255, 255)');assert.equal(styles.background,'rgb(23, 77, 50)');assert.equal(styles.font,styles.cellFont);assert.equal(styles.weight,'700');
  const editorBox=await page.locator('#editor').boundingBox(),cellBox=await cell('A1').boundingBox();assert.deepEqual(editorBox,cellBox);
  await page.keyboard.type('!');assert.equal(await page.locator('#editor').inputValue(),'Hello world'.slice(0,offset)+'!'+'Hello world'.slice(offset));assert.equal(await raw('A1'),'Hello world');
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/cell-inline-caret.png'),caret:'initial'});
  await page.locator('#editor').press('Enter');assert.equal(await page.locator('#editor').evaluate(el=>getComputedStyle(el).opacity),'0');assert.equal(await page.locator('#name-box').inputValue(),'A2');
  await page.locator('#undo').click();assert.equal(await raw('A1'),'Hello world');await page.locator('#redo').click();assert.ok((await raw('A1')).includes('!'));
  await doubleClick('B2');assert.equal(await page.locator('#editor').evaluate(el=>getComputedStyle(el).textAlign),'right');await page.locator('#editor').fill('999');await page.locator('#editor').press('Escape');assert.equal(await raw('B2'),'12345');
  await go('B2');await page.keyboard.press('F2');assert.equal(await page.locator('#editor').inputValue(),'12345');await page.keyboard.press('End');await page.keyboard.type('6');await page.locator('#editor').press('Tab');assert.equal(await raw('B2'),'123456');assert.equal(await page.locator('#name-box').inputValue(),'C2');
  await go('B2');await page.keyboard.type('ABC');assert.equal(await page.locator('#editor').inputValue(),'ABC');await page.locator('#editor').press('Enter');assert.equal(await raw('B2'),'ABC');
  await doubleClick('C3');assert.equal(await page.locator('#editor').inputValue(),'=SUM(10,20)');await page.locator('#editor').fill('=SUM(20,30)');await page.locator('#editor').press('Enter');assert.equal(await page.evaluate(()=>sheetStudio.evaluate('C3')),50);
  await doubleClick('D4');assert.deepEqual(await page.locator('#editor').boundingBox(),await cell('D4').boundingBox());await page.locator('#editor').fill('병합 편집 완료');await cell('A8').click();assert.equal(await raw('D4'),'병합 편집 완료');
  await doubleClick('A8');await page.locator('#editor').fill('한글 입력 중');await page.locator('#editor').dispatchEvent('keydown',{key:'Enter',isComposing:true});assert.equal(await raw('A8'),'한글 편집');await page.locator('#editor').press('Enter');assert.equal(await raw('A8'),'한글 입력 중');
  await go('B2');await page.getByRole('button',{name:'보기',exact:true}).click();await page.locator('[data-action="freeze-menu"]').click();await page.locator('[data-action="freeze-selection"]').click();
  await page.locator('#zoom').evaluate(el=>{el.value='150';el.dispatchEvent(new Event('input',{bubbles:true}));});await page.locator('#grid').evaluate(el=>{el.scrollLeft=120;el.scrollTop=120;});await frame();
  await doubleClick('A1');assert.deepEqual(await page.locator('#editor').boundingBox(),await cell('A1').boundingBox());assert.equal(await page.locator('#editor').evaluate(el=>getComputedStyle(el).fontSize),'40px');await page.locator('#editor').press('Escape');
  await page.reload();await cell('A1').waitFor();assert.equal(await raw('D4'),'병합 편집 완료');assert.equal(await raw('A8'),'한글 입력 중');assert.deepEqual(errors,[]);
  console.log('Inline editing UI passed: delayed double-click across redraws, clicked caret position, font/color/geometry, in-cell typing, Enter/Tab/Escape, undo/redo, F2/replacement typing, formulas, merged cells, IME guard, frozen/zoomed cells and reload.');
}finally{await app.close();}
