import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged'),version=createRequire(import.meta.url)('../package.json').version;
const dataDir=path.join(root,'.package','ime-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];await page.setViewportSize({width:1280,height:900});page.on('pageerror',e=>errors.push(e.message));
const editor=page.locator('#editor'),name=page.locator('#name-box'),raw=a=>page.evaluate(a=>sheetStudio.getWorkbook().sheets[0].cells[a]?.raw,a);
const go=async a=>{await name.fill(a);await name.press('Enter');};
const ready=async()=>{assert.equal(await page.evaluate(()=>document.activeElement.id),'editor');assert.equal(await editor.evaluate(el=>el.readOnly),false);assert.equal(await editor.evaluate(el=>getComputedStyle(el).opacity),'0');};
const session=await page.context().newCDPSession(page);
const compose=text=>session.send('Input.imeSetComposition',{text,selectionStart:text.length,selectionEnd:text.length});
const confirm=text=>session.send('Input.insertText',{text});
try{
 await page.locator('.cell').first().waitFor();await fs.mkdir('test-results',{recursive:true});
 const b=createWorkbook();b.name='Native IME entry';b.sheets[0].cells={B1:{raw:'이전 값'},A4:{raw:'기존 '}};
 await page.locator('#file-input').setInputFiles({name:'ime.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Native IME entry');
 await go('A1');await page.keyboard.press('ArrowRight');assert.equal(await name.inputValue(),'B1');await ready();
 // These use Chromium's real composition/input path, not fill() or fabricated
 // CompositionEvents. There must be no focus change from the first jamo onward.
 await editor.evaluate(el=>{el.dataset.focusChanges='0';for(const type of ['blur','focus'])el.addEventListener(type,()=>el.dataset.focusChanges=String(+el.dataset.focusChanges+1));});
 for(const text of ['ㅎ','하','한','한ㄱ','한그','한글']){await compose(text);assert.equal(await editor.inputValue(),text);assert.equal(await raw('B1'),'이전 값');assert.equal(await page.evaluate(()=>document.activeElement.id),'editor');}
 assert.equal(await editor.getAttribute('data-focus-changes'),'0');await page.screenshot({path:'test-results/ime-first-hangul.png',caret:'initial'});
 await page.keyboard.press('ArrowRight');assert.equal(await raw('B1'),'이전 값');assert.equal(await name.inputValue(),'B1');
 await confirm('한글');await page.keyboard.press('Enter');assert.equal(await raw('B1'),'한글');assert.equal(await name.inputValue(),'B2');await ready();
 // No render delay is inserted between cell movement and composition.
 for(const [key,cell,value] of [['ArrowRight','C2','다음'],['Tab','D2','문자'],['Enter','D3','입력']]){
  await page.keyboard.press(key);assert.equal(await name.inputValue(),cell);await ready();await compose(value);await confirm(value);await page.keyboard.press('ArrowDown');assert.equal(await raw(cell),value);await ready();
  // Return to the committed cell so the next iteration moves from that origin.
  await page.keyboard.press('ArrowUp');
 }
 await go('A4');await page.keyboard.press('F2');await page.keyboard.press('End');await compose('한');await confirm('한글');assert.equal(await editor.inputValue(),'기존 한글');await page.keyboard.press('ArrowLeft');assert.equal(await name.inputValue(),'A4');await page.keyboard.press('Enter');assert.equal(await raw('A4'),'기존 한글');
 await go('A6');await page.keyboard.type('ABC ');await compose('한');await confirm('한글');await page.keyboard.press('Tab');assert.equal(await raw('A6'),'ABC 한글');await ready();
 await page.getByRole('gridcell',{name:'B6',exact:true}).click();await ready();await compose('취소');await compose('');await page.keyboard.press('Escape');assert.equal(await raw('B6'),undefined);await ready();
 await go('A8');await compose('이동');await page.getByRole('gridcell',{name:'B8',exact:true}).click();assert.equal(await raw('A8'),'이동');assert.equal(await name.inputValue(),'B8');await ready();assert.equal(await editor.inputValue(),'');await compose('다음');await confirm('다음');await page.keyboard.press('Enter');assert.equal(await raw('B8'),'다음');
 for(const [key,cell] of [['ArrowLeft','C10'],['ArrowRight','E10'],['ArrowUp','D9'],['ArrowDown','D11']]){await go('D10');await page.keyboard.press(key);await compose('ㅎ');await compose('한글');await confirm('한글');await page.keyboard.press('Enter');assert.equal(await raw(cell),'한글');}
 await page.evaluate(()=>sheetStudio.flushRecovery());await page.reload();await page.locator('.cell').first().waitFor();await go('B1');assert.equal(await raw('B1'),'한글');assert.equal(await raw('A4'),'기존 한글');assert.deepEqual(errors,[]);
 console.log('PASS: native Chromium IME composition from first jamo, stable focus after arrows/Tab/Enter/click, replacement, F2 caret editing, mixed text, cancel and reload');
}catch(e){await page.screenshot({path:'test-results/ime-entry-failure.png'});throw e;}finally{await session.detach();await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
