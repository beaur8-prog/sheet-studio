import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';
import {createWorkbook} from '../src/engine.mjs';

const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged');
const version=createRequire(import.meta.url)('../package.json').version;
await fs.mkdir(path.join(root,'.package'),{recursive:true});
const dataDir=await fs.mkdtemp(path.join(root,'.package/entry-navigation-'));
const env={...process.env,SHEET_STUDIO_DATA:dataDir};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow(),errors=[];page.setDefaultTimeout(10000);page.on('pageerror',e=>errors.push(e.message));
const editor=page.locator('#editor'),name=page.locator('#name-box');
const go=async ref=>{await name.fill(ref);await name.press('Enter');};
const raw=ref=>page.evaluate(ref=>sheetStudio.getWorkbook().sheets[0].cells[ref]?.raw,ref);
const entered=async(ref,text,key,target)=>{
 await go(ref);await page.keyboard.type(text);await editor.press(key);
 assert.equal(await raw(ref),text,`${key} must commit the entry`);
 assert.equal(await editor.evaluate(el=>getComputedStyle(el).opacity),'0',`${key} must leave entry mode`);
 assert.equal(await editor.inputValue(),'');
 assert.equal(await name.inputValue(),target,`${key} must select ${target}`);
 assert.equal(await page.evaluate(()=>document.activeElement.id),'editor');
};
try{
 await page.locator('.cell').first().waitFor();const book=createWorkbook();book.name='Entry navigation regression';
 const sheet=book.sheets[0];sheet.cells={B20:{raw:'original'},C20:{raw:'caret text'},D20:{raw:'=1+2'}};
 sheet.rowHeights={5:0};sheet.columnWidths={5:0};sheet.merges=['H10:I11'];
 await page.locator('#file-input').setInputFiles({name:'entry.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(book))});
 await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Entry navigation regression');
 for(const [key,target] of [['ArrowRight','D3'],['ArrowLeft','B3'],['ArrowUp','C2'],['ArrowDown','C4']])await entered('C3','hello',key,target);
 await entered('B20','123','ArrowRight','C20');await page.keyboard.press('ControlOrMeta+z');assert.equal(await raw('B20'),'original');
 await entered('E5','hidden column','ArrowRight','G5');await entered('E5','hidden row','ArrowDown','E7');
 await entered('H10','merged right','ArrowRight','J10');await entered('H10','merged down','ArrowDown','H12');
 await entered('A1','left edge','ArrowLeft','A1');await entered('A1','top edge','ArrowUp','A1');
 await entered('B22','=1+2','ArrowDown','B23');assert.equal(await page.evaluate(()=>sheetStudio.evaluate('B22')),3);
 // A formula awaiting a reference still points to cells instead of committing.
 await go('D22');await page.keyboard.type('=');await editor.press('ArrowLeft');assert.equal(await editor.inputValue(),'=C22');assert.equal(await raw('D22'),undefined);await editor.press('Escape');
 // F2 and double-click are explicit text editing, including after another input event.
 await go('C20');await page.keyboard.press('F2');await editor.fill('changed text');await editor.press('End');await editor.press('ArrowLeft');
 assert.equal(await editor.evaluate(el=>el.selectionStart),11);assert.equal(await editor.isVisible(),true);assert.equal(await raw('C20'),'caret text');await editor.press('Escape');
 await go('C20');await page.getByRole('gridcell',{name:'C20',exact:true}).dblclick();await editor.fill('double click');await editor.press('End');await editor.press('ArrowLeft');
 assert.equal(await editor.evaluate(el=>el.selectionStart),11);assert.equal(await raw('C20'),'caret text');await editor.press('Escape');
 await go('D20');await page.keyboard.press('F2');await editor.press('ArrowLeft');assert.equal(await editor.inputValue(),'=1+2');assert.equal(await editor.evaluate(el=>el.selectionStart),3);await editor.press('Escape');
 await go('C20');await page.locator('#formula').fill('formula bar text');await page.locator('#formula').press('ArrowLeft');assert.equal(await raw('C20'),'caret text');await page.locator('#formula').press('Escape');
 await go('B24');await page.keyboard.type('abc');await editor.press('Shift+ArrowLeft');assert.equal(await editor.isVisible(),true);assert.equal(await raw('B24'),undefined);await editor.press('Escape');
 // Composition flags must prevent an IME navigation key from committing early.
 await go('B25');await page.keyboard.type('a');await editor.dispatchEvent('compositionstart');await editor.fill('한글 입력');
 await editor.dispatchEvent('keydown',{key:'ArrowRight',bubbles:true});assert.equal(await raw('B25'),undefined);assert.equal(await editor.isVisible(),true);
 await editor.dispatchEvent('compositionend');await editor.press('ArrowRight');assert.equal(await raw('B25'),'한글 입력');assert.equal(await name.inputValue(),'C25');
 assert.deepEqual(errors,[]);await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/entry-navigation.png')});
 console.log('PASS: direct input arrow commit/navigation, numbers, formulas/references, hidden/merged cells, edges, undo, F2/double-click/formula-bar caret editing and IME guards');
}finally{await app.close();await fs.rm(dataDir,{recursive:true,force:true});}
