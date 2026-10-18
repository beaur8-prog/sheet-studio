import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import path from 'node:path';
import platform from '../scripts/platform.cjs';
import {createRequire} from 'node:module';
import {createWorkbook} from '../src/engine.mjs';
import fs from 'node:fs/promises';
const root=path.resolve(import.meta.dirname,'..');
const packaged=process.argv.includes('--packaged');
const version=createRequire(import.meta.url)('../package.json').version;
const env={...process.env,SHEET_STUDIO_DATA:path.join(root,'.package','sheet-lifecycle-'+Date.now())};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch({executablePath:process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env});
const page=await app.firstWindow();
page.setDefaultTimeout(8000);const findings=[];
const settle=()=>page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const load=async b=>{await page.locator('#file-input').setInputFiles({name:'review.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(n=>sheetStudio.getWorkbook().name===n,b.name);await settle();};
const rename=async(name)=>{await page.locator('[data-sheet].active').click({button:'right'});await page.locator('[data-action="rename-sheet"]').click();await page.locator('#prompt-value').fill(name);await page.getByRole('dialog').getByRole('button',{name:'확인',exact:true}).click();await settle();};
try{
await page.locator('.cell').first().waitFor();
let b=createWorkbook();b.name='delete reference';b.sheets[0].cells.A1={raw:'=Source!A1'};let s=structuredClone(createWorkbook().sheets[0]);s.id='source';s.name='Source';s.cells.A1={raw:'42'};b.sheets.push(s);await load(b);
const before=await page.evaluate(()=>sheetStudio.evaluate('A1'));
await page.locator('[data-sheet="source"]').click({button:'right'});await page.locator('[data-action="delete-sheet"]').click();await page.getByRole('dialog').getByRole('button',{name:'삭제',exact:true}).click();await settle();
const afterDelete=await page.evaluate(()=>({raw:sheetStudio.getWorkbook().sheets[0].cells.A1.raw,value:sheetStudio.evaluate('A1')}));
await page.locator('#undo').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets.length),2);
await page.locator('#redo').click();await settle();assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets.length),1);
await page.getByRole('button',{name:'새 시트',exact:true}).click();await settle();await rename('Source');await page.locator('#name-box').fill('A1');await page.locator('#name-box').press('Enter');await page.keyboard.press('F2');await page.locator('#editor').fill('999');await page.locator('#editor').press('Enter');await page.locator('[data-sheet="sheet-1"]').click();await settle();
findings.push({case:'deleted sheet reference revives',before,afterDelete,afterRecreate:await page.evaluate(()=>sheetStudio.evaluate('A1'))});
b=createWorkbook();b.name='rename conditional reference';b.sheets[0].cells.A1={raw:'1'};b.sheets[0].conditionalRules=[{id:'cf',range:'A1:A1',type:'expression',formulae:['Sheet1!A1=1'],style:{background:'#ff0000'}}];await load(b);
const oldColor=await page.locator('[role="gridcell"][aria-label="A1"]').evaluate(e=>getComputedStyle(e).backgroundColor);await rename('Renamed');
findings.push({case:'rename conditional formula',before:oldColor,after:await page.locator('[role="gridcell"][aria-label="A1"]').evaluate(e=>getComputedStyle(e).backgroundColor),rules:await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].conditionalRules)});
const attempts=[];for(let i=0;i<3;i++){await page.locator('[data-sheet].active').dblclick();await settle();attempts.push(await page.locator('#prompt-value').count());if(attempts.at(-1))await page.keyboard.press('Escape');}findings.push({case:'sheet tab double-click rename',dialogCounts:attempts});
assert.equal(before,42);assert.equal(afterDelete.raw,'=#REF!');assert.equal(afterDelete.value,'#REF!');
assert.equal(findings[0].afterRecreate,'#REF!');assert.equal(findings[1].after,oldColor);assert.equal(findings[1].rules[0].formulae[0],"'Renamed'!A1=1");assert.deepEqual(attempts,[1,1,1]);
// An inactive tab must retain its identity during the first click too.
await page.getByRole('button',{name:'새 시트',exact:true}).click();await settle();await page.locator('[data-sheet="sheet-1"]').dblclick();await page.locator('#prompt-value').waitFor();await page.keyboard.press('Escape');
await page.reload();await page.locator('.cell').first().waitFor();await settle();assert.equal(await page.locator('[role="gridcell"][aria-label="A1"]').evaluate(e=>getComputedStyle(e).backgroundColor),oldColor);
await fs.mkdir('test-results/review',{recursive:true});
await fs.writeFile('test-results/review/boundaries-fixed.json',JSON.stringify(findings,null,2));console.log(JSON.stringify(findings,null,2));await page.screenshot({path:'test-results/review/conditional-rename-fixed.png'});
}finally{await app.close();}
