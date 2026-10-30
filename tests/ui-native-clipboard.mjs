import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createWorkbook} from '../src/engine.mjs';

// This intentionally replaces the host clipboard. Run only on a disposable CI
// host, never as a local verification command against a user's pasteboard.
if(process.env.GITHUB_ACTIONS!=='true')throw Error('Native clipboard verification requires a disposable GitHub Actions host; the local clipboard is untouched.');
const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','native-clipboard-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:profile};delete env.ELECTRON_RUN_AS_NODE;
let app,page,failed=false;const checks=[];
async function step(name,operation){console.log('native-clipboard: '+name);const start=Date.now();let timer;try{const result=await Promise.race([operation(),new Promise((_,reject)=>{timer=setTimeout(()=>reject(Error('Native clipboard phase timed out: '+name)),30000);})]);checks.push({name,ms:Date.now()-start,status:'passed'});return result;}catch(error){checks.push({name,ms:Date.now()-start,status:'failed',error:error.message});throw error;}finally{clearTimeout(timer);}}
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
try{
 await step('launch isolated app',async()=>{app=await _electron.launch(desktopLaunch(root,env));page=await app.firstWindow();page.setDefaultTimeout(15000);await page.locator('.cell').first().waitFor();});
 await step('prepare native clipboard source',async()=>{const b=createWorkbook();b.name='Native clipboard control';b.sheets[0].cells={A1:{raw:'한글 표',style:{bold:true,background:'#fff2cc'}},A2:{raw:'=1+2'},B2:{raw:'줄1\n줄2'}};b.sheets[0].merges=['A1:B1'];await page.locator('#file-input').setInputFiles({name:'native-clipboard.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Native clipboard control');await settle();await page.locator('#name-box').fill('A1:B2');await page.locator('#name-box').press('Enter');await settle();});
 await step('real trusted IPC writes OS HTML and TSV',async()=>{await page.locator('[data-action=copy]').click();await page.getByText('선택한 셀을 복사했습니다.',{exact:true}).waitFor();});
 await step('read back the app-owned native payload',async()=>{const actual=await app.evaluate(({clipboard})=>({text:clipboard.readText(),html:clipboard.readHTML()}));assert.ok(actual.text.includes('한글 표'));assert.ok(actual.text.includes('"줄1\n줄2"'));assert.ok(actual.html.includes('colspan="2"'));assert.ok(!actual.html.includes('=1+2'));});
 await step('real OS paste preserves own formula and merged cells',async()=>{await page.locator('#name-box').fill('D1');await page.locator('#name-box').press('Enter');await settle();await page.locator('[data-action=paste]').click();await page.waitForFunction(()=>sheetStudio.getWorkbook().sheets[0].cells.D2?.raw==='=1+2');await settle();const s=await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0]);assert.equal(s.cells.E2.raw,'줄1\n줄2');assert.ok(s.merges.includes('D1:E1'));});
 await step('native external HTML and TSV paste',async()=>{await app.evaluate(({clipboard})=>clipboard.write({text:'외부 표\t42',html:'<table><tr style="height:32pt"><td>외부 표</td><td>42</td></tr></table>'}));await page.locator('#name-box').fill('G1');await page.locator('#name-box').press('Enter');await settle();await page.locator('[data-action=paste]').click();await page.waitForFunction(()=>sheetStudio.getWorkbook().sheets[0].cells.G1?.raw==='외부 표');await settle();const s=await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0]);assert.equal(s.cells.H1.raw,'42');assert.ok(Math.abs(s.rowHeights[0]-32*4/3)<1e-6);});
 await step('normal process close',async()=>{await app.close();app=null;});
 console.log('PASS: disposable-host native OS clipboard HTML/TSV, trusted desktop IPC, formula/merged-cell paste and external table height');
}catch(error){failed=true;console.error(error.stack||error.message);if(app)app.process().kill('SIGKILL');}
finally{await fs.mkdir(path.join(root,'test-results'),{recursive:true});await fs.writeFile(path.join(root,'test-results/native-clipboard.json'),JSON.stringify({at:new Date().toISOString(),platform:process.platform,disposableCiHost:true,checks},null,2));await fs.rm(profile,{recursive:true,force:true});if(failed)process.exit(1);}
