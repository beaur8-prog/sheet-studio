import {chromium} from 'playwright';
import assert from 'node:assert/strict';
const browser=await chromium.launch({headless:true,channel:'msedge'});
const context=await browser.newContext({viewport:{width:1540,height:950},permissions:['clipboard-read','clipboard-write']});
const page=await context.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));
const go=async a=>{await page.locator('#name-box').fill(a);await page.locator('#name-box').press('Enter');await page.waitForTimeout(70)};
const enter=async(a,v)=>{await go(a);await page.keyboard.press('F2');await page.locator('#editor').fill(v);await page.locator('#editor').press('Enter');await page.waitForTimeout(70)};
const val=async a=>page.evaluate(a=>sheetStudio.evaluate(a),a);
try{
await page.goto('http://127.0.0.1:5192');await page.locator('.cell').first().waitFor();
await enter('A1','7');await enter('B1','=A1*3');await go('B1');await page.getByRole('button',{name:'복사',exact:true}).click();await go('B2');await page.getByRole('button',{name:'붙여넣기',exact:true}).click();await page.waitForTimeout(150);assert.equal(await val('B2'),0);await enter('A2','4');assert.equal(await val('B2'),12);
await go('D1:E1');await page.getByRole('button',{name:'병합하고 가운데 맞춤',exact:true}).click();await page.waitForTimeout(70);assert.deepEqual(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].merges),['D1:E1']);
await go('A3');await page.keyboard.press('F2');await page.locator('#editor').fill('draft');await page.getByRole('button',{name:'새 시트',exact:true}).click();assert.equal(await val('A1'),'');assert.equal(await page.evaluate(()=>sheetStudio.getWorkbook().sheets[0].cells.A3.raw),'draft');
await page.getByRole('button',{name:'도움말',exact:true}).click();await page.getByRole('button',{name:'예제 통합 문서',exact:true}).click();await go('A5:F10');await page.getByRole('button',{name:'데이터',exact:true}).click();await page.getByRole('button',{name:'내림차순',exact:true}).click();assert.equal(await val('A5'),'6월');assert.equal(await val('D5'),8800000);
await go('A1');await page.getByRole('button',{name:'텍스트 필터',exact:true}).click();await page.getByRole('dialog').locator('input').fill('3월');await page.getByRole('button',{name:'확인',exact:true}).click();await page.waitForTimeout(100);assert.equal(await page.locator('#status').innerText(),'필터 모드');assert.equal(await page.getByRole('gridcell',{name:'A8',exact:true}).innerText(),'3월');await page.getByRole('button',{name:'필터 해제',exact:true}).click();
await page.locator('.search-command').click();await page.locator('#find-text').fill('신제품');await page.getByRole('button',{name:'다음 찾기',exact:true}).click();assert.match(await page.locator('#find-status').innerText(),/1개 발견/);await page.getByRole('button',{name:'닫기',exact:true}).click();
const downloaded=page.waitForEvent('download');await page.getByRole('button',{name:'파일',exact:true}).click();await page.getByRole('button',{name:'앱 문서 저장 (.sheet.json)',exact:true}).click();const d=await downloaded;const path=await d.path();assert(path);await page.locator('#file-input').setInputFiles({name:'roundtrip.sheet.json',mimeType:'application/json',buffer:await (await import('node:fs/promises')).readFile(path)});await page.waitForTimeout(100);assert.equal(await val('D12'),41300000);
assert.deepEqual(errors,[]);console.log('UI features: clipboard relative refs, merge, pending draft sheet switch, sort formulas, filtering, find, JSON roundtrip passed.');
}finally{await browser.close();}
