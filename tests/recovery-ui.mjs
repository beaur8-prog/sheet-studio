import assert from 'node:assert/strict';

export async function assertBlankStart(page) {
  const state=await page.evaluate(()=>{const book=sheetStudio.getWorkbook();return {name:book.name,sheets:book.sheets.length,cells:book.sheets.reduce((n,s)=>n+Object.keys(s.cells).length,0),charts:book.sheets.reduce((n,s)=>n+(s.charts?.length||0),0)};});
  assert.deepEqual(state,{name:'통합 문서1',sheets:1,cells:0,charts:0});
}
export async function recoverPrevious(page) {
  await page.locator('[data-tab="파일"]').click();
  await page.locator('.menu [data-action="recover"]').click();
  await page.getByRole('button',{name:'복구',exact:true}).click();
  await page.locator('.workbook-busy').waitFor({state:'detached'});
  await page.evaluate(()=>sheetStudio.whenIdle());
}
