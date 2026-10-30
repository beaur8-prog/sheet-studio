import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {createWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {captureClipboardStyles,adoptClipboardStyles} from '../src/cell-clipboard.mjs';
import {pastePlan} from '../src/cell-paste.mjs';
import {createColorResolver} from '../src/spreadsheet-colors.mjs';
import {clipboardColorFixture} from './fixtures/clipboard-colors.mjs';
test('cross-document copies resolve source theme/tint/indexed colors including advanced fills/borders and rich-text runs',async()=>{
 const source=await readSpreadsheet(await clipboardColorFixture(),'color-source.xlsx'),sheet=source.sheets[0],data=[[structuredClone(sheet.cells.A1)],[structuredClone(sheet.cells.A2)]],before=JSON.stringify(source),styles=captureClipboardStyles(source,sheet,data,[0,1],[0]),resolve=createColorResolver(source.chartTheme,'<styleSheet>'+source.xlsxColors+'</styleSheet>'),color=v=>'FF'+resolve(v).slice(1).toUpperCase();
 const dest=createWorkbook(),s=dest.sheets[0],clip={data,xlsxStyles:styles,r:0,c:0};const plan=pastePlan(s,clip,{r:0,c:0}),added=adoptClipboardStyles(dest,clip,plan.updates);dest.xlsxStyles=added;for(const u of plan.updates)s.cells[u.key]=u.cell;
 const excel=new ExcelJS.Workbook();await excel.xlsx.load((await writeSpreadsheet(dest)).data);const check=excel.worksheets[0];assert.equal(check.getCell('A1').font.color.argb,color({theme:4,tint:.2}));assert.equal(check.getCell('A1').border.diagonal.color.argb,'FF0088AA');assert.equal(check.getCell('A1').fill.stops[0].color.argb,'FFCC2200');assert.equal(check.getCell('A1').fill.stops[1].color.argb,'FF0088AA');assert.equal(check.getCell('A2').value.richText[0].font.color.argb,'FFCC2200');assert.equal(check.getCell('A2').value.richText[1].font.color.argb,'FF0088AA');assert.equal(JSON.stringify(source),before);
});
