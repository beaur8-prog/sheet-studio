import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {resetCellFormat} from '../src/cell-format-reset.mjs';
import {formatCellRecord} from '../src/cell-paste.mjs';
import {captureClipboardStyles,adoptClipboardStyles} from '../src/cell-clipboard.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {worksheetCellStyle} from '../src/cell-style.mjs';
import {styleDefaultsFixture} from './fixtures/style-defaults.mjs';
test('format reset overrides row/column inheritance in existing and unallocated cells without changing content or notes',async()=>{
 const book=await readSpreadsheet(await styleDefaultsFixture(),'reset.xlsx'),s=book.sheets[0];s.cells.B3=resetCellFormat({raw:'12',note:'keep'});s.cells.B6=resetCellFormat();s.cells.C5=resetCellFormat({raw:'.25'});assert.equal(worksheetCellStyle(book,s,2,1).color,'#004488');assert.equal(worksheetCellStyle(book,s,4,2).fontFamily,'Arial');assert.equal(s.cells.B3.raw,'12');assert.equal(s.cells.B3.note,'keep');
 const out=await writeSpreadsheet(book),excel=new ExcelJS.Workbook();await excel.xlsx.load(out.data);const t=excel.worksheets[0];assert.equal(t.getCell('B3').font.size,14);assert.equal(t.getCell('B3').border.diagonal,undefined);assert.equal(t.getCell('C5').alignment.textRotation,undefined);assert.equal(t.getCell('B3').numFmt,undefined);const round=await readSpreadsheet(out.data,'reset-again.xlsx');assert.equal(worksheetCellStyle(round,round.sheets[0],5,1).color,'#004488');
});
test('reset and formats-only application remove character overrides while preserving values, hyperlinks, array definitions and protected lock overrides',()=>{
 const old={raw:"'001",note:'keep',xlsxProtection:{locked:false},xlsxValue:{raw:"'001",value:{richText:[{text:'001',font:{bold:true,color:{argb:'FFFF0000'}}}]}},xlsxFormula:{shareType:'array',ref:'A1:A2'},cachedValue:7};const reset=resetCellFormat(old,{protection:{locked:false}});assert.deepEqual(reset.xlsxValue.value.richText,[{text:'001'}]);assert.deepEqual(reset.xlsxProtection,{locked:false});assert.equal(old.xlsxValue.value.richText[0].font.bold,true);assert.deepEqual(reset.xlsxFormula,old.xlsxFormula);
 const painted=formatCellRecord(old,{style:{bold:false},xlsxStyle:4});assert.equal(painted.raw,"'001");assert.equal(painted.xlsxStyle,4);assert.equal(painted.note,'keep');assert.equal(painted.xlsxProtection,undefined);assert.deepEqual(painted.xlsxValue.value.richText,[{text:'001'}]);assert.equal(painted.cachedValue,7);
 const link=resetCellFormat({raw:'link',xlsxValue:{raw:'link',value:{text:'link',hyperlink:'https://example.com'}}});assert.equal(link.xlsxValue.value.hyperlink,'https://example.com');
});
test('format painter portable definitions include explicit protection and source advanced formatting',async()=>{
 const b=await readSpreadsheet(await styleDefaultsFixture(),'paint.xlsx'),s=b.sheets[0];s.cells.B3={raw:'12',xlsxProtection:{locked:false}};s.cells.D2={raw:'destination',note:'keep'};const data=[[structuredClone(s.cells.B3)]],xlsxStyles=captureClipboardStyles(b,s,data,[2],[1]),updates=[{key:'D2',cell:formatCellRecord(s.cells.D2,data[0][0])}],added=adoptClipboardStyles(b,{xlsxStyles},updates,{mode:'formats'});b.xlsxStyles.push(...added);s.cells.D2=updates[0].cell;const excel=new ExcelJS.Workbook();await excel.xlsx.load((await writeSpreadsheet(b)).data);const cell=excel.worksheets[0].getCell('D2');assert.equal(cell.value,'destination');assert.equal(cell.border.diagonal.up,true);assert.equal(cell.protection.locked,false);assert.equal(cell.numFmt,'00000');
});
