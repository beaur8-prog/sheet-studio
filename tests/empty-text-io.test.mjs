import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {createWorkbook,createEvaluator} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';

test('explicit empty text remains distinct from a physical blank through repeated XLSX saves',async()=>{
 let b=createWorkbook();b.sheets[0].cells={A1:{raw:"'"},A2:{raw:'',style:{background:'#abcdef'}},A3:{raw:'0'},B1:{raw:'=SORT(A1:A3)'},D1:{raw:'=UNIQUE(A1:A3)'},F1:{raw:'=ISBLANK(A1)'},F2:{raw:'=ISTEXT(A1)'},F3:{raw:'=COUNTA(A1:A3)'},F4:{raw:'=ISBLANK(A2)'},F5:{raw:'=ISTEXT(A2)'}};
 const before=createEvaluator(b,b.activeSheetId),expected=Object.fromEntries(['B1','B2','B3','D1','D2','D3','F1','F2','F3','F4','F5'].map(ref=>[ref,before(ref)]));
 assert.equal(expected.F1,false);assert.equal(expected.F2,true);assert.equal(expected.F3,2);assert.equal(expected.F4,true);assert.equal(expected.F5,false);
 for(let i=0;i<3;i++){b=await readSpreadsheet((await writeSpreadsheet(b)).data,'empty-text.xlsx');const s=b.sheets[0],e=createEvaluator(b,s.id);assert.equal(s.cells.A1.raw,"'");assert.equal(s.cells.A2.raw,'');for(const [ref,value] of Object.entries(expected))assert.equal(e(ref),value,ref+' round '+i);}
});

test('shared-string, inline-string and rich-text empty values import as text while untyped styled blanks remain blank',async()=>{
 const source=new ExcelJS.Workbook(),s=source.addWorksheet('Input');s.getCell('A1').value='';s.getCell('B1').value='placeholder';s.getCell('C1').value={richText:[{font:{bold:true},text:'placeholder'}]};s.getCell('D1').fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFABCDEF'}};
 const z=await JSZip.loadAsync(await source.xlsx.writeBuffer()),part='xl/worksheets/sheet1.xml';let xml=await z.file(part).async('string');xml=xml.replace(/<c\b[^>]*r="B1"[^>]*>[\s\S]*?<\/c>/,'<c r="B1" t="inlineStr"><is><t></t></is></c>').replace(/<c\b[^>]*r="C1"[^>]*>[\s\S]*?<\/c>/,'<c r="C1" t="inlineStr"><is><r><rPr><b/></rPr><t></t></r></is></c>');z.file(part,xml);
 const b=await readSpreadsheet(await z.generateAsync({type:'uint8array'}),'independent-empty.xlsx');for(const ref of ['A1','B1','C1'])assert.equal(b.sheets[0].cells[ref]?.raw,"'",ref);assert.equal(b.sheets[0].cells.D1.raw,'');
});

test('explicit empty formula caches survive repeated saves and are distinguished from a missing imported cache',async()=>{
 const source=new ExcelJS.Workbook(),s=source.addWorksheet('Input');s.getCell('A1').value={formula:'_xlfn.FUTURE_EMPTY()',result:''};s.getCell('B1').value={formula:'_xlfn.FUTURE_EMPTY()'};
 let b=await readSpreadsheet(await source.xlsx.writeBuffer(),'cached-empty.xlsx');
 assert.equal(b.sheets[0].cells.B1.cachedValue,undefined);
 for(let i=0;i<3;i++){const sheet=b.sheets[0],e=createEvaluator(b,sheet.id);assert.equal(sheet.cells.A1.cachedValue,'');assert.equal(e.result('A1').cachedValue,'');assert.equal(e.result('A1').cached,true);b=await readSpreadsheet((await writeSpreadsheet(b)).data,'cached-empty.xlsx');}
});
