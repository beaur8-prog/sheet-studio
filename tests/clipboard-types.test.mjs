import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {createWorkbook,createEvaluator,validateWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {entryRaw} from '../src/cell-values.mjs';
import {pastePlan} from '../src/cell-paste.mjs';
import {captureClipboardStyles,adoptClipboardStyles} from '../src/cell-clipboard.mjs';
import {worksheetCellStyle} from '../src/cell-style.mjs';
import {styleDefaultsFixture} from './fixtures/style-defaults.mjs';
test('Text entry preserves numbers, percent, booleans, errors and incomplete formula strings without changing existing numeric types',()=>{
 const b=createWorkbook(),s=b.sheets[0],texts=['0012','1e4','10%','TRUE','#N/A','=SUM(','12345678901234567890',"'001",'ordinary'];
 texts.forEach((text,i)=>s.cells['A'+(i+1)]={raw:entryRaw(text,{format:'text'})});const e=createEvaluator(b,s.id);
 texts.forEach((text,i)=>assert.equal(e('A'+(i+1)),text.startsWith("'")?text.slice(1):text));
 assert.equal(entryRaw('=1+2',{format:'general'}),'=1+2');assert.equal(entryRaw('001',{numberFormat:'@'}),"'001");s.cells.B1={raw:'12',style:{format:'text'}};assert.equal(createEvaluator(b,s.id)('B1'),12);
});
test('plain text paste retains destination styles and notes, Text entry is distinct from numeric values-only paste',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'old',note:'keep',style:{format:'text',bold:true,background:'#ffff00'},xlsxStyle:0};b.xlsxStyles=[{style:{alignment:{textRotation:30}},baseline:{}}];
 const plain=pastePlan(s,{r:0,c:0,plainText:true,data:[[{raw:'0012'}]]},{r:0,c:0});assert.equal(plain.updates[0].cell.raw,"'0012");assert.equal(plain.updates[0].cell.style.bold,true);assert.equal(plain.updates[0].cell.note,'keep');assert.equal(plain.updates[0].cell.xlsxStyle,0);
 const values=pastePlan(s,{r:0,c:0,data:[[{raw:'12'}]],values:[[12]]},{r:0,c:0},{mode:'values'});assert.equal(values.updates[0].cell.raw,'12');assert.equal(values.updates[0].cell.style.format,'text');
 const inherited=pastePlan(s,{r:0,c:0,plainText:true,data:[[{raw:'=A1'}]]},{r:1,c:0},{styleAt:()=>({format:'text'})});assert.equal(inherited.updates[0].cell.raw,"'=A1");
});
test('clipboard owns portable Normal, axis, table and advanced style definitions across workbooks and XLSX reopen',async()=>{
 const source=await readSpreadsheet(await styleDefaultsFixture(),'source.xlsx'),s=source.sheets[0];s.cells.B3={raw:'12'};s.cells.C5={raw:'.25'};
 const data=[[structuredClone(s.cells.B3),{raw:'normal'}],[{raw:'9'},structuredClone(s.cells.C5)]],rows=[2,4],columns=[1,2],styles=captureClipboardStyles(source,s,data,rows,columns),clip={r:2,c:1,sourceRows:rows,sourceCols:columns,data,xlsxStyles:styles};
 const target=createWorkbook(),t=target.sheets[0];target.xlsxStyles=[{style:{font:{name:'Consolas',color:{argb:'FF00FF00'}}},baseline:{fontFamily:'Consolas',color:'#00ff00'}}];target.xlsxDefaultStyle=0;
 const plan=pastePlan(t,clip,{r:0,c:0}),added=adoptClipboardStyles(target,clip,plan.updates);assert.equal(target.xlsxStyles.length,1);target.xlsxStyles.push(...added);for(const u of plan.updates)t.cells[u.key]=u.cell;
 validateWorkbook(target);assert.equal(worksheetCellStyle(target,t,0,0).color.toLowerCase(),'#ff0000');assert.equal(worksheetCellStyle(target,t,1,1).fontFamily,'Georgia');assert.equal(worksheetCellStyle(target,t,0,1).color,'#004488');
 const out=await writeSpreadsheet(target),excel=new ExcelJS.Workbook();await excel.xlsx.load(out.data);const check=excel.worksheets[0];assert.equal(check.getCell('A1').numFmt,'00000');assert.equal(check.getCell('A1').border.diagonal.up,true);assert.equal(check.getCell('B2').alignment.textRotation,30);assert.equal(check.getCell('B2').alignment.shrinkToFit,true);assert.equal(check.getCell('B2').protection.locked,false);assert.equal(check.getCell('B1').font.size,14);
 const round=await readSpreadsheet(out.data,'round.xlsx');assert.equal(worksheetCellStyle(round,round.sheets[0],0,1).italic,true);
 const second=pastePlan(t,clip,{r:5,c:0});assert.equal(adoptClipboardStyles(target,clip,second.updates).length,0);
});
test('format-only paste adopts source styles while formula/value-only paste keeps target style indices',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells.C1={raw:'99',style:{color:'#008800'},xlsxStyle:0};b.xlsxStyles=[{style:{font:{color:{argb:'FF008800'}}},baseline:{color:'#008800'}}];
 const clip={r:0,c:0,data:[[{raw:'=A2',style:{bold:true},xlsxStyle:0}]],values:[[12]],xlsxStyles:[{style:{font:{bold:true}},baseline:{bold:true}}]};
 const f=pastePlan(s,clip,{r:0,c:2},{mode:'formats'});assert.equal(f.updates[0].cell.raw,'99');assert.equal(adoptClipboardStyles(b,clip,f.updates,{mode:'formats'}).length,1);assert.equal(f.updates[0].cell.xlsxStyle,1);
 for(const mode of ['values','formulas']){const p=pastePlan(s,clip,{r:0,c:2},{mode});assert.deepEqual(adoptClipboardStyles(b,clip,p.updates,{mode}),[]);assert.equal(p.updates[0].cell.xlsxStyle,0);assert.equal(p.updates[0].cell.style.color,'#008800');}
});
