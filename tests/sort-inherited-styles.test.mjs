import test from 'node:test';
import assert from 'node:assert/strict';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import ExcelJS from 'exceljs';
import {createWorkbook,validateWorkbook} from '../src/engine.mjs';
import {sortRange,sortFilterRange} from '../src/filters.mjs';
import {worksheetCellStyle} from '../src/cell-style.mjs';
import {styleDefaultsFixture} from './fixtures/style-defaults.mjs';
test('sorting inherited row styles has the same result before or after XLSX materializes their cell formats',async()=>{
 const original=await readSpreadsheet(await styleDefaultsFixture(),'sort-defaults.xlsx'),s=original.sheets[0];for(let r=2;r<=5;r++)s.cells['C'+r]={raw:String(7-r)};
 const saved=await readSpreadsheet((await writeSpreadsheet(original)).data,'saved.xlsx');const before=structuredClone(original),a=before.sheets[0],b=saved.sheets[0],range={r1:1,r2:4,c1:2,c2:2};sortRange(before,a,range,2,false);sortRange(saved,b,range,2,false);assert.equal(a.cells.C2.raw,'2');assert.equal(b.cells.C2.raw,'2');assert.deepEqual(worksheetCellStyle(before,a,1,2),worksheetCellStyle(saved,b,1,2));
});

test('blank cells, advanced row/column formatting and Normal overrides travel with sorted records',async()=>{
 const book=await readSpreadsheet(await styleDefaultsFixture(),'blank.xlsx'),s=book.sheets[0];
 // The first Normal record moves into the styled row before that row moves out.
 for(let row=2;row<=5;row++)s.cells['C'+row]={raw:String(row===2?1:7-row)};
 s.cells.D5={raw:'row',note:'keep'};s.cells.B5={raw:'25'};
 sortRange(book,s,{r1:1,r2:4,c1:1,c2:3},2,false);
 assert.equal(s.cells.C5.raw,'4');assert.equal(worksheetCellStyle(book,s,4,2).fontFamily,'Arial');
 assert.equal(s.cells.D3.raw,'row');assert.equal(s.cells.D3.note,'keep');assert.equal(worksheetCellStyle(book,s,2,3).fontFamily,'Georgia');
 assert.equal(worksheetCellStyle(book,s,4,3).fontFamily,'Arial');assert.equal(s.cells.D5.xlsxNormal,true);
 const bytes=(await writeSpreadsheet(book)).data,out=new ExcelJS.Workbook();await out.xlsx.load(bytes);
 assert.equal(out.worksheets[0].getCell('D3').alignment.textRotation,30);assert.equal(out.worksheets[0].getCell('B5').border.diagonal.up,true);assert.equal(out.worksheets[0].getCell('B3').font.name,'Georgia');assert.equal(out.worksheets[0].getCell('B5').fill.fgColor.argb,'FFFFFFCC');assert.equal(out.worksheets[0].getCell('D5').font.name,'Arial');
 const loaded=await readSpreadsheet(bytes,'sorted.xlsx');assert.equal(worksheetCellStyle(loaded,loaded.sheets[0],4,3).fontFamily,'Arial');validateWorkbook(loaded);
});
test('filtered sorting preserves inherited formats and excluded rows while table banding stays on destination rows',async()=>{
 const book=await readSpreadsheet(await styleDefaultsFixture(),'filtered.xlsx'),s=book.sheets[0];for(let r=2;r<=5;r++)s.cells['C'+r]={raw:String(7-r)};
 s.autoFilter={range:{r1:0,r2:4,c1:2,c2:3},columns:{},hiddenRows:[2]};const excluded=structuredClone(s.cells.C3);
 s.tables=[{id:'t',name:'Banding',range:'C1:D5',headerRow:true,totalsRow:false,columns:[{id:1,name:'Value'},{id:2,name:'Other'}],style:{name:'TableStyleMedium2',showRowStripes:true}}];
 const background=worksheetCellStyle(book,s,1,2).background;sortFilterRange(book,s,2,false);
 assert.equal(s.cells.C2.raw,'2');assert.deepEqual(s.cells.C3,excluded);assert.equal(worksheetCellStyle(book,s,1,2).fontFamily,'Georgia');assert.equal(worksheetCellStyle(book,s,1,2).background,background);assert.equal(s.cells.C2.style?.background,undefined);
 const saved=await readSpreadsheet((await writeSpreadsheet(book)).data,'filtered-sorted.xlsx');assert.equal(saved.sheets[0].cells.C2.raw,'2');assert.equal(worksheetCellStyle(saved,saved.sheets[0],1,2).fontFamily,'Georgia');
});
test('sort style snapshots reuse definitions and unstyled sorting adds no metadata',()=>{
 const book=createWorkbook(),s=book.sheets[0],style={style:{font:{name:'Georgia',size:16,color:{argb:'FF0000FF'}}},baseline:{fontFamily:'Georgia',fontSize:16,color:'#0000FF'}};book.xlsxStyles=[style];s.xlsxStyleDefaults={rows:{},columns:{}};
 for(let row=0;row<10000;row++){s.cells['A'+(row+1)]={raw:String(10000-row)};if(row>=5000)s.xlsxStyleDefaults.rows[row]=0;}
 sortRange(book,s,{r1:0,r2:9999,c1:0,c2:0},0,false);assert.equal(s.cells.A1.raw,'1');assert.equal(s.cells.A10000.xlsxNormal,true);assert.equal(book.xlsxStyles.length,2);assert.equal(Object.keys(s.cells.A1).length,3);
 const plain=createWorkbook(),sheet=plain.sheets[0];sheet.cells={A1:{raw:'2'},A2:{raw:'1'}};sortRange(plain,sheet,{r1:0,r2:1,c1:0,c2:0},0,false);assert.equal(plain.xlsxStyles,undefined);assert.deepEqual(sheet.cells.A1,{raw:'1'});
});
test('native XLSM row-template sort is rejected before any records or format definitions change',async()=>{
 const book=await readSpreadsheet(await styleDefaultsFixture(),'native.xlsx'),s=book.sheets[0];book.nativeSource={format:'xlsm'};s.cells.C2={raw:'3'};s.cells.C5={raw:'1'};const before=JSON.stringify(book);assert.throws(()=>sortRange(book,s,{r1:1,r2:4,c1:2,c2:2},2,false),/XLSM/);assert.equal(JSON.stringify(book),before);
});
