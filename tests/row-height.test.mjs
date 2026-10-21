import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {createWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {pastePlan} from '../src/cell-paste.mjs';

const near=(actual,expected)=>assert.ok(Math.abs(actual-expected)<1e-7,`${actual} != ${expected}`);
const clip=heights=>({r:0,c:0,data:heights.map((_,r)=>[{raw:`한글 ${r+1}행`,style:{fontSize:18}}]),rowHeights:heights});

test('XLSX default row height survives import and export alongside explicit and hidden rows',async()=>{
 const native=new ExcelJS.Workbook(),s=native.addWorksheet('행 높이');s.properties.defaultRowHeight=32;
 for(let r=1;r<=4;r++){s.getCell(`A${r}`).value=`한글 ${r}행`;s.getCell(`A${r}`).font={size:18};}
 s.getRow(2).height=40;s.getRow(3).hidden=true;s.getRow(4).height=60;s.getRow(4).hidden=true;
 const imported=await readSpreadsheet(await native.xlsx.writeBuffer(),'default-height.xlsx'),sheet=imported.sheets[0];
 near(sheet.defaultRowHeight,32*4/3);near(sheet.rowHeights[1],40*4/3);assert.equal(sheet.rowHeights[2],0);assert.equal(sheet.rowHeights[3],0);
 near(sheet.hiddenSizes.rows[2],32*4/3);near(sheet.hiddenSizes.rows[3],60*4/3);
 assert.equal(sheet.rowHeights[0],undefined);assert.equal(sheet.cells.A1.style.fontSize,18);
 const output=await writeSpreadsheet(imported),again=(await readSpreadsheet(output.data,'again.xlsx')).sheets[0];
 near(again.defaultRowHeight,32*4/3);near(again.rowHeights[1],40*4/3);assert.equal(again.rowHeights[2],0);
 near(again.hiddenSizes.rows[2],32*4/3);near(again.hiddenSizes.rows[3],60*4/3);
});

test('XLSX used-range metadata never compresses explicit unequal heights, merged cells or empty height-only rows',async()=>{
 const native=new ExcelJS.Workbook(),s=native.addWorksheet('명시 높이'),heights=[24,40,28,36];
 heights.forEach((h,r)=>{s.getRow(r+1).height=h;s.getCell(`A${r+1}`).value=`한글 ${r+1}행`;});
 s.mergeCells('A1:A2');s.getRow(7).height=52;
 const zip=await JSZip.loadAsync(await native.xlsx.writeBuffer()),part='xl/worksheets/sheet1.xml';
 zip.file(part,(await zip.file(part).async('string')).replace(/<dimension\s+ref="[^"]+"\s*\/>/,'<dimension ref="A1:A1"/>'));
 const book=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'small-dimension.xlsx');
 for(const current of [book,(await readSpreadsheet((await writeSpreadsheet(book)).data,'again.xlsx'))]){
  heights.forEach((h,r)=>near(current.sheets[0].rowHeights[r],h*4/3));
  near(current.sheets[0].rowHeights[6],52*4/3);assert.deepEqual(current.sheets[0].merges,['A1:A2']);
 }
});

test('external table row heights follow visible destination rows and repeated ranges without shrinking existing rows',()=>{
 const sheet=createWorkbook().sheets[0];sheet.rowHeights={3:80,4:0};
 const p=pastePlan(sheet,clip([32*4/3,40*4/3]),{r:3,c:1},{selection:{r1:3,c1:1,r2:6,c2:1},hiddenRow:r=>r===4});
 assert.deepEqual(p.rowHeightUpdates,[{row:5,height:40*4/3},{row:6,height:32*4/3},{row:7,height:40*4/3}]);
 assert.deepEqual(sheet.rowHeights,{3:80,4:0});assert.equal(sheet.cells.B4,undefined);
 const tall=createWorkbook().sheets[0];tall.defaultRowHeight=90;
 assert.deepEqual(pastePlan(tall,clip([42]),{r:0,c:0}).rowHeightUpdates,[]);
});

test('row height import leaves values-only, formula-only, transposed, arithmetic and skipped blank pastes unchanged',()=>{
 const sheet=createWorkbook().sheets[0],source=clip([60]);source.data[0][0].raw='2';source.values=[[2]];
 for(const options of [{mode:'values'},{mode:'formulas'},{mode:'formats'},{transpose:true},{operation:'add'}])assert.deepEqual(pastePlan(sheet,source,{r:0,c:0},options).rowHeightUpdates,[]);
 assert.deepEqual(pastePlan(sheet,{...source,data:[[{raw:''}]]},{r:0,c:0},{skipBlanks:true}).rowHeightUpdates,[]);
 for(const height of [NaN,Infinity,-1,0,10001,'60'])assert.deepEqual(pastePlan(sheet,clip([height]),{r:0,c:0}).rowHeightUpdates,[]);
 assert.deepEqual(pastePlan(sheet,{r:0,c:0,data:[[{raw:'plain text'}]]},{r:0,c:0}).rowHeightUpdates,[]);
});
