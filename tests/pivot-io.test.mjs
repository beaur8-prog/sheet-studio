import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {fixture,spec} from './fixtures/pivot.mjs';
import {applyPivot} from '../src/pivot.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';

test('XLSX includes native pivot, caches, typed records and durable editable settings',async()=>{
  const b=fixture(),s=spec(),target={id:'pivot',name:'피벗',charts:[]};applyPivot(b,target,s);b.sheets.push(target);b.activeSheetId=target.id;
  b.sheets[0].cells.C2.raw='40'; // Export must not combine old displayed results with a new cache.
  const out=await writeSpreadsheet(b),zip=await JSZip.loadAsync(out.data);
  assert.ok(zip.file('xl/pivotTables/pivotTable1.xml'));
  assert.match(await zip.file('xl/pivotCache/pivotCacheRecords1.xml').async('string'),/<n v="40"\/>/);
  assert.match(await zip.file('xl/_rels/workbook.xml.rels').async('string'),/customXml/);
  const read=await readSpreadsheet(out.data,'피벗.xlsx');assert.equal(read.sheets[1].cells.D4.raw,'190');assert.equal(target.cells.D4.raw,'160');
  assert.equal(read.sheets[1].pivot.sourceSheetId,read.sheets[0].id);assert.deepEqual(read.sheets[1].pivot.values,s.values);
  // Excel renames custom parts. App settings must still be discoverable.
  zip.file('customXml/item7.xml',await zip.file('customXml/sheetStudioPivots.xml').async('string'));zip.remove('customXml/sheetStudioPivots.xml');
  const renamed=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'renamed.xlsx');assert.ok(renamed.sheets[1].pivot);
});
test('report filters reserve their Excel rows; empty reports retain zero results without an incorrect native filter',async()=>{
  const b=fixture(),s={...spec(),rows:['상품'],columns:[],filters:[{field:'지역',values:['동부']}]},target={id:'pivot',name:'피벗',charts:[]};applyPivot(b,target,s);b.sheets.push(target);
  const out=await writeSpreadsheet(b),zip=await JSZip.loadAsync(out.data);
  assert.match(await zip.file('xl/pivotTables/pivotTable1.xml').async('string'),/ref="A3:B6"/);
  const read=await readSpreadsheet(out.data,'filter.xlsx');assert.equal(read.sheets[1].cells.B4.raw,'60');
  target.pivot.filters[0].values=[];const empty=await writeSpreadsheet(b),ez=await JSZip.loadAsync(empty.data);assert.equal(ez.file('xl/pivotTables/pivotTable1.xml'),null);
  const reread=await readSpreadsheet(empty.data,'empty.xlsx');assert.equal(reread.sheets[1].cells.B2.raw,'0');assert.deepEqual(reread.sheets[1].pivot.filters[0].values,[]);
});

test('report filter placement keeps dependent formulas correct across XLSX roundtrip',async()=>{
  const b=fixture(),s={...spec(),rows:['상품'],columns:[],filters:[{field:'지역',values:['동부']}]},target={id:'pivot',name:'피벗',charts:[]};applyPivot(b,target,s);b.sheets.push(target);
  b.sheets[0].cells.F1={raw:"='피벗'!B4",style:{}};
  const out=await writeSpreadsheet(b),excel=new ExcelJS.Workbook();await excel.xlsx.load(out.data);
  assert.equal(excel.getWorksheet('원본').getCell('F1').formula,"'피벗'!B6");assert.equal(excel.getWorksheet('원본').getCell('F1').result,60);
  const read=await readSpreadsheet(out.data,'links.xlsx');assert.equal(read.sheets[0].cells.F1.raw,"='피벗'!B4");assert.equal(read.sheets[1].cells.B4.raw,'60');
});
