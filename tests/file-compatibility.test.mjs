import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {readCompatibilityRisks,compatibilityRisks,compatibilitySummary,validateCompatibilityReview} from '../src/file-compatibility.mjs';

async function base(){const b=new ExcelJS.Workbook();b.addWorksheet('문서').getCell('A1').value='원본 유지';return JSZip.loadAsync(await b.xlsx.writeBuffer());}

test('plain XLSX and app pivot Custom XML do not trigger unrelated feature warnings',async()=>{
  const zip=await base();assert.equal(await readCompatibilityRisks(zip,'plain.xlsx'),undefined);
  zip.file('customXml/item7.xml','<pivots xmlns="urn:sheet-studio:pivots:1">[]</pivots>');
  assert.equal(await readCompatibilityRisks(zip,'plain.xlsx'),undefined);
});

test('known unsupported OLE/control/comment/model/slicer data is detected before destructive XLSX export',async()=>{
  const zip=await base();zip.file('xl/embeddings/original.bin',new Uint8Array([0,255,13]));zip.file('xl/activeX/activeX1.xml','<control/>');zip.file('xl/threadedComments/threadedComment1.xml','<ThreadedComments/>');zip.file('xl/model/item.data',new Uint8Array([9]));zip.file('xl/slicerCaches/slicerCache1.xml','<slicerCache/>');zip.file('customXml/item1.xml','<businessData>keep original</businessData>');
  const bytes=await zip.generateAsync({type:'uint8array'}),book=await readSpreadsheet(bytes,'Original.xlsx');
  assert.deepEqual(new Set(book.xlsxCompatibility.risks.map(r=>r.kind)),new Set(['ole','activex','threaded','model','slicer','custom']));
  assert.equal(compatibilityRisks(book).length,6);assert.equal(compatibilitySummary(book)[0].mode,'편집 가능');
  await assert.rejects(writeSpreadsheet(book),/파일 호환성/);
  const copy=await writeSpreadsheet(book,'xlsx',{allowFeatureLoss:true}),saved=await JSZip.loadAsync(copy.data);assert.equal(saved.file('xl/embeddings/original.bin'),null);
  assert.deepEqual(book.xlsxCompatibility,parseWorkbook(serializeWorkbook(book)).xlsxCompatibility);
  assert.equal((await readSpreadsheet(copy.data,'copy.xlsx')).xlsxCompatibility,undefined);
  assert.ok(zip.file('xl/embeddings/original.bin'),'in-memory source remains unchanged');
});

test('content types recognize relocated objects and VBA in an incorrectly named XLSX',async()=>{
  const zip=await base();zip.file('different/object.bin',new Uint8Array([1]));zip.file('xl/vbaProject.bin',new Uint8Array([2]));
  const path='[Content_Types].xml';zip.file(path,(await zip.file(path).async('string')).replace('</Types>','<Override PartName="/different/object.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>'));
  assert.deepEqual((await readCompatibilityRisks(zip,'bad-name.xlsx')).risks.map(r=>r.kind),['ole','macro']);
  assert.deepEqual((await readCompatibilityRisks(zip,'macro.xlsm')).risks.map(r=>r.kind),['ole']);
});

test('review metadata rejects unknown/duplicate kinds, counts and injected source paths',()=>{
  const value={version:1,sourceName:'문서.xlsx',risks:[{kind:'ole',count:1}]};validateCompatibilityReview(value);
  for(const bad of [{...value,sourceName:'\0bad'}, {...value,risks:[{kind:'invented',count:1}]},{...value,risks:[{kind:'ole',count:0}]},{...value,risks:[...value.risks,...value.risks]}])assert.throws(()=>validateCompatibilityReview(bad),/호환성/);
});
