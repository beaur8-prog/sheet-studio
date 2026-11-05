import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {fixture,spec} from './fixtures/pivot.mjs';
import {applyPivot} from '../src/pivot.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {nativePivotIndex,preservedPivots,removePreservedPivots,validateNativePivotArchive} from '../src/native-pivots.mjs';
import {cellEditError} from '../src/cell-constraints.mjs';
import {structuralWorkbookPlan} from '../src/workbook-structure.mjs';

async function sample(){const book=fixture(),target={id:'pivot',name:'피벗',charts:[]};applyPivot(book,target,spec());book.sheets.push(target);return {book,zip:await JSZip.loadAsync((await writeSpreadsheet(book)).data)};}
const bytes=zip=>zip.generateAsync({type:'uint8array'});
async function changeResult(zip,value=40){const path='xl/worksheets/sheet2.xml',before=await zip.file(path).async('string');assert.match(before,/<c r="D4"/);zip.file(path,before.replace(/(<c r="D4"[^>]*>[\s\S]*?<v>)160(<\/v>)/,(_,a,b)=>a+value+b));}

test('external aggregate change keeps saved results, formula positions and native pivot through three roundtrips',async()=>{
  const {zip}=await sample();
  zip.file('xl/pivotTables/pivotTable1.xml',(await zip.file('xl/pivotTables/pivotTable1.xml').async('string')).replace('subtotal="sum"','subtotal="average"'));
  await changeResult(zip);
  let book=await readSpreadsheet(await bytes(zip),'external.xlsx');
  assert.equal(book.sheets[1].pivot,undefined);assert.equal(book.sheets[1].cells.D4.raw,'40');assert.match(book.sheets[1].pivotImportNotice,/외부/);
  assert.match(cellEditError(book,book.sheets[1],'D4'),/원본 피벗/);assert.equal(cellEditError(book,book.sheets[0],'C2'),null);
  assert.throws(()=>structuralWorkbookPlan(book,book.sheets[0].id,{axis:'r',index:1,count:1,delta:1}),/원본 피벗/);
  book.sheets[0].cells.C2.raw='999';
  for(let i=0;i<3;i++){
    book=parseWorkbook(serializeWorkbook(book));
    const out=await writeSpreadsheet(book),saved=await JSZip.loadAsync(out.data),entries=await nativePivotIndex(saved,book);
    assert.equal(entries.length,1);assert.match(await saved.file(entries[0].path).async('string'),/subtotal="average"/);
    assert.match(await saved.file(entries[0].cache).async('string'),/worksheetSource[^>]*ref="A1:D6"/);
    const records=Object.keys(saved.files).filter(p=>/pivotCacheRecords1\.xml$/.test(p));assert.equal(records.length,1);
    book=await readSpreadsheet(out.data,'external.xlsx');assert.equal(book.sheets[1].cells.D4.raw,'40');assert.equal(book.sheets[0].cells.C2.raw,'999');assert.equal(preservedPivots(book,book.sheets[1].id).length,1);
    assert.ok(!Object.keys(saved.files).some(p=>/sheetStudioPivotArchive\/xl\/sheetStudioPivotArchive/.test(p)));
  }
});

test('legacy app metadata cannot overwrite an external native field/filter change',async()=>{
  const {zip}=await sample(),meta='customXml/sheetStudioPivots.xml';
  // Rebuild metadata explicitly to model files created before layout proofs.
  const {book}=await sample();zip.file(meta,'<pivots xmlns="urn:sheet-studio:pivots:1">'+JSON.stringify([{sheet:'피벗',source:'원본',spec:book.sheets[1].pivot,rowOffset:0}]).replace(/&/g,'&amp;').replace(/</g,'&lt;')+'</pivots>');
  const original=await bytes(zip),unchanged=await readSpreadsheet(original,'legacy.xlsx');assert.ok(unchanged.sheets[1].pivot);assert.equal(unchanged.xlsxPivotArchive,undefined);
  zip.file('xl/pivotTables/pivotTable1.xml',(await zip.file('xl/pivotTables/pivotTable1.xml').async('string')).replace('subtotal="sum"','subtotal="average"'));
  await changeResult(zip);
  const changed=await readSpreadsheet(await bytes(zip),'legacy-edited.xlsx');assert.equal(changed.sheets[1].pivot,undefined);assert.equal(changed.sheets[1].cells.D4.raw,'40');assert.ok(changed.xlsxPivotArchive);
});

test('native pivots without app metadata retain fields/cache and custom number formats',async()=>{
  const {zip}=await sample();zip.remove('customXml/sheetStudioPivots.xml');
  const table='xl/pivotTables/pivotTable1.xml';zip.file(table,(await zip.file(table).async('string')).replace('subtotal="sum"','subtotal="sum" numFmtId="201"'));
  const styles='xl/styles.xml';let s=await zip.file(styles).async('string');
  if(s.includes('<numFmts'))s=s.replace(/<numFmts count="(\d+)"/,(_,n)=>'<numFmts count="'+(+n+1)+'"').replace('</numFmts>','<numFmt numFmtId="201" formatCode="0.0000 &quot;units&quot;"/></numFmts>');
  else s=s.replace(/(<styleSheet\b[^>]*>)/,'$1<numFmts count="1"><numFmt numFmtId="201" formatCode="0.0000 &quot;units&quot;"/></numFmts>');
  zip.file(styles,s);
  const book=await readSpreadsheet(await bytes(zip),'native.xlsx');assert.ok(book.xlsxPivotArchive);assert.equal(book.sheets[1].pivot,undefined);
  const out=await writeSpreadsheet(book),saved=await JSZip.loadAsync(out.data),[entry]=await nativePivotIndex(saved,book);
  const {all,attr,xml}=await import('../src/chart-xml.mjs'),field=all(entry.table,'dataField')[0],format=attr(field,'numFmtId');
  assert.ok(all(xml(await saved.file(styles).async('string')),'numFmt').some(n=>attr(n,'numFmtId')===format&&attr(n,'formatCode')==='0.0000 "units"'));
  const changed=structuredClone(book);changed.sheets[0].name='changed';await assert.rejects(writeSpreadsheet(changed),/시트 연결/);
  removePreservedPivots(book,book.sheets[1].id);assert.equal(book.xlsxPivotArchive,undefined);assert.equal(cellEditError(book,book.sheets[1],'D4'),null);
  const plain=await writeSpreadsheet(book);assert.equal((await nativePivotIndex(await JSZip.loadAsync(plain.data),book)).length,0);
});

test('pivot archive validates identities, package paths and format payloads',async()=>{
  const {zip}=await sample();zip.remove('customXml/sheetStudioPivots.xml');const book=await readSpreadsheet(await bytes(zip),'native.xlsx');
  validateNativePivotArchive(book.xlsxPivotArchive);
  const invalid=structuredClone(book.xlsxPivotArchive);invalid.sheets[0].path='missing';assert.throws(()=>validateNativePivotArchive(invalid),/연결/);
  const bad=structuredClone(book.xlsxPivotArchive);bad.parts['../escape.xml']={data:'AA==',contentType:'application/xml'};assert.throws(()=>validateNativePivotArchive(bad),/연결 파일/);
});

test('advanced native pivot dependencies block lossy XLSX saving while app documents retain the archive',async()=>{
  const {zip}=await sample();zip.remove('customXml/sheetStudioPivots.xml');
  zip.file('xl/pivotCache/pivotCacheDefinition1.xml',(await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string')).replace('type="worksheet"','type="external"'));
  const book=await readSpreadsheet(await bytes(zip),'external-source.xlsx');assert.match(book.xlsxPivotArchive.blockReason,/외부 데이터/);
  const recovered=parseWorkbook(serializeWorkbook(book));assert.deepEqual(recovered.xlsxPivotArchive,book.xlsxPivotArchive);
  await assert.rejects(writeSpreadsheet(book),/원본 피벗의 XLSX/);
  removePreservedPivots(book,book.sheets[1].id);await writeSpreadsheet(book);
});
