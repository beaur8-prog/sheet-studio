import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {createWorkbook,createEvaluator,formulaForXlsx,formulaForEditing,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,child,children,all,attr,stringify,readRelationships,resolvePart} from '../src/chart-xml.mjs';
import {valueMetadataFixture} from './fixtures/value-metadata.mjs';
import {structuralWorkbookPlan} from '../src/workbook-structure.mjs';

const fixture=()=>fs.readFile(new URL('./fixtures/xlsxwriter-dynamic.xlsx',import.meta.url));
async function cellXML(bytes,ref){const zip=await JSZip.loadAsync(bytes),doc=xml(await zip.file('xl/worksheets/sheet1.xml').async('string'));return all(doc,'c').find(c=>attr(c,'r')===ref);}

test('independent XlsxWriter mixed-type arrays ignore poisoned caches and retain boolean, empty text and element errors',async()=>{
 let b=await readSpreadsheet(await fs.readFile(new URL('./fixtures/xlsxwriter-array-controls.xlsx',import.meta.url)),'typed-arrays.xlsx');
 for(let i=0;i<3;i++){
  const s=b.sheets[0],e=createEvaluator(b,s.id);assert.equal(s.cells.A5.raw,"'");assert.equal(s.cells.A6?.raw||'','');
  assert.deepEqual(Array.from({length:6},(_,r)=>e('F'+(r+1))),[true,false,'#N/A',2,'',0]);
  assert.deepEqual(Array.from({length:6},(_,r)=>e('L'+(r+1))),[true,true,false,false,false,false]);
  assert.equal(e.result('F1').error,null);assert.equal(e.result('F3').error,'#N/A');assert.equal(e('N1'),3);
  b=await readSpreadsheet((await writeSpreadsheet(b)).data,'typed-arrays.xlsx');
 }
});

test('independently generated XlsxWriter arrays retain dynamic metadata and recalculate stale followers separately from legacy CSE',async()=>{
  let b=await readSpreadsheet(await fixture(),'independent.xlsx');
  for(let i=0;i<3;i++){
    b=parseWorkbook(serializeWorkbook(b));const s=b.sheets[0],e=createEvaluator(b,s.id);assert.equal(s.cells.A1.xlsxFormula.dynamic,true);assert.equal(s.cells.F1.xlsxFormula.dynamic,true);assert.equal(s.cells.H1.xlsxFormula.dynamic,undefined);assert.equal(s.cells.A2?.raw||'','');assert.equal(e('A3'),5);assert.equal(e('B3'),6);assert.equal(e('D1'),21);assert.equal(e('F3'),3);assert.equal(e('H3'),10);
    const bytes=(await writeSpreadsheet(b)).data,zip=await JSZip.loadAsync(bytes),metadata=xml(await zip.file('xl/metadata.xml').async('string'));
    assert.equal(children(child(metadata.documentElement,'metadataTypes'),'metadataType').filter(n=>attr(n,'name')==='XLDAPR').length,1);assert.equal(all(metadata,'dynamicArrayProperties').length,1);assert.equal(attr(all(metadata,'dynamicArrayProperties')[0],'fDynamic'),'1');assert.equal(attr(await cellXML(bytes,'A1'),'cm'),'1');assert.equal(attr(await cellXML(bytes,'H1'),'cm'),'');
    assert.match(child(await cellXML(bytes,'D1'),'f').textContent,/_xlfn\.ANCHORARRAY\(A1\)/);assert.match(child(await cellXML(bytes,'F1'),'f').textContent,/_xlfn\._xlws\.SORT\(_xlfn\.UNIQUE/);
    b=await readSpreadsheet(bytes,'again.xlsx');
  }
});
test('new arrays preserve text/error/boolean types and resize or clear cached followers after reopening',async()=>{
  let b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'=FILTER({"#N/A";#N/A;TRUE;"001"},{1;1;1;1})'},C1:{raw:'=SEQUENCE(4)'},E1:{raw:'=SUM(C1#)'}};
  let bytes=(await writeSpreadsheet(b)).data;const excel=new ExcelJS.Workbook();await excel.xlsx.load(bytes);const source=excel.worksheets[0];assert.equal(source.getCell('A1').result,'#N/A');assert.deepEqual(source.getCell('A2').value,{error:'#N/A'});assert.equal(source.getCell('A3').value,true);assert.equal(source.getCell('A4').value,'001');
  b=await readSpreadsheet(bytes,'types.xlsx');s=b.sheets[0];let e=createEvaluator(b,s.id);assert.deepEqual(e.result('A1'),{value:'#N/A',error:null});assert.equal(e.result('A2').error,'#N/A');assert.equal(e('A3'),true);assert.equal(e('A4'),'001');
  s.cells.C1.raw='=SEQUENCE(2)';bytes=(await writeSpreadsheet(b)).data;assert.equal(attr(child(await cellXML(bytes,'C1'),'f'),'ref'),'C1:C2');assert.equal(child(await cellXML(bytes,'C3'),'v'),undefined);
  b=await readSpreadsheet(bytes,'shrink.xlsx');s=b.sheets[0];e=createEvaluator(b,s.id);assert.equal(e('C2'),2);assert.equal(e('C3'),'');assert.equal(e('E1'),3);
  delete s.cells.C1;delete s.cells.E1;bytes=(await writeSpreadsheet(b)).data;b=await readSpreadsheet(bytes,'clear.xlsx');e=createEvaluator(b,b.sheets[0].id);assert.equal(e('C2'),'');assert.equal(e.spills().filter(a=>a.ref==='C1').length,0);
});
test('blocked spills save a single-cell error with dynamic metadata and CSV includes live spill followers',async()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'=SEQUENCE(3)'},A2:{raw:'block'},C1:{raw:'=SEQUENCE(3,2)'}};
  const bytes=(await writeSpreadsheet(b)).data,a=await cellXML(bytes,'A1');assert.equal(attr(child(a,'f'),'ref'),'A1');assert.equal(attr(a,'t'),'e');assert.equal(child(a,'v').textContent,'#SPILL!');assert.notEqual(attr(a,'cm'),'');
  delete s.cells.A2;delete s.cells.C1;const csv=new TextDecoder().decode((await writeSpreadsheet(b,'csv')).data).replace(/^\uFEFF/,'');assert.equal(csv,'1\r\n2\r\n3');
});
test('unknown dynamic functions keep their labelled original cached range even when caches are marked stale',async()=>{
  const zip=await JSZip.loadAsync(await fixture()),doc=xml(await zip.file('xl/worksheets/sheet1.xml').async('string')),anchor=all(doc,'c').find(c=>attr(c,'r')==='A1');child(anchor,'f').textContent='_xlfn.UNKNOWN_DYNAMIC(3,2)';zip.file('xl/worksheets/sheet1.xml',stringify(doc));
  let b=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'future.xlsx'),s=b.sheets[0],e=createEvaluator(b,s.id);assert.equal(e('A1'),'#NAME?');assert.equal(e('B3'),6);assert.equal(e.result('B3').cached,true);assert.equal(e.result('B3').stale,false);assert.equal(e.spillAt('B3').ref,'A1');
  b.referenceCachesValid=false;e=createEvaluator(b,s.id);assert.equal(e('B3'),'#NAME?');assert.equal(e.result('B3').stale,true);assert.equal(e('D1'),'#NAME?');
  const bytes=(await writeSpreadsheet(b)).data;assert.equal(attr(child(await cellXML(bytes,'A1'),'f'),'ref'),'A1:B3');assert.equal(child(await cellXML(bytes,'A3'),'v').textContent,'5');assert.equal(child(await cellXML(bytes,'B3'),'v').textContent,'6');assert.match(child(await cellXML(bytes,'A1'),'f').textContent,/_xlfn\.UNKNOWN_DYNAMIC/);
  const reopened=await readSpreadsheet(bytes,'future-again.xlsx');assert.equal(reopened.referenceCachesValid,false);assert.equal(createEvaluator(reopened,reopened.activeSheetId)('B3'),'#NAME?');
});
test('unrelated metadata records, original cm/vm indices and linked rich-data parts survive adding and resizing spills',async()=>{
  const zip=await JSZip.loadAsync(await fixture()),metadata=xml(await zip.file('xl/metadata.xml').async('string')),ns=metadata.documentElement.namespaceURI,types=child(metadata.documentElement,'metadataTypes'),custom=metadata.createElementNS(ns,'metadataType');custom.setAttribute('name','APP_PRIVATE');types.insertBefore(custom,types.firstChild);types.setAttribute('count','2');for(const rc of all(metadata,'rc'))rc.setAttribute('t','2');
  const cells=child(metadata.documentElement,'cellMetadata'),bk=metadata.createElementNS(ns,'bk'),rc=metadata.createElementNS(ns,'rc');rc.setAttribute('t','1');rc.setAttribute('v','0');bk.appendChild(rc);cells.insertBefore(bk,cells.firstChild);cells.setAttribute('count','2');zip.file('xl/metadata.xml',stringify(metadata));
  const sheet=xml(await zip.file('xl/worksheets/sheet1.xml').async('string'));for(const c of all(sheet,'c'))if(attr(c,'cm'))c.setAttribute('cm','2');const j=sheet.createElementNS(ns,'c');j.setAttribute('r','J1');j.setAttribute('cm','1');j.setAttribute('vm','1');j.setAttribute('t','n');const v=sheet.createElementNS(ns,'v');v.textContent='42';j.appendChild(v);all(sheet,'row')[0].appendChild(j);zip.file('xl/worksheets/sheet1.xml',stringify(sheet));
  const rels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string')),rel=rels.createElementNS(rels.documentElement.namespaceURI,'Relationship');rel.setAttribute('Id','rIdRichPrivate');rel.setAttribute('Type','http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue');rel.setAttribute('Target','richData/private.xml');rels.documentElement.appendChild(rel);zip.file('xl/_rels/workbook.xml.rels',stringify(rels));zip.file('xl/richData/private.xml','<private xmlns="urn:test">preserved opaque metadata</private>');
  let bytes=await zip.generateAsync({type:'uint8array'});
  for(let i=0;i<3;i++){
    const b=await readSpreadsheet(bytes,'mixed.xlsx'),s=b.sheets[0];s.cells.L1={raw:'=SEQUENCE(2)'};bytes=(await writeSpreadsheet(b)).data;const out=await JSZip.loadAsync(bytes),meta=xml(await out.file('xl/metadata.xml').async('string'));
    assert.equal(children(child(meta.documentElement,'metadataTypes'),'metadataType').length,2);assert.equal(children(child(meta.documentElement,'cellMetadata'),'bk').length,2);assert.equal(attr(await cellXML(bytes,'J1'),'cm'),'1');assert.equal(attr(await cellXML(bytes,'J1'),'vm'),'1');assert.equal(attr(await cellXML(bytes,'L1'),'cm'),'2');
    const rich=(await readRelationships(out,'xl/workbook.xml')).find(r=>r.type.endsWith('/rdRichValue'));assert.match(await out.file(resolvePart('xl/workbook.xml',rich.target)).async('string'),/preserved opaque metadata/);
  }
});
test('spill and implicit-intersection syntax round-trips with canonical future prefixes without touching string literals',()=>{
  for(const f of ['=SUM(A1#)','=FILTER(A1#,A1#>2)','=TEXT(@A1:A3,"0")','=IF(A1,@(B1#),@D1:D4)','=SUM("A1#",\'한글 시트\'!$A$1#)']){
    const exported=formulaForXlsx(f),editing=formulaForEditing(exported);assert.equal(formulaForXlsx(editing),exported);assert.equal(formulaForXlsx(exported),exported);
  }
  assert.equal(formulaForXlsx('=FILTER(A1#,A1#>2)'),'=_xlfn._xlws.FILTER(_xlfn.ANCHORARRAY(A1),_xlfn.ANCHORARRAY(A1)>2)');
});

test('imported supported spills release cache-only cells while retaining explicit styles and notes',async()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'=SEQUENCE(10000)'},A2:{raw:'',style:{background:'#abcdef'},note:'Keep this note'}};
 const loaded=await readSpreadsheet((await writeSpreadsheet(b)).data,'compact.xlsx'),target=loaded.sheets[0];assert.equal(Object.keys(target.cells).length,2);assert.equal(target.cells.A2.style.background.toUpperCase(),'#ABCDEF');assert.equal(target.cells.A2.note,'Keep this note');assert.equal(target.cells.A2.cachedValue,undefined);assert.equal(createEvaluator(loaded,loaded.activeSheetId)('A10000'),10000);
 target.cells.A1.raw='=SEQUENCE(2)';const again=await readSpreadsheet((await writeSpreadsheet(loaded)).data,'small.xlsx');assert.equal(Object.keys(again.sheets[0].cells).length,2);assert.equal(createEvaluator(again,again.activeSheetId)('A3'),'');
});

test('unchanged dynamic followers retain their original value-metadata pointers instead of losing them during cache compaction',async()=>{
 let bytes=await valueMetadataFixture();
 for(let round=0;round<3;round++){const b=await readSpreadsheet(bytes,'metadata.xlsx'),s=b.sheets[0];assert.equal(s.cells.A2.raw,'');assert.equal(s.cells.A2.xlsxMetadata.vm,1);bytes=(await writeSpreadsheet(parseWorkbook(serializeWorkbook(b)))).data;for(const ref of ['A1','A2','D2','G2'])assert.equal(attr(await cellXML(bytes,ref),'vm'),'1',ref);}
});

test('value metadata follows the original typed result, never a changed or removed spill value',async()=>{
 const b=await readSpreadsheet(await valueMetadataFixture(),'metadata.xlsx'),s=b.sheets[0];s.cells.A1.raw='=SEQUENCE(2)';s.cells.D1.raw='=SEQUENCE(3,1,10)';delete s.cells.G1;b.referenceCachesValid=false;const bytes=(await writeSpreadsheet(b)).data;
 for(const ref of ['A1','A2','D2','G2'])assert.equal(attr(await cellXML(bytes,ref),'vm'),'',ref);assert.equal(child(await cellXML(bytes,'D2'),'v').textContent,'11');assert.equal(s.cells.G2.note,'metadata note');
 const same=await readSpreadsheet(await valueMetadataFixture(),'same.xlsx');same.sheets[0].cells.Z1={raw:'unrelated edit'};same.referenceCachesValid=false;const unchanged=(await writeSpreadsheet(same)).data;for(const ref of ['A1','A2','D2','G2'])assert.equal(attr(await cellXML(unchanged,ref),'vm'),'1',ref);
 const future=await readSpreadsheet(await valueMetadataFixture({unknown:true}),'future.xlsx');future.referenceCachesValid=false;const preserved=(await writeSpreadsheet(future)).data;for(const ref of ['A1','A2'])assert.equal(attr(await cellXML(preserved,ref),'vm'),'1');assert.equal(createEvaluator(future,future.activeSheetId).result('A2').stale,true);
 const broken=parseWorkbook(serializeWorkbook(same));broken.sheets[0].cells.A2.xlsxMetadata.array='A0';assert.throws(()=>parseWorkbook(JSON.stringify(broken)),/메타데이터/);
});

test('structural movement rebases metadata owners and formula/cache signatures together',async()=>{
 let b=await readSpreadsheet(await valueMetadataFixture(),'metadata.xlsx'),id=b.activeSheetId;b=structuralWorkbookPlan(b,id,{axis:'r',index:0,count:1,delta:1});assert.equal(b.sheets[0].cells.A3.xlsxMetadata.array,'A2');assert.equal(createEvaluator(b,id)('A3'),2);let bytes=(await writeSpreadsheet(b)).data;assert.equal(attr(await cellXML(bytes,'A2'),'vm'),'1');assert.equal(attr(await cellXML(bytes,'A3'),'vm'),'1');
 b=structuralWorkbookPlan(b,id,{axis:'c',index:0,count:1,delta:1});assert.equal(b.sheets[0].cells.B3.xlsxMetadata.array,'B2');bytes=(await writeSpreadsheet(b)).data;assert.equal(attr(await cellXML(bytes,'B3'),'vm'),'1');
 const future=await readSpreadsheet(await valueMetadataFixture({unknown:true}),'future.xlsx');future.sheets[0].cells.A1.raw='=_xlfn.FUTUREARRAY(A2,3)';future.sheets[0].cells.A1.cachedFormula=future.sheets[0].cells.A1.raw;future.sheets[0].cells.A1.xlsxFormula.raw=future.sheets[0].cells.A1.raw;future.sheets[0].cells.A1.xlsxFormula.formula='_xlfn.FUTUREARRAY(A2,3)';future.sheets[0].cells.A1.xlsxMetadata.raw=future.sheets[0].cells.A1.raw;
 const moved=structuralWorkbookPlan(future,future.activeSheetId,{axis:'r',index:0,count:1,delta:1});assert.equal(moved.sheets[0].cells.A2.cachedFormula,'=_xlfn.FUTUREARRAY(A3,3)');assert.equal(moved.sheets[0].cells.A3.xlsxMetadata.array,'A2');bytes=(await writeSpreadsheet(moved)).data;assert.equal(attr(await cellXML(bytes,'A3'),'vm'),'1');
});
