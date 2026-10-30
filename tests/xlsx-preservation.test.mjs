import {preservationFixture as fixture} from './xlsx-preservation-fixture.mjs';
import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {validateWorkbook,serializeWorkbook,parseWorkbook,formulaForXlsx} from '../src/engine.mjs';
import {xml,all,attr,child} from '../src/chart-xml.mjs';

test('XLSX roundtrips advanced styles, rich text, hyperlinks, validation, protection, arrays and chartless names',async()=>{
  let bytes=await fixture();let original;
  for(let round=0;round<3;round++){
    let book=validateWorkbook(await readSpreadsheet(bytes,'preservation.xlsx'));book=parseWorkbook(serializeWorkbook(book));
    const s=book.sheets[0];assert.ok(s.xlsxFeatures.dataValidations.includes('예,아니오'));assert.ok(s.xlsxFeatures.sheetProtection.includes('hashValue'));
    assert.equal(book.chartNames.find(n=>n.name==='Revenue').formula.replace(/'/g,''),'Details!$A$3');
    bytes=(await writeSpreadsheet(book)).data;const check=new ExcelJS.Workbook();await check.xlsx.load(bytes);const out=check.worksheets[0];
    assert.deepEqual(out.getCell('A1').alignment,{horizontal:'left',shrinkToFit:true,indent:2,textRotation:45});assert.equal(out.getCell('A1').font.strike,true);
    assert.deepEqual(out.getCell('A1').value.richText.map(run=>[run.text,run.font.bold||false,run.font.italic||false]),[['중요 ',true,false],['메모',false,true]]);
    assert.equal(out.getCell('A2').value.hyperlink,'https://example.org/manual?q=a&x=b');assert.equal(s.cells.A2.xlsxValue.value.tooltip,'원본 도움말');const zip=await JSZip.loadAsync(bytes);assert.equal(attr(all(xml(await zip.file('xl/worksheets/sheet1.xml').async('string')),'hyperlink')[0],'tooltip'),'원본 도움말');
    assert.equal(out.getCell('A3').fill.type,'gradient');assert.equal(out.getCell('A3').fill.degree,30);assert.equal(out.getCell('A3').border.diagonal.style,'double');assert.equal(out.getCell('A3').border.diagonal.up,true);
    assert.deepEqual(out.getCell('A4').protection,{locked:false,hidden:true});assert.equal(out.getCell('D1').dataValidation.error,'목록에서 선택하세요.');
    assert.equal(out.getCell('B1').value.shareType,'array');assert.equal(out.getCell('B1').value.ref,'B1:B2');assert.equal(out.getCell('C1').formula,'_xlfn._xlws.FILTER(A3:A3,A3:A3>0)');
    if(!round)original=out.sheetProtection;else assert.deepEqual(out.sheetProtection,original);
    assert.deepEqual(check.definedNames.getRanges('Revenue').ranges,['Details!$A$3']);
  }
});
test('basic formatting edits preserve unrelated advanced fields and update rich-text runs',async()=>{
  const book=await readSpreadsheet(await fixture(),'edit.xlsx'),s=book.sheets[0];s.cells.A1.style={...s.cells.A1.style,bold:true,color:'#006600'};s.cells.A2.raw='바뀐 제목';s.cells.A3.style={...s.cells.A3.style,format:'percent'};s.cells.C1.raw='=FILTER(A3:A3,A3:A3>5)';
  const result=await writeSpreadsheet(book),out=new ExcelJS.Workbook();await out.xlsx.load(result.data);const cells=out.worksheets[0];
  assert.equal(cells.getCell('A1').alignment.textRotation,45);assert.equal(cells.getCell('A1').font.strike,true);
  assert.ok(cells.getCell('A1').value.richText.every(run=>run.font.bold&&run.font.color.argb==='FF006600'));
  assert.equal(cells.getCell('A2').value.text,'바뀐 제목');assert.equal(cells.getCell('A2').value.hyperlink,'https://example.org/manual?q=a&x=b');
  assert.equal(cells.getCell('A3').fill.type,'gradient');assert.equal(cells.getCell('A3').numFmt,'0.00%');
  assert.equal(cells.getCell('C1').formula,'_xlfn._xlws.FILTER(A3:A3,A3:A3>5)');assert.equal(cells.getCell('C1').value.shareType,'array');
  assert.equal(formulaForXlsx('=FILTER(A1:A4,A1:A4>1)+N("FILTER")','_xlfn._xlws.FILTER(A1:A4,A1:A4>0)'), '=_xlfn._xlws.FILTER(A1:A4,A1:A4>1)+N("FILTER")');
});
test('whole-column validation stays a compact original range without allocating cells',async()=>{
  const zip=await JSZip.loadAsync(await fixture());let text=await zip.file('xl/worksheets/sheet1.xml').async('string');text=text.replace('sqref="D1"','sqref="D1:D1048576"');zip.file('xl/worksheets/sheet1.xml',text);
  const book=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'whole-column.xlsx');assert.equal(Object.keys(book.sheets[0].cells).length,7);assert.ok(book.sheets[0].xlsxFeatures.dataValidations.includes('D1:D1048576'));
  const saved=await JSZip.loadAsync((await writeSpreadsheet(book)).data),doc=xml(await saved.file('xl/worksheets/sheet1.xml').async('string'));assert.equal(attr(all(doc,'dataValidation')[0],'sqref'),'D1:D1048576');assert.equal(all(doc,'c').length,7);
});
test('unsupported color scales and icon sets remain saveable with their original priorities and differential styles',async()=>{
  const source=new ExcelJS.Workbook(),s=source.addWorksheet('Conditional');[1,2,3].forEach((value,i)=>s.getCell(i+1,1).value=value);
  s.addConditionalFormatting({ref:'A1:A3',rules:[{type:'colorScale',priority:1,cfvo:[{type:'min'},{type:'max'}],color:[{argb:'FFFF0000'},{argb:'FF00FF00'}]},{type:'cellIs',operator:'greaterThan',formulae:['2'],priority:2,style:{font:{bold:true,color:{argb:'FF0000FF'}}}}]});
  s.addConditionalFormatting({ref:'B1:B3',rules:[{type:'iconSet',iconSet:'3TrafficLights1',priority:3,cfvo:[{type:'percent',value:0},{type:'percent',value:33},{type:'percent',value:67}]}]});
  let bytes=await source.xlsx.writeBuffer();
  for(let i=0;i<3;i++){
    const book=validateWorkbook(await readSpreadsheet(bytes,'conditional.xlsx'));assert.equal(book.sheets[0].unsupportedConditionalCount,2);bytes=(await writeSpreadsheet(book)).data;
    const zip=await JSZip.loadAsync(bytes),doc=xml(await zip.file('xl/worksheets/sheet1.xml').async('string')),styles=xml(await zip.file('xl/styles.xml').async('string'));
    assert.deepEqual(all(doc,'cfRule').map(n=>[attr(n,'type'),attr(n,'priority')]),[['colorScale','1'],['cellIs','2'],['iconSet','3']]);
    const id=Number(attr(all(doc,'cfRule')[1],'dxfId')),dxf=child(styles.documentElement,'dxfs').getElementsByTagName('dxf')[id];assert.equal(attr(all(dxf,'color')[0],'rgb'),'FF0000FF');
  }
});
