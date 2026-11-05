import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {createWorkbook,createEvaluator,validateWorkbook,rewriteWorkbookSheetReferences} from '../src/engine.mjs';
import {createConditionalEvaluator,shiftWorksheetFeatures} from '../src/conditional.mjs';
import {validateConditionalRules} from '../src/worksheet-schema.mjs';
import {dataBarCSS} from '../src/conditional-visual.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {buildSheetPrintDocument} from '../src/sheet-print.mjs';
import {xml,all,attr} from '../src/chart-xml.mjs';
const scale=(overrides={})=>({id:'scale',range:'A1:A9',type:'colorScale',formulae:[],style:{},stopIfTrue:false,cfvo:[{type:'min'},{type:'percent',value:50},{type:'max'}],colors:['#ff0000','#ffffff','#00ff00'],...overrides});
const bar=(overrides={})=>({id:'bar',range:'B1:B9',type:'dataBar',formulae:[],style:{},stopIfTrue:false,cfvo:[{type:'min'},{type:'max'}],barColor:'#638ec6',...overrides});
function fixture(){const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={};for(const c of ['A','B'])[-10,0,10,'',"'10",'TRUE','#N/A','=""','=5+0'].forEach((raw,i)=>sheet.cells[c+(i+1)]={raw:String(raw)});sheet.conditionalRules=[scale(),bar()];return {book,sheet};}
test('two/three-color scales and base data bars compute range thresholds from numbers, not blanks, text, Boolean or errors',()=>{
  const {book,sheet}=fixture(),evaluate=createEvaluator(book,sheet.id),style=createConditionalEvaluator(book,sheet,evaluate);
  assert.deepEqual(style(0,0,-10),{background:'#ff0000'});assert.deepEqual(style(1,0,0),{background:'#ffffff'});assert.deepEqual(style(2,0,10),{background:'#00ff00'});assert.deepEqual(style(8,0,5),{background:'#80ff80'});
  assert.equal(style(0,1,-10).dataBar.percent,10);assert.equal(style(1,1,0).dataBar.percent,50);assert.equal(style(2,1,10).dataBar.percent,90);
  for(const r of [3,4,5,6,7]){assert.deepEqual(style(r,0,evaluate('A'+(r+1))),{});assert.deepEqual(style(r,1,evaluate('B'+(r+1))),{});}
  sheet.conditionalRules=[scale({cfvo:[{type:'num',value:-20},{type:'num',value:20}],colors:['#000000','#ffffff']})];const two=createConditionalEvaluator(book,sheet);assert.equal(two(1,0,0).background,'#808080');
});
test('percentile thresholds use the sorted distribution and are different from a percentage of the numeric interval',()=>{
  const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'0'},A2:{raw:'1'},A3:{raw:'2'},A4:{raw:'100'}};
  sheet.conditionalRules=[scale({range:'A1:A4',cfvo:[{type:'min'},{type:'percentile',value:50},{type:'max'}]})];const percentile=createConditionalEvaluator(book,sheet);assert.equal(percentile(0,0,1.5).background,'#ffffff');
  sheet.conditionalRules=[scale({range:'A1:A4'})];assert.notEqual(createConditionalEvaluator(book,sheet)(0,0,1.5).background,'#ffffff');
});
test('visual rules obey existing property priority and ordinary stop rules, and include computed spill followers in range statistics',()=>{
  const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'=SEQUENCE(3)'}};sheet.conditionalRules=[scale({range:'A1:A3'})];const evaluate=createEvaluator(book,sheet.id);evaluate('A1');let style=createConditionalEvaluator(book,sheet,evaluate);assert.equal(style(2,0,evaluate('A3')).background,'#00ff00');
  sheet.conditionalRules.unshift({id:'stop',range:'A1:A3',type:'cellIs',operator:'equal',formulae:['2'],style:{background:'#0000ff',bold:true},stopIfTrue:true});style=createConditionalEvaluator(book,sheet,evaluate);assert.deepEqual(style(1,0,2),{background:'#0000ff',bold:true});assert.equal(style(2,0,3).background,'#00ff00');
});
test('invalid visual thresholds, colors and unsupported stop flags reject without generating invalid files',()=>{
  for(const r of [scale({colors:['#fff']}),scale({cfvo:[{type:'num',value:10},{type:'num',value:0}],colors:['#000000','#ffffff']}),scale({cfvo:[{type:'min'},{type:'percentile',value:101},{type:'max'}]}),scale({stopIfTrue:true}),bar({cfvo:[{type:'min'},{type:'num',value:NaN}]}),bar({barColor:'url(evil)'})])assert.throws(()=>validateConditionalRules([r]));
  assert.equal(dataBarCSS({dataBar:{color:'url(evil)',percent:50}}),'');
});
test('print and grid use the same conditional colors and gradient bar dimensions',()=>{
  const {book,sheet}=fixture(),print=buildSheetPrintDocument(book,sheet.id,{scope:'selection',selection:{r1:0,r2:8,c1:0,c2:1}});
  assert.match(print.html,/background:#ff0000/);assert.match(print.html,/background:#00ff00/);assert.match(print.html,/background-size:10%/);assert.match(print.html,/background-size:50%/);assert.match(print.html,/background-size:90%/);assert.match(print.html,/linear-gradient/);
});
test('editable visual rules, values and row-transformed ranges survive three XLSX saves and independent ExcelJS reads',async()=>{
  let {book,sheet}=fixture();shiftWorksheetFeatures(sheet,'r',0,1,sheet.name);sheet.cells=Object.fromEntries(Object.entries(sheet.cells).map(([key,cell])=>[key.replace(/\d+$/,n=>Number(n)+1),cell]));assert.equal(sheet.conditionalRules[0].range,'A2:A10');
  for(let i=0;i<3;i++){
    const {data}=await writeSpreadsheet(book),zip=await JSZip.loadAsync(data),doc=xml(await zip.file('xl/worksheets/sheet1.xml').async('string'));
    assert.equal(all(doc,'colorScale').length,1);assert.equal(all(doc,'dataBar').length,1);assert.equal(all(doc,'cfRule').length,2);assert.deepEqual(all(doc,'cfRule').map(n=>attr(n,'type')),['colorScale','dataBar']);
    const independent=new ExcelJS.Workbook();await independent.xlsx.load(data);assert.equal(independent.worksheets[0].conditionalFormattings[0].rules[0].type,'colorScale');assert.equal(independent.worksheets[0].conditionalFormattings[1].rules[0].type,'dataBar');
    book=validateWorkbook(await readSpreadsheet(data,'visual.xlsx'));sheet=book.sheets[0];assert.equal(sheet.unsupportedConditionalCount,undefined);assert.equal(sheet.conditionalRules[0].range,'A2:A10');assert.equal(sheet.conditionalRules[1].barColor,'#638ec6');assert.equal(createConditionalEvaluator(book,sheet)(2,0,0).background,'#ffffff');
  }
});
test('advanced x14 data bars keep their source XML and are not offered as editable basic rules',async()=>{
  const source=new ExcelJS.Workbook(),s=source.addWorksheet('Advanced');s.getCell('A1').value=-10;s.getCell('A2').value=20;s.addConditionalFormatting({ref:'A1:A2',rules:[{type:'dataBar',priority:1,cfvo:[{type:'min'},{type:'max'}],color:{argb:'FF638EC6'},gradient:false,negativeFillColor:{argb:'FFFF0000'},axisColor:{argb:'FF000000'}}]});
  const input=await source.xlsx.writeBuffer(),book=await readSpreadsheet(input,'advanced.xlsx');assert.ok(book.sheets[0].unsupportedConditionalCount);assert.ok(!book.sheets[0].conditionalRules?.length);
  const saved=await JSZip.loadAsync((await writeSpreadsheet(book)).data),text=await saved.file('xl/worksheets/sheet1.xml').async('string');assert.match(text,/dataBar/);assert.match(text,/x14:id/);assert.match(text,/<x14:dataBar/);assert.match(text,/FFFF0000/);assert.match(text,/x14:negativeFillColor/);
});

test('advanced definitions keep base GUID links, negative colors, relative ranges and names through three transformed roundtrips',async()=>{
  const source=new ExcelJS.Workbook(),s=source.addWorksheet('Advanced');source.addWorksheet('Bounds').getCell('A1').value=10;
  s.getCell('A1').value=-10;s.getCell('A2').value=20;s.addConditionalFormatting({ref:'A1:A2',rules:[{type:'dataBar',priority:1,cfvo:[{type:'formula',value:'-Bounds!$A$1'},{type:'max'}],color:{argb:'FF638EC6'},gradient:false,negativeFillColor:{argb:'FFFF0000'},axisColor:{argb:'FF000000'}}]});
  let book=await readSpreadsheet(await source.xlsx.writeBuffer(),'advanced.xlsx');let sheet=book.sheets[0];
  shiftWorksheetFeatures(sheet,'r',0,1,sheet.name);sheet.cells={A2:{raw:'-10'},A3:{raw:'20'}};rewriteWorkbookSheetReferences(book,'Bounds','Limits');book.sheets[1].name='Limits';
  for(let round=0;round<3;round++){
    const data=(await writeSpreadsheet(book)).data,zip=await JSZip.loadAsync(data),doc=xml(await zip.file('xl/worksheets/sheet1.xml').async('string'));
    const legacy=all(doc,'id').find(n=>n.namespaceURI==='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'),advanced=all(doc,'cfRule').find(n=>n.namespaceURI==='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main');
    assert.ok(legacy?.textContent);assert.equal(attr(advanced,'id'),legacy.textContent);assert.ok(all(advanced,'f').some(n=>n.textContent==="-'Limits'!$A$1"));assert.ok(all(doc,'cfvo').some(n=>attr(n,'val')==="-'Limits'!$A$1"));assert.ok(all(advanced,'negativeFillColor').some(n=>attr(n,'rgb')==='FFFF0000'));
    assert.equal(all(doc,'sqref').find(n=>n.namespaceURI?.includes('excel/2006/main'))?.textContent,'A2:A3');
    assert.equal(attr(all(doc,'conditionalFormatting').find(n=>n.namespaceURI==='http://schemas.openxmlformats.org/spreadsheetml/2006/main'),'sqref'),'A2:A3');
    book=validateWorkbook(await readSpreadsheet(data,'advanced.xlsx'));sheet=book.sheets[0];assert.ok(sheet.unsupportedConditionalCount);assert.equal(sheet.xlsxFeatures.conditionalExtensions.length,1);
  }
  shiftWorksheetFeatures(sheet,'r',1,-2,sheet.name);assert.equal(sheet.unsupportedConditionalCount,undefined);assert.deepEqual(sheet.xlsxFeatures.conditionalExtensions,[]);
});
