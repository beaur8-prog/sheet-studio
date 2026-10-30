import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator,setExternalEvaluator,formatValue} from '../src/engine.mjs';
import {buildSheetPrintDocument} from '../src/sheet-print.mjs';
import {readSpreadsheet} from '../src/io.mjs';
import {tableFixture} from './tables-fixture.mjs';

test('printed sheets retain merges, dimensions, filters, format sections, current arrays and typed errors',()=>{
  const book=createWorkbook(),s=book.sheets[0];book.name='<Book>';s.name='Report';s.merges=['A1:C1'];s.columnWidths={0:130,1:100,2:0};s.rowHeights={3:0};
  s.cells={A1:{raw:'Merged <title>',style:{bold:true,background:'#e2efda'}},A2:{raw:'hidden'},A3:{raw:'-12',style:{numberFormat:'0;[Red](0)',borders:{bottom:{style:'medium',color:'#008800'}}}},B3:{raw:'={5;6}',xlsxFormula:{raw:'={5;6}',formula:'{5;6}',shareType:'array',ref:'B3:B4'}},A4:{raw:'manual hidden'},A5:{raw:'="#N/A"',style:{numberFormat:'0;0;0;"text: "@'}},B5:{raw:'=NA()',style:{numberFormat:'0;0;0;"text: "@'}}};
  s.autoFilter={range:{r1:0,r2:4,c1:0,c2:1},columns:{},hiddenRows:[1]};
  const before=JSON.stringify(book),out=buildSheetPrintDocument(book,s.id,{gridLines:true});assert.equal(JSON.stringify(book),before);
  assert.equal(out.dataCells,6);assert.match(out.html,/rowspan="1" colspan="2"/);assert.match(out.html,/Merged &lt;title&gt;/);assert.doesNotMatch(out.html,/manual hidden|>hidden<|data-row="1"/);assert.match(out.html,/color:#ff0000/);assert.match(out.html,/border-bottom:2px solid #008800/);assert.match(out.html,/data-ref="B3"[^>]*><div[^>]*>5</);assert.match(out.html,/>text: #N\/A</);assert.match(out.html,/data-ref="B5"[^>]*><div[^>]*>#N\/A</);assert.match(out.html,/width:230px/);assert.match(out.html,/Content-Security-Policy/);
  const selected=buildSheetPrintDocument(book,s.id,{scope:'selection',selection:{r1:0,r2:2,c1:1,c2:2}});assert.match(selected.html,/rowspan="1" colspan="1"/);assert.match(selected.html,/Merged &lt;title&gt;/);
});
test('Excel table print colors, current calculated totals and conditional styles are included',async()=>{
  const book=await readSpreadsheet(await tableFixture(),'print.xlsx'),s=book.sheets[0];s.cells.A3.raw='6';s.conditionalRules=[{id:'high',range:'C2:C3',type:'cellIs',operator:'greaterThan',formulae:['20'],style:{background:'#ffcc00'}}];
  const out=buildSheetPrintDocument(book,s.id);assert.match(out.html,/data-ref="C3"[^>]*><div[^>]*>30</);assert.match(out.html,/data-ref="C4"[^>]*><div[^>]*>36</);assert.match(out.html,/background:#ffcc00/);assert.match(out.html,/font-weight:bold/);
});
test('individual chart pages remain available and chart-only sheets print actual SVG rather than an empty grid',()=>{
  const book=createWorkbook(),s=book.sheets[0];s.cells={A1:{raw:'Jan'},B1:{raw:'2'},A2:{raw:'Feb'},B2:{raw:'5'}};s.charts=[{id:'c1',type:'bar',title:'Sales & costs',range:'A1:B2',firstRowNames:false,firstColCategories:true,width:500,height:300,legend:false}];
  const out=buildSheetPrintDocument(book,s.id,{orientation:'landscape',objectLayout:'pages'});assert.equal(out.chartCount,1);assert.equal(out.pageCount,2);assert.match(out.html,/size:A4 landscape/);assert.match(out.html,/class="chart-point"/);assert.match(out.html,/Sales &amp; costs/);assert.ok(out.html.indexOf('<table')<out.html.indexOf('<figure'));assert.equal(buildSheetPrintDocument(book,s.id,{includeCharts:false}).chartCount,0);
  const chartSheet={id:'chart',name:'Chart',kind:'chart',cells:{},charts:[{...s.charts[0],series:[{name:'Sales',valRef:{sheetId:s.id,range:'B1:B2'},catRef:{sheetId:s.id,range:'A1:A2'}}]}]};book.sheets.push(chartSheet);const chart=buildSheetPrintDocument(book,'chart');assert.equal(chart.dataCells,0);assert.equal(chart.chartCount,1);assert.match(chart.html,/class="chart-point"/);assert.doesNotMatch(chart.html,/<table/);
});
test('selected and chart-only printing avoid large data ranges and use a shared calculation snapshot',()=>{
  const book=createWorkbook(),s=book.sheets[0];s.cells={A1:{raw:'=RAND()'},A1000000:{raw:'last'}};s.charts=[{id:'random',type:'bar',range:'A1:A1',firstRowNames:false,firstColCategories:false,title:'Random'}];
  assert.throws(()=>buildSheetPrintDocument(book,s.id),/20,000/);assert.equal(buildSheetPrintDocument(book,s.id,{scope:'charts'}).chartCount,1);
  assert.throws(()=>buildSheetPrintDocument(book,s.id,{scope:'selection'}),/선택/);
  const evaluate=createEvaluator(book,s.id),result=evaluate.result('A1'),snapshot={...book},shared=(id,ref)=>evaluate(ref);shared.result=(id,ref)=>evaluate.result(ref);shared.isComputed=(id,ref)=>evaluate.isComputed(ref);setExternalEvaluator(snapshot,shared);
  const print=buildSheetPrintDocument(snapshot,s.id,{scope:'selection',selection:{r1:0,r2:0,c1:0,c2:0},evaluate});assert.equal(print.dataCells,1);assert.ok(print.html.includes(formatValue(result.value,{},book)));assert.equal(evaluate('A1'),result.value);
  const hidden=createWorkbook(),hs=hidden.sheets[0];hs.cells.A1={raw:'hidden'};hs.rowHeights={0:0};const empty=buildSheetPrintDocument(hidden,hs.id);assert.equal(empty.dataCells,0);assert.match(empty.html,/인쇄할 표시 셀이 없습니다/);assert.doesNotMatch(empty.html,/>hidden</);
});
