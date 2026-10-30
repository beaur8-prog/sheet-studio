import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator,formatValue} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {readPrintSettings,planPrintPages} from '../src/print-layout.mjs';
import {buildSheetPrintDocument} from '../src/sheet-print.mjs';
import {worksheetExtrasFixture} from './fixtures/worksheet-extras.mjs';

function definition(book,sheet,name,refs){
  book.chartNames||=[];book.chartNames.push({name:'_xlnm.'+name,sheetId:sheet.id,formula:refs.map(ref=>"'"+sheet.name.replace(/'/g,"''")+"'!"+ref).join(',')});
}
const pageRefs=html=>[...html.matchAll(/<section class="print-page"[\s\S]*?<\/table>/g)].map(m=>[...m[0].matchAll(/data-ref="([A-Z]+\d+)"/g)].map(m=>m[1]));

test('saved Excel print areas, paper, orientation, margins and fit settings affect output after repeated round trips',async()=>{
  let book=await readSpreadsheet(await worksheetExtrasFixture(),'print.xlsx');
  for(let round=0;round<3;round++){
    const s=book.sheets[0];s.cells.A1000000={raw:'outside private content'};s.cells.B2.raw='=A2*4';
    const before=JSON.stringify(book),settings=readPrintSettings(book,s),out=buildSheetPrintDocument(book,s.id);
    assert.equal(JSON.stringify(book),before);assert.equal(settings.areas.length,1);assert.equal(out.dataCells,60);assert.equal(out.pageCount,1);assert.match(out.html,/size:A4 landscape/);assert.match(out.html,/margin:12\.7mm 10\.16mm 15\.24mm 7\.62mm/);assert.match(out.html,/>12<\/div>/);assert.doesNotMatch(out.html,/outside private content|data-ref="A1000000"/);
    book=await readSpreadsheet((await writeSpreadsheet(book)).data,'again.xlsx');
  }
});
test('nonadjacent areas stay separate and sheet names containing commas, quotes and exclamation marks are parsed safely',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.name="Q1, Bob's!";s.cells={A1:{raw:'first'},C3:{raw:'second'},B2:{raw:'not selected'}};definition(b,s,'Print_Area',['$A$1:$A$2','$C$3']);
  const out=buildSheetPrintDocument(b,s.id);assert.equal(out.pageCount,2);assert.equal(out.dataCells,3);assert.deepEqual(pageRefs(out.html),[['A1','A2'],['C3']]);assert.doesNotMatch(out.html,/not selected/);
});
test('manual row/column breaks repeat titles on every page and page order follows saved Excel settings',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.xlsxFeatures={pageSetup:'<pageSetup paperSize="9" scale="100" pageOrder="downThenOver"/>',rowBreaks:'<rowBreaks><brk id="4" min="0" max="16383" man="1"/></rowBreaks>',colBreaks:'<colBreaks><brk id="3" min="0" max="1048575" man="true"/></colBreaks>'};
  definition(b,s,'Print_Area',['A1:E8']);definition(b,s,'Print_Titles',['$1:$1','$A:$A']);
  const range={r1:0,r2:7,c1:0,c2:4},settings=readPrintSettings(b,s),plan=planPrintPages(s,[range],settings);
  assert.equal(plan.pages.length,4);assert.equal(plan.dataCells,40);assert.deepEqual(plan.pages.map(p=>[p.rows,p.cols]),[[[0,1,2,3],[0,1,2]],[[0,4,5,6,7],[0,1,2]],[[0,1,2,3],[0,3,4]],[[0,4,5,6,7],[0,3,4]]]);
  settings.pageOrder='overThenDown';const over=planPrintPages(s,[range],settings);assert.deepEqual(over.pages[1].cols,[0,3,4]);assert.deepEqual(over.pages[1].rows,[0,1,2,3]);
  s.cells.A1={raw:'Repeat <title>'};const out=buildSheetPrintDocument(b,s.id);assert.equal(out.pageCount,4);assert.equal((out.html.match(/Repeat &lt;title&gt;/g)||[]).length,4);
});
test('automatic page sizing, hidden dimensions, outside title rows and visible manual boundaries work together',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.defaultRowHeight=200;s.defaultColumnWidth=250;s.columnWidths={0:100,2:0};s.rowHeights={2:0};s.xlsxFeatures={pageSetup:'<pageSetup paperSize="9" scale="100"/>',rowBreaks:'<rowBreaks><brk id="3" man="1"/></rowBreaks>'};definition(b,s,'Print_Titles',['1:1','A:A']);
  const settings=readPrintSettings(b,s),out=planPrintPages(s,[{r1:1,r2:7,c1:1,c2:4}],settings);assert.equal(out.pages.length,6);assert.equal(out.dataCells,18);
  for(const p of out.pages){assert.equal(p.rows[0],0);assert.equal(p.cols[0],0);assert.ok(!p.rows.includes(2));assert.ok(!p.cols.includes(2));assert.ok(p.rows.reduce((n,r)=>n+(s.rowHeights[r]??200),0)*p.scale<=settings.contentHeight+.01);}
});
test('fit-to-page suppresses manual breaks while explicit percentages preserve them',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.xlsxFeatures={pageSetup:'<pageSetup fitToWidth="1" fitToHeight="1"/>',sheetPr:'<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',rowBreaks:'<rowBreaks><brk id="3" man="1"/></rowBreaks>'};
  const range={r1:0,r2:89,c1:0,c2:9};let out=planPrintPages(s,[range],readPrintSettings(b,s));assert.equal(out.pages.length,1);assert.ok(out.pages[0].scale<1);
  s.xlsxFeatures.sheetPr='<sheetPr/>';s.xlsxFeatures.pageSetup='<pageSetup scale="50"/>';out=planPrintPages(s,[range],readPrintSettings(b,s));assert.equal(out.pages.length,3);assert.equal(out.pages[0].scale,.5);assert.deepEqual(out.pages[0].rows,[0,1,2]);
});
test('selection and chart-only printing bypass bad saved areas/titles without mutating workbook settings',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'=RAND()'};s.charts=[{id:'c1',type:'bar',range:'A1:A1',firstRowNames:false,firstColCategories:false}];definition(b,s,'Print_Area',['#REF!']);definition(b,s,'Print_Titles',['#REF!']);
  assert.throws(()=>buildSheetPrintDocument(b,s.id),/인쇄 영역/);const before=JSON.stringify(b),evaluate=createEvaluator(b,s.id),value=evaluate('A1');const out=buildSheetPrintDocument(b,s.id,{scope:'selection',selection:{r1:0,r2:0,c1:0,c2:0},includeCharts:false,evaluate});assert.equal(out.dataCells,1);assert.ok(out.html.includes(formatValue(value,{},b)));assert.equal(buildSheetPrintDocument(b,s.id,{scope:'charts'}).chartCount,1);assert.equal(JSON.stringify(b),before);
  b.chartNames=b.chartNames.filter(n=>!n.name.endsWith('Print_Titles'));assert.equal(buildSheetPrintDocument(b,s.id,{ignorePrintArea:true}).dataCells,1);
});
test('oversized or invalid print metadata fails before output expansion rather than printing unrelated cells',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'safe'};definition(b,s,'Print_Area',['A1:A20001']);assert.throws(()=>buildSheetPrintDocument(b,s.id),/20,000/);
  b.chartNames=[];definition(b,s,'Print_Titles',['1:1048576']);assert.throws(()=>buildSheetPrintDocument(b,s.id),/반복 제목이 너무 큽니다/);
  b.chartNames=[];s.xlsxFeatures={pageMargins:'<pageMargins left="20" right="20"/>'};assert.throws(()=>buildSheetPrintDocument(b,s.id),/여백/);
  s.xlsxFeatures={pageSetup:'<pageSetup paperSize="999"/>'};const out=buildSheetPrintDocument(b,s.id);assert.match(out.notes[0],/A4/);
  s.xlsxFeatures.pageSetup='<pageSetup paperSize="13" orientation="landscape"/>';assert.match(buildSheetPrintDocument(b,s.id).html,/size:257mm 182mm/);
  s.defaultRowHeight=1;s.xlsxFeatures={pageSetup:'<pageSetup scale="100"/>'};const settings=readPrintSettings(b,s);settings.repeatRows={start:0,end:999};settings.contentHeight=1005;
  assert.throws(()=>planPrintPages(s,[{r1:1000,r2:1999,c1:0,c2:0}],settings),/인쇄 분량이 너무 커졌습니다/);
});
test('merged text is printed once across page boundaries while a clipped selection still shows the original merged value',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'MERGED_VALUE'};s.merges=['A1:C6'];s.xlsxFeatures={pageSetup:'<pageSetup scale="100"/>',rowBreaks:'<rowBreaks><brk id="3" man="1"/></rowBreaks>',colBreaks:'<colBreaks><brk id="2" man="1"/></colBreaks>'};
  const out=buildSheetPrintDocument(b,s.id);assert.equal(out.pageCount,4);assert.equal((out.html.match(/MERGED_VALUE/g)||[]).length,1);
  const selected=buildSheetPrintDocument(b,s.id,{scope:'selection',selection:{r1:4,r2:5,c1:1,c2:2}});assert.equal((selected.html.match(/MERGED_VALUE/g)||[]).length,1);
});
test('fit-to-two-pages includes the space consumed by repeated titles on both pages',()=>{
  const b=createWorkbook(),s=b.sheets[0];s.defaultColumnWidth=150;s.xlsxFeatures={pageSetup:'<pageSetup fitToWidth="2" fitToHeight="0"/>',sheetPr:'<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'};definition(b,s,'Print_Titles',['A:A']);
  const out=planPrintPages(s,[{r1:0,r2:1,c1:0,c2:7}],readPrintSettings(b,s));assert.equal(out.pages.length,2);assert.ok(out.pages[0].scale<1);assert.deepEqual(out.pages[1].cols,[0,5,6,7]);
});
