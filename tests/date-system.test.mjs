import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {createWorkbook,evaluateFormula,evaluateCell,formatValue,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {chartData} from '../src/chart-model.mjs';
import {renderChartSVG} from '../src/chart-render.mjs';
import {applyFilters,filterOptions} from '../src/filters.mjs';
import {nativeFilterPlan} from '../src/filter-native.mjs';
import {pastePlan} from '../src/cell-paste.mjs';
import {copyTable} from '../src/table-exchange.mjs';
import {xml,all,attr} from '../src/chart-xml.mjs';

const calc=(book,formula)=>evaluateFormula(book,book.activeSheetId,formula);
function dates(){
  const book=createWorkbook();book.date1904=true;
  book.sheets[0].cells={A1:{raw:'날짜'},B1:{raw:'값'},A2:{raw:'39267',style:{format:'date'}},A3:{raw:'39268',style:{format:'date'}},B2:{raw:'10'},B3:{raw:'20'}};
  return book;
}

test('the 1904 calendar is local to each workbook and applies to all date adapters',()=>{
  const modern=createWorkbook(),old=createWorkbook();old.date1904=true;
  for(const [formula,expected] of [
    ['=DATE(2011,7,5)',39267],['=YEAR(39267)',2011],['=MONTH(39267)',7],['=DAY(39267)',5],
    ['=DATEVALUE("2011-07-05")',39267],['=WEEKDAY(39267)',3],['=YEAR(0)',1904],['=WEEKDAY(0)',6],
    ['=TEXT(39267,"yyyy-mm-dd")','2011-07-05'],['=EOMONTH(39267,0)',39293],['=EDATE(39267,1)',39298],
    ['=NETWORKDAYS(DATE(2011,7,4),DATE(2011,7,8))',5],['=WORKDAY(39267,5)',39274],
  ])assert.equal(calc(old,formula),expected,formula);
  assert.equal(calc(modern,'=DATE(2011,7,5)'),40729);
  assert.equal(calc(old,'=TODAY()'),calc(modern,'=TODAY()')-1462);
  assert.equal(formatValue(39267,{format:'date'},old),'2011-07-05');
  assert.equal(formatValue(39267,{format:'date'},modern),'2007-07-04');
  assert.equal(parseWorkbook(serializeWorkbook(old)).date1904,true);
});

test('1904 XLSX preserves raw date constants, formula results and the workbook epoch over repeated saves',async()=>{
  const source=new ExcelJS.Workbook();source.properties.date1904=true;
  const sheet=source.addWorksheet('Dates');
  sheet.getCell('A1').value=new Date('2011-07-05T00:00:00Z');sheet.getCell('A1').numFmt='yyyy-mm-dd';
  sheet.getCell('B1').value={formula:'DATE(2011,7,5)-39267',result:0};
  sheet.getCell('C1').value={formula:'A1+1',result:new Date('2011-07-06T00:00:00Z')};sheet.getCell('C1').numFmt='yyyy-mm-dd';
  let book=await readSpreadsheet(await source.xlsx.writeBuffer(),'1904.xlsx');
  for(let i=0;i<3;i++){
    const current=book.sheets[0];assert.equal(book.date1904,true);assert.equal(current.cells.A1.raw,'39267');
    assert.equal(evaluateCell(book,current.id,'B1'),0);assert.equal(evaluateCell(book,current.id,'C1'),39268);
    assert.equal(current.cells.C1.cachedValue,39268);
    const output=await writeSpreadsheet(book),check=new ExcelJS.Workbook();await check.xlsx.load(output.data);
    assert.equal(check.properties.date1904,true);assert.equal(check.worksheets[0].getCell('B1').result,0);
    assert.equal(check.worksheets[0].getCell('A1').value.toISOString(),'2011-07-05T00:00:00.000Z');
    book=await readSpreadsheet(output.data,'1904.xlsx');
  }
});

test('1904 conditional date functions and filters share the document calendar',async()=>{
  const book=dates(),sheet=book.sheets[0];
  assert.equal(calc(book,'=COUNTIF(A2:A3,">2011-07-05")'),1);
  assert.equal(calc(book,'=SUMIFS(B2:B3,A2:A3,">7/5/2011")'),20);
  sheet.autoFilter={range:{r1:0,r2:2,c1:0,c2:1},columns:{0:{type:'custom',kind:'date',join:'and',conditions:[{operator:'greaterThan',value:'2011-07-05'}]}},hiddenRows:[]};
  applyFilters(book,sheet);assert.deepEqual(sheet.autoFilter.hiddenRows,[1]);
  assert.deepEqual(filterOptions(book,sheet,0).values.map(v=>v.label),['2011-07-05','2011-07-06']);
  assert.equal(nativeFilterPlan(sheet.autoFilter,book).filters[0].criteria1,'>39267');
  let output=await writeSpreadsheet(book),zip=await JSZip.loadAsync(output.data),doc=xml(await zip.file('xl/worksheets/sheet1.xml').async('string'));
  assert.equal(attr(all(doc,'customFilter')[0],'val'),'39267');
  let loaded=await readSpreadsheet(output.data,'dates.xlsx');applyFilters(loaded,loaded.sheets[0]);assert.deepEqual(loaded.sheets[0].autoFilter.hiddenRows,[1]);
  sheet.autoFilter.columns={0:{type:'values',kind:'date',values:['39267']}};applyFilters(book,sheet);assert.deepEqual(sheet.autoFilter.hiddenRows,[2]);
  output=await writeSpreadsheet(book);zip=await JSZip.loadAsync(output.data);doc=xml(await zip.file('xl/worksheets/sheet1.xml').async('string'));
  const day=all(doc,'dateGroupItem')[0];assert.deepEqual(['year','month','day'].map(key=>attr(day,key)),['2011','7','5']);
  loaded=await readSpreadsheet(output.data,'dates.xlsx');assert.deepEqual(loaded.sheets[0].autoFilter.columns[0].values,['39267']);
});

test('1904 dates render consistently in charts, chart XML and rich clipboard text',async()=>{
  const book=dates(),sheet=book.sheets[0],chart={id:'dates',type:'line',range:'A1:B3',firstColumnAsLabels:true,width:600,height:360};
  sheet.charts.push(chart);
  assert.deepEqual(chartData(chart,sheet,book)[0].labels,['2011-07-05','2011-07-06']);
  assert.match(renderChartSVG(chart,sheet,book).svg,/2011-07-05/);
  const output=await writeSpreadsheet(book),zip=await JSZip.loadAsync(output.data),doc=xml(await zip.file('xl/charts/sheetStudioChart1.xml').async('string'));
  assert.equal(attr(all(doc,'date1904')[0],'val'),'1');
  const clip={r:1,c:0,data:[[sheet.cells.A2]],values:[[39267]],merges:[],date1904:true};
  assert.equal(copyTable(clip,{date1904:true}).text,'2011-07-05');
  const target=createWorkbook().sheets[0];
  assert.equal(pastePlan(target,clip,{r:0,c:0},{date1904:false}).updates[0].cell.raw,'40729');
  assert.equal(pastePlan(target,{...clip,date1904:false,data:[[{raw:'40729',style:{format:'date'}}]],values:[[40729]]},{r:0,c:0},{date1904:true,mode:'values'}).updates[0].cell.raw,'39267');
  const numberClip={...clip,data:[[{raw:'39267'}]]};assert.equal(pastePlan(target,numberClip,{r:0,c:0},{date1904:false}).updates[0].cell.raw,'39267');
});
