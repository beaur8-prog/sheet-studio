import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {createWorkbook,evaluateFormula,formatValue,numberFormatColor,numberFormatKind} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {applyFilters,filterOptions,filterValue} from '../src/filters.mjs';
import {createEvaluator} from '../src/engine.mjs';
import {effectiveNumberFormat,validateNumberFormat} from '../src/number-formats.mjs';

test('documented custom numeric codes preserve digits, optional places, scaling and scientific notation',()=>{
  const cases=[
    [1234.59,'####.#','1234.6'],[8.9,'#.000','8.900'],[.631,'0.#','0.6'],[12,'#.0#','12.0'],[1234.568,'#.0#','1234.57'],[.47,'#.##','.47'],
    [12000,'#,###','12,000'],[12000,'#,','12'],[12200000,'0.0,,"M"','12.2M'],[12200000,'0.00E+00','1.22E+07'],[12200000,'#0.0E+0','12.2E+6'],
    [12,'00000','00012'],[-12.5,'00000','-00013'],[.125,'0.0%','12.5%'],[12,'0"%"','12%'],[.125,'0.0\\%','0.1%'],
    [5.25,'# ?/?','5 1/4'],[5.3,'# ??/??','5  3/10'],[1.25,'# ?/8','1 2/8'],[.125,'?/8','1/8'],[1.5,'??/??',' 3/ 2'],
    [123456789,'000-00-0000','123-45-6789'],[8.9,'0.0?','8.9 '],[0,'#',''],[1234.5,'[$€-407]#,##0.00','€1,234.50']
  ];
  for(const [n,code,want] of cases)assert.equal(formatValue(n,{numberFormat:code}),want,code+' '+n);
});
test('positive, negative, zero, text, conditions, escaped semicolons and colors use their own sections',()=>{
  const code='[Blue]#,##0.00;[Red](#,##0.00);"-";"item: "@';
  assert.equal(formatValue(12.5,{numberFormat:code}),'12.50');assert.equal(formatValue(-12.5,{numberFormat:code}),'(12.50)');assert.equal(formatValue(0,{numberFormat:code}),'-');assert.equal(formatValue('abc',{numberFormat:code}),'item: abc');
  assert.equal(numberFormatColor(-1,{numberFormat:code}),'#ff0000');assert.equal(numberFormatColor(1,{numberFormat:code}),'#0000ff');
  for(const value of [-1,0,1])assert.equal(formatValue(value,{numberFormat:';;;'}),'');
  assert.equal(formatValue('text',{numberFormat:';;;'}),'');assert.equal(formatValue(1,{numberFormat:'0"; x";(0)'}),'1; x');assert.equal(formatValue(-1,{numberFormat:'0\\;;(0)'}),'(1)');
  assert.equal(formatValue(200,{numberFormat:'[>=100]0.0;"small"'}),'200.0');assert.equal(formatValue(20,{numberFormat:'[>=100]0.0;"small"'}),'small');
  assert.equal(formatValue(0,{numberFormat:'[>0]0;[<0]0'}),'########');assert.equal(formatValue(-20,{numberFormat:'[Red]General'}),'-20');
  assert.equal(numberFormatKind('@'),'text');assert.equal(numberFormatKind('0"months"'),'number');assert.equal(formatValue('#N/A',{numberFormat:code},{error:'#N/A'}),'#N/A');
});
test('TEXT and cells share section, numeric string, mask and fraction behavior',()=>{
  const b=createWorkbook(),s=b.sheets[0];for(const [raw,want] of [
    ['=TEXT(-1234.5,"#,##0.00;(#,##0.00)")','(1,234.50)'],['=TEXT("12","00000")','00012'],['=TEXT(12200000,"0.0,,""M""")','12.2M'],
    ['=TEXT(5.25,"# ?/?")','5 1/4'],['=TEXT("abc","0;0;0;""item: ""@")','item: abc'],['=TEXT(0,"0;(0);""none""")','none']
  ])assert.equal(evaluateFormula(b,s.id,raw),want,raw);
});
test('elapsed and fractional seconds, single-letter months and decimal edits keep meaning',()=>{
  assert.equal(formatValue(61.5625,{numberFormat:'mmmmm d, yyyy h:mm am/pm'}),'M 1, 1900 1:30 pm');
  assert.equal(formatValue(3735.8/86400,{numberFormat:'[ss].00'}),'3735.80');assert.equal(formatValue(3735.8/86400,{numberFormat:'h:mm:ss.00'}),'1:02:15.80');
  assert.equal(formatValue(61+86399.9999/86400,{numberFormat:'yyyy-mm-dd hh:mm:ss.00'}),'1900-03-02 00:00:00.00');
  assert.equal(formatValue(61+86399.9999/86400,{numberFormat:'yyyy-mm-dd'}),'1900-03-01');
  assert.equal(effectiveNumberFormat({format:'number',numberFormat:'0.0,,"M";(0.0,,"M")',decimals:3}),'0.000,,"M";(0.000,,"M")');
  for(const code of ['0"x','0[Red','0;0;0;@;0','0\\'])assert.throws(()=>validateNumberFormat(code));
});
test('number-code colors participate in filters and absent array followers are not treated as empty',()=>{
  const book=createWorkbook(),s=book.sheets[0],code='0;[Red](0)';
  s.cells={A1:{raw:'Value'},A2:{raw:'2',style:{numberFormat:code}},A3:{raw:'-2',style:{numberFormat:code}},B1:{raw:'Array'},B2:{raw:'={1;2}',xlsxFormula:{raw:'={1;2}',formula:'{1;2}',shareType:'array',ref:'B2:B3'}}};
  s.autoFilter={range:{r1:0,r2:2,c1:0,c2:1},columns:{0:{type:'color',target:'color',color:'#ff0000'}}};applyFilters(book,s);assert.deepEqual(s.autoFilter.hiddenRows,[1]);
  assert.equal(filterOptions(book,s,0).colors.color.find(c=>c.color==='#ff0000').count,1);assert.equal(filterValue(s,createEvaluator(book,s.id),2,1),2);
});
test('custom sections and masks survive three XLSX saves with original numeric and text values',async()=>{
  const b=createWorkbook(),s=b.sheets[0],styles=['00000','#,##0.00;[Red](#,##0.00);"-";"item: "@','0.0,,"M"','# ?/?','0.00E+00','@'],raws=['12','-1234.5','12200000','5.25','12200000',"'123"];
  styles.forEach((code,i)=>s.cells['A'+(i+1)]={raw:raws[i],style:{format:numberFormatKind(code),numberFormat:code}});
  let book=b;
  for(let round=0;round<3;round++){
    const out=await writeSpreadsheet(book,'xlsx'),excel=new ExcelJS.Workbook();await excel.xlsx.load(out.data);assert.equal(excel.worksheets[0].getCell('A1').value,12);assert.equal(excel.worksheets[0].getCell('A6').value,'123');
    book=await readSpreadsheet(out.data,'format.xlsx');for(let i=0;i<styles.length;i++){assert.equal(book.sheets[0].cells['A'+(i+1)].style.numberFormat,styles[i]);assert.equal(formatValue(i===5?'123':Number(raws[i]),book.sheets[0].cells['A'+(i+1)].style),formatValue(i===5?'123':Number(raws[i]),{numberFormat:styles[i]}));}
  }
});
