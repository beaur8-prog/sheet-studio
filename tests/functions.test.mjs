import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {spawnSync} from 'node:child_process';
import {createWorkbook,evaluateFormula,evaluateCell,shiftFormula,formulaForXlsx} from '../src/engine.mjs';
import {FUNCTION_CATALOG,FUNCTION_MAP,searchFunctions} from '../src/function-catalog.mjs';
import {formulaAssistance} from '../src/function-assist.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';

const fixture=JSON.parse(fs.readFileSync(new URL('./fixtures/excel-functions.json',import.meta.url),'utf8'));
const edges=JSON.parse(fs.readFileSync(new URL('./fixtures/excel-function-edges.json',import.meta.url),'utf8'));
function book(cells={}){
  const wb=createWorkbook();wb.sheets[0].cells=Object.fromEntries(Object.entries({A1:1,A2:2,A3:3,A4:4,A5:5,B1:'=SUM(A1:A5)',...cells}).map(([ref,raw])=>[ref,{raw:String(raw)}]));return wb;
}
const calc=(formula,cells={})=>evaluateFormula(book(cells),'sheet-1',formula);
const near=(actual,expected,label)=>{assert.equal(typeof actual,'number',label);assert.ok(Math.abs(actual-expected)<=1e-7*Math.max(1,Math.abs(expected)),`${label}: ${actual} ≠ ${expected}`);};
function complex(text){
  text=text.replace(/([+-])([ij])$/,'$11$2');
  const number='[+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][+-]?\\d+)?';
  const match=new RegExp(`^(${number})(${number})[ij]$`).exec(text);
  if(match)return [+match[1],+match[2]];
  if(/[ij]$/.test(text)){const value=text.slice(0,-1);return [0,value===''||value==='+'?1:value==='-'?-1:Number(value)];}
  return [Number(text),0];
}
for(const category of new Set(FUNCTION_CATALOG.map(f=>f.category))){
  test(`Excel oracle: ${category}`,()=>{
    const cases=fixture.cases.filter(c=>FUNCTION_MAP.get(c.name)?.category===category);assert.ok(cases.length);
    for(const item of cases){
      const actual=calc(item.formula),expected=item.expected;
      if(typeof expected==='number')near(actual,expected,item.formula);
      else if(item.name.startsWith('IM')&&typeof expected==='string'&&!expected.startsWith('#')){
        assert.equal(typeof actual,'string',item.formula);complex(actual).forEach((v,i)=>near(v,complex(expected)[i],item.formula));
      }else assert.equal(actual,expected,item.formula);
    }
  });
}

test('catalogue has unique implemented names, valid signatures and complete coverage',()=>{
  assert.equal(FUNCTION_CATALOG.length,379);assert.equal(FUNCTION_MAP.size,379);
  const covered=new Set(fixture.cases.map(c=>c.name));
  for(const fn of FUNCTION_CATALOG){
    assert.ok(fn.description&&fn.syntax&&fn.example.startsWith('='));assert.ok(fn.min<=fn.max);
    assert.ok(covered.has(fn.name)||['NOW','TODAY','RAND','RANDBETWEEN','XLOOKUP','XMATCH','TEXTBEFORE','TEXTAFTER'].includes(fn.name),fn.name);
    if(fn.min)assert.equal(calc(`=${fn.name}()`),'#VALUE!',fn.name);
  }
  assert.ok(searchFunctions('평균').some(f=>f.name==='AVERAGEIFS'));
  assert.deepEqual(searchFunctions('PMT','재무').map(f=>f.name),['CUMIPMT','IPMT','ISPMT','PMT','PPMT']);
  assert.equal(calc('=constructor(1)'),'#NAME?');assert.equal(calc('=utils.parseNumber(1)'),'#NAME?');
});

test('Excel oracle: 111 edge cases across types, errors, domains and old calendar rules',()=>{
  assert.equal(edges.cases.length,111);
  for(const item of edges.cases){const actual=calc(item.formula);if(typeof item.expected==='number')near(actual,item.expected,item.formula);else assert.equal(actual,item.expected,item.formula);}
});

test('XLSX future function prefixes preserve literals and do not grow on repeated export',()=>{
  const formula='=STDEV.S(A1:A5)+IFNA(NA(),0)&"TEXTJOIN("';
  const output='=_xlfn.STDEV.S(A1:A5)+_xlfn.IFNA(NA(),0)&"TEXTJOIN("';
  assert.equal(formulaForXlsx(formula),output);assert.equal(formulaForXlsx(output),output);
  assert.equal(formulaForXlsx('=ERROR.TYPE(NA())'),'=ERROR.TYPE(NA())');
});

test('array constants, expression arrays, prefixed names and literal operators compose',()=>{
  assert.equal(calc('=SUMPRODUCT((A1:A5>2)*(A1:A5<5),A1:A5)'),7);
  assert.equal(calc('=SUM({1,2;3,4})'),10);
  assert.equal(calc('=SUM({1,2;3,4}*{10;20})'),170);
  assert.equal(calc('=SUMPRODUCT({1,2},{3;4})'),'#VALUE!');
  assert.equal(calc('=SUM({1,2;3})'),'#VALUE!');
  assert.equal(calc('=SUM({A1,2})'),'#VALUE!');
  assert.equal(calc('=_xlfn.STDEV.S(A1:A5)'),Math.sqrt(2.5));
  assert.equal(calc('=TEXTJOIN(",",TRUE,"-","+","=","(",")")'), '-,+,=,(,)');
  assert.equal(shiftFormula('=SUMPRODUCT(A1:A5,{1;2;3;4;5})',1,1),'=SUMPRODUCT(B2:B6,{1;2;3;4;5})');
});

test('reference functions retain coordinates, sheet identity and dependency cycles',()=>{
  const wb=book({D10:'=ROW()+COLUMN()',A8:'=ROW(A8)',A9:'=INDIRECT("A9")'});
  wb.sheets.push({id:'other',name:"다른 '시트'",cells:{C3:{raw:'42'}},charts:[]});
  assert.equal(evaluateCell(wb,'sheet-1','D10'),14);assert.equal(evaluateCell(wb,'sheet-1','A8'),8);
  assert.equal(evaluateCell(wb,'sheet-1','A9'),'#CYCLE!');
  assert.equal(evaluateFormula(wb,'sheet-1',"=INDIRECT(\"'다른 ''시트'''!C3\")"),42);
  assert.equal(calc('=SUM(INDEX(A1:A5,0))'),15);
  assert.equal(calc('=SUM(OFFSET(A1,1,0,3,1))'),9);
  assert.equal(calc('=ROW(OFFSET(A1,2,0))'),3);
  assert.equal(calc('=ISREF(INDEX(A1:A5,2))'),true);
  assert.equal(calc('=ISREF(2)'),false);
  assert.equal(calc('=INDIRECT("SUM(A1:A5)")'),'#REF!');
  assert.equal(calc('=OFFSET(A1,-1,0)'),'#REF!');
  assert.equal(calc('=SUM(OFFSET(A1,0,0,100001,1))'),'#NUM!');
  assert.equal(calc('=SUBTOTAL(9,A1:A6)',{A6:'=SUBTOTAL(9,A1:A5)'}),15);
  assert.equal(calc('=SUM(A1:A6)',{A6:'=SUBTOTAL(9,A1:A5)'}),30);
});

test('criteria pairs, blank handling, numeric text and errors keep Excel semantics',()=>{
  const cells={A1:1,A2:2,A3:"'9",A4:'TRUE',A5:'',B1:'x',B2:'x',B3:'y',B4:'y',B5:'=""'};
  assert.equal(calc('=AVERAGEIFS(A1:A5,B1:B5,"x")',cells),1.5);
  assert.equal(calc('=MINIFS(A1:A5,B1:B5,"missing")',cells),0);
  assert.equal(calc('=COUNTIFS(B1:B5,"x",A1:A5,">1")',cells),1);
  assert.equal(calc('=SUMIFS(A1:A5,B1:B2,"x")',cells),'#VALUE!');
  assert.equal(calc('=SUMIFS(A1:A5,B1:B5)',cells),'#VALUE!');
  assert.equal(calc('=COUNTBLANK(A1:B5)',cells),2);
  assert.equal(calc('=PRODUCT(A1:A5)',cells),2);
  assert.equal(calc('=MEDIAN(A1:A5,"7",TRUE)',cells),1.5);
  assert.equal(calc('=VAR.S(A1:A5)',cells),0.5);
  assert.equal(calc('=VAR.S(A5)',cells),'#DIV/0!');
  assert.equal(calc('=STDEV.S(A1:A5)',{A3:'=1/0'}),'#DIV/0!');
  assert.equal(calc('=SUMIFS(A1:A5,B1:B5,"x")',{...cells,A3:'=1/0'}),3);
});

test('logical and information functions inspect errors without eager unused branches',()=>{
  assert.equal(calc('=IFNA(NA(),7)'),7);assert.equal(calc('=IFNA(1/0,7)'),'#DIV/0!');
  assert.equal(calc('=IFS(FALSE,1/0,TRUE,7)'),7);
  assert.equal(calc('=SWITCH(2,1,1/0,2,7,UNKNOWN())'),7);
  assert.equal(calc('=CHOOSE(2,1/0,7,UNKNOWN())'),7);
  assert.equal(calc('=ISNUMBER("42")'),false);assert.equal(calc('=ISBLANK(A7)'),true);
  assert.equal(calc('=ISBLANK(A7)',{A7:'=""'}),false);
  assert.equal(calc('=ISERROR(1/0)'),true);assert.equal(calc('=ERROR.TYPE(NA())'),7);
  assert.equal(calc('=TYPE({1,2})'),64);assert.equal(calc('=N("42")'),0);
});

test('modern lookups and text extraction cover modes, optional arguments and lazy defaults',()=>{
  assert.equal(calc('=XLOOKUP(2,{1,2,3},{10,20,30},1/0)'),20);
  assert.equal(calc('=XMATCH(2,{2,1,2},0,-1)'),3);
  assert.equal(calc('=XMATCH("b*",{"aa","BB","cc"},2)'),2);
  assert.equal(calc('=XMATCH(2.5,{1,2,3},-1)'),2);
  assert.equal(calc('=TEXTAFTER("a-b-c","-",-1)'), 'c');
  assert.equal(calc('=TEXTBEFORE("a-b-c","-",-1)'), 'a-b');
  assert.equal(calc('=TEXTBEFORE("Marcus Aurelius"," ",,,1)'), 'Marcus');
  assert.equal(calc('=TEXTBEFORE("Socrates"," ",,,1)'), 'Socrates');
  assert.equal(calc('=TEXTAFTER("a-b","-",1,0,0,1/0)'), 'b');
  assert.equal(calc('=TEXTAFTER("ab","-",1,0,0,"none")'),'none');
  assert.equal(calc('=TEXTBEFORE("ab","-",0)'),'#VALUE!');
  assert.equal(calc('=TEXTBEFORE("ab","-",3)'),'#VALUE!');
  assert.equal(calc('=UNICODE(UNICHAR(128512))'),128512);
});

test('calendar functions cover leap days, weekends, holidays, negative offsets and time',()=>{
  assert.equal(calc('=EDATE(DATE(2024,1,31),1)'),calc('=DATE(2024,2,29)'));
  assert.equal(calc('=EOMONTH(DATE(1900,2,1),0)'),59);
  assert.equal(calc('=NETWORKDAYS(DATE(2024,1,7),DATE(2024,1,1))'),-5);
  assert.equal(calc('=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),"0000011",{45292,45292})'),4);
  assert.equal(calc('=WORKDAY(DATE(2024,1,8),-1)'),calc('=DATE(2024,1,5)'));
  assert.equal(calc('=WORKDAY.INTL(DATE(2024,1,1),1,"1111111")'),'#NUM!');
  assert.equal(calc('=DATEVALUE("2023-02-29")'),'#VALUE!');
  assert.equal(calc('=TIMEVALUE("12:30 PM")'),12.5/24);
  assert.equal(calc('=ISOWEEKNUM(DATE(2021,1,1))'),53);
  for(const zone of ['UTC','America/New_York','Asia/Seoul']){
    const code=`import {createWorkbook,evaluateFormula} from './src/engine.mjs';console.log(evaluateFormula(createWorkbook(),'sheet-1','=EDATE(DATE(2024,3,31),-1)+NETWORKDAYS(DATE(2024,3,1),DATE(2024,3,31))'));`;
    const result=spawnSync(process.execPath,['--input-type=module','-e',code],{cwd:new URL('..',import.meta.url),env:{...process.env,TZ:zone},encoding:'utf8'});
    assert.equal(result.status,0,result.stderr);assert.equal(Number(result.stdout.trim()),45372,zone);
  }
});

test('math and financial adapters preserve large integers, domains and bounded iterations',()=>{
  assert.equal(calc('=BITOR(1099511627776,3)'),1099511627779);
  assert.equal(calc('=BITLSHIFT(1,47)'),140737488355328);
  assert.equal(calc('=BITLSHIFT(1,48)'),'#NUM!');
  assert.equal(calc('=COMBIN(1000000,1)'),1000000);
  assert.equal(calc('=QUOTIENT(1,0)'),'#DIV/0!');
  assert.equal(calc('=FACT(1000000)'),'#NUM!');
  assert.equal(calc('=REPT("x",1000000000)'),'#VALUE!');
  assert.equal(calc('=BINOM.DIST.RANGE(10,0.5,1,1000000000)'),'#NUM!');
  assert.equal(calc('=BINOM.INV(100000,0.5,0.5)'),50000);
  assert.equal(calc('=EXP(1E999)'),'#NUM!');
  assert.equal(calc('=XIRR({-1,1},{43831,43831})'),0.1);
  assert.equal(calc('=XIRR({1,2},{43831,44197})'),'#NUM!');
  assert.equal(calc('=XIRR({-100,1},{43831,44197})')<0,true);
  assert.equal(calc('=XNPV(0.1,{-100,110},{43831})'),'#NUM!');
  near(calc('=T.DIST(-1,10,TRUE)'),1-calc('=T.DIST(1,10,TRUE)'),'negative t distribution');
  for(const f of ['RAND()','RANDBETWEEN(2,3)','NOW()','TODAY()'])assert.equal(typeof calc('='+f),'number');
  const random=calc('=RAND()'),between=calc('=RANDBETWEEN(2,3)'),today=calc('=TODAY()'),now=calc('=NOW()');
  assert.ok(random>=0&&random<1);assert.ok(between===2||between===3);assert.ok(Number.isInteger(today)&&now>=today&&now<today+1);
});

test('function assistant understands nesting, quotes, arrays and prefixed names',()=>{
  assert.deepEqual(formulaAssistance('="SUMI').suggestions,[]);
  assert.equal(formulaAssistance('=SUMI').suggestions[0].name,'SUMIF');
  assert.deepEqual(formulaAssistance("='SUMI").suggestions,[]);
  assert.deepEqual(formulaAssistance('=Sheet1!SU').suggestions,[]);
  assert.equal(formulaAssistance('=IF(SUM(A1:A5)>1,ROUND(1.5,').active.name,'ROUND');
  assert.equal(formulaAssistance('=IF(SUM(A1:A5)>1,ROUND(1.5,').active.argument,1);
  assert.equal(formulaAssistance('=SUMPRODUCT({1,2;3,4},').active.argument,1);
  assert.equal(formulaAssistance('=_xlfn.NORM.DIST(1,0,1,').active.name,'NORM.DIST');
});

test('extended formulas and calculated caches survive an XLSX roundtrip',async()=>{
  const wb=book({C1:'=SUMIFS(A1:A5,A1:A5,">2")',C2:'=PMT(0.01,12,1000)',C3:'=_xlfn.STDEV.S(A1:A5)',C4:'=SUMPRODUCT((A1:A5>2)*A1:A5)',C5:'=TEXTJOIN(",",TRUE,"a","b")'});
  const data=await writeSpreadsheet(wb,'xlsx');
  const restored=await readSpreadsheet(data.data,'functions.xlsx');
  for(let i=1;i<=5;i++)assert.equal(evaluateCell(restored,restored.sheets[0].id,'C'+i),evaluateCell(wb,'sheet-1','C'+i));
});
