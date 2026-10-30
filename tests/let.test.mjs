import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {createWorkbook,createEvaluator,formulaForXlsx,shiftFormula,transformFormulaReferences} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
const setup=()=>{const b=createWorkbook();return {b,s:b.sheets[0]};},value=(b,s,raw)=>{s.cells.Z1={raw};return createEvaluator(b,s.id)('Z1');};
test('LET supports sequential bindings, Unicode names, nested shadowing and imported parameter prefixes',()=>{
 const {b,s}=setup();assert.equal(value(b,s,'=LET(x,5,SUM(x,1))'),6);assert.equal(value(b,s,'=LET(x,5,y,x*2,x+y)'),15);assert.equal(value(b,s,'=LET(x,5,LET(x,x+1,x)+x)'),11);assert.equal(value(b,s,'=_xlfn.LET(_xlpm.x,2,_xlpm.y,_xlpm.x+1,_xlpm.y*3)'),9);assert.equal(value(b,s,'=LET(매출,40,세율,0.1,매출*(1+세율))'),44);assert.equal(value(b,s,'=LET(label,"001",label&"-"&label)'),'001-001');
});
test('LET local scopes cannot leak into cells, workbook names or later formulas',()=>{
 const {b,s}=setup();b.chartNames=[{name:'x',formula:'100'},{name:'outer',formula:'x+1'}];s.cells.A1={raw:'=x'};assert.equal(value(b,s,'=LET(x,2,A1+x)'),102);assert.equal(value(b,s,'=LET(x,2,outer+x)'),103);assert.equal(value(b,s,'=LET(x,2,Sheet1!x+x)'),102);assert.equal(value(b,s,'=LET(x,2,x)+x'),102);assert.equal(value(b,s,'=LET(x,x+1,x)'),101);
 delete b.chartNames;assert.equal(value(b,s,'=LET(x,2,A1)'),'#NAME?');assert.equal(value(b,s,'=LET(x,y,y,2,x)'),'#NAME?');assert.equal(value(b,s,'=x'),'#NAME?');
});
test('LET reuses random calculations across dependency scheduling and evaluates only required values',()=>{
 const {b,s}=setup();s.cells.A1={raw:'=2+3'};s.cells.A2={raw:'=LET(x,RAND(),y,A1,x-x+y)'};s.cells.A3={raw:'=LET(x,NA(),IF(FALSE,x,7))'};s.cells.A4={raw:'=LET(x,1/0,8)'};
 let count=0;const previous=Math.random;Math.random=()=>{count++;return .25;};try{const e=createEvaluator(b,s.id);assert.equal(e('A2'),5);assert.equal(count,1);assert.equal(e('A3'),7);assert.equal(e('A4'),8);assert.equal(e('A2'),5);assert.equal(count,1);}finally{Math.random=previous;}
});
test('LET range variables retain reference metadata, sparse whole-column calculations, array results and copy/structure references',()=>{
 const {b,s}=setup();s.cells.A1={raw:'1'};s.cells.A2={raw:'2'};s.cells.A1048576={raw:'3'};assert.equal(value(b,s,'=LET(data,A:A,SUM(data)+COUNTBLANK(data))'),1048579);assert.equal(value(b,s,'=LET(data,A2:A3,ROW(data))'),2);assert.equal(value(b,s,'=LET(data,A1:A2,INDEX(data,2,1))'),2);assert.equal(value(b,s,'=LET(data,A1:A2,ISREF(data))'),true);assert.equal(value(b,s,'=LET(data,{1;2;3},SUM(data*2))'),12);
 s.cells.C1={raw:'=LET(data,A1:A2,data*3)',xlsxFormula:{raw:'=LET(data,A1:A2,data*3)',formula:'LET(data,A1:A2,data*3)',shareType:'array',ref:'C1:C2'}};assert.equal(createEvaluator(b,s.id)('C2'),6);
 assert.equal(shiftFormula('=LET(data,A1:A2,SUM(data))',2,1),'=LET(data,B3:B4,SUM(data))');assert.equal(transformFormulaReferences('=LET(data,A1:A2,SUM(data))',{axis:'r',index:0,delta:2}),'=LET(data,A3:A4,SUM(data))');
});
test('LET validates binding names, argument pairs, duplicates and the 126-pair limit',()=>{
 const {b,s}=setup();for(const raw of ['=LET()','=LET(x,1)','=LET(x,1,y,2)','=LET(A1,1,2)','=LET(c,1,c)','=LET(r,1,r)','=LET("x",1,2)','=LET(x,1,X,2,x)'])assert.equal(value(b,s,raw),'#VALUE!',raw);
 const make=count=>'=LET('+Array.from({length:count},(_,i)=>'v_'+i+','+(i?'v_'+(i-1)+'+1':'1')).join(',')+',v_'+(count-1)+')';assert.equal(value(b,s,make(126)),126);assert.equal(value(b,s,make(127)),'#VALUE!');assert.equal(value(b,s,'=IFERROR(LET(x,1,x/0),"caught")'),'caught');
});
test('LET uses the OOXML future-function prefix and current typed caches through save/reopen',async()=>{
 let {b,s}=setup();s.cells.A1={raw:'4'};s.cells.B1={raw:'=LET(x,A1+1,x*2)'};s.cells.B2={raw:'=_xlfn.LET(_xlpm.t,"001",_xlpm.t)'};
 for(let i=0;i<3;i++){const out=await writeSpreadsheet(b),zip=await JSZip.loadAsync(out.data);assert.match(await zip.file('xl/worksheets/sheet1.xml').async('string'),/_xlfn.LET/);const check=new ExcelJS.Workbook();await check.xlsx.load(out.data);assert.equal(check.worksheets[0].getCell('B1').result,10);assert.equal(check.worksheets[0].getCell('B2').result,'001');b=await readSpreadsheet(out.data,'let.xlsx');s=b.sheets[0];assert.equal(createEvaluator(b,s.id)('B1'),10);assert.equal(createEvaluator(b,s.id)('B2'),'001');}
 assert.equal(formulaForXlsx('=LET(x,2,x+1)'),'=_xlfn.LET(x,2,x+1)');
});
