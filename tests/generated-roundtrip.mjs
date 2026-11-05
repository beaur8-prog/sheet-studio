import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {performance} from 'node:perf_hooks';
import path from 'node:path';
import {createWorkbook,createEvaluator,address,parseAddress,formatValue} from '../src/engine.mjs';
import {calculationDisplay} from '../src/calculation-display.mjs';
import {worksheetCellStyle} from '../src/cell-style.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';
import {chartData} from '../src/chart-model.mjs';
import {createHistory} from '../src/history.mjs';
import {replaceWorkbookPlan,replacementPatch,applyReplacementPatch} from '../src/find-replace.mjs';
// Optional generated invariant audit. It is deliberately separate from npm test.
// Only synthetic data and the requested output directory are used.
const output=path.resolve(process.argv[2]||'test-results/generated-roundtrip');await fs.mkdir(output,{recursive:true});
let state=0x51ee738;const random=()=>{state^=state<<13;state^=state>>>17;state^=state<<5;return (state>>>0)/4294967296;},pick=a=>a[Math.floor(random()*a.length)];
const raws=['0','-12.5','2.34567',"'001",'TRUE','FALSE','가나다\n둘째 줄','<태그>& 기호',"'=SUM(A1)","'#N/A",'=#N/A','="#N/A"','12.5%','', '60','45235'];
const start=performance.now(),results=[];
const close=(a,b)=>typeof a==='number'&&typeof b==='number'?Math.abs(a-b)<=1e-10*Math.max(1,Math.abs(a),Math.abs(b)):a===b;
function snapshot(book,refs){
 return book.sheets.map((sheet,i)=>{const e=createEvaluator(book,sheet.id);return {name:sheet.name,cells:refs[i].map(ref=>{const result=calculationDisplay(e.result(ref)),p=parseAddress(ref),style=worksheetCellStyle(book,sheet,p.r,p.c);return {ref,value:result.value,error:result.error,display:formatValue(result.value,style,{date1904:book.date1904,error:result.error}),note:sheet.cells[ref]?.note||''};}),charts:sheet.charts.map(chart=>chartData(chart,sheet,book).slice(0,chart.type==='pie'?1:128).map(s=>({name:s.name,values:s.values,labels:s.labels,xs:s.xs})))};});
}
for(let index=0;index<120;index++){
 const book=createWorkbook();book.name='Roundtrip '+index;book.date1904=!!(index%2);book.sheets[0].name="매출 '24";const second=createWorkbook().sheets[0];second.id=crypto.randomUUID();second.name='자료2';book.sheets.push(second);
 const refs=[];for(const [si,s] of book.sheets.entries()){
  for(let r=0;r<7;r++)for(let c=0;c<4;c++){if(random()<.15)continue;const raw=pick(raws),style={fontFamily:'Arial',fontSize:11,format:pick(['general','number','percent','date']),decimals:Math.floor(random()*4),color:pick(['#123456','#654321','#000000']),background:pick(['#ffffff','#ffeeaa'])};if(style.format==='date'){style.numberFormat='yyyy-mm-dd';delete style.decimals;}s.cells[address(r,c)]={raw,style,...(random()<.1?{note:'메모 '+index+' '+r}: {})};}
  s.cells.A1={raw:String(index+1),style:{format:'number',decimals:2,fontFamily:'Arial'}};s.cells.J1={raw:'marker',style:{bold:true}};s.cells.J2={raw:'=A1*2',style:{format:'number',decimals:3}};s.cells.J3={raw:si?"='매출 ''24'!A1+1":'=J2+1',style:{}};s.cells.D11={raw:'0'};s.cells.D12={raw:'1'};s.cells.E10={raw:'빈칸'};s.cells.E11={raw:'숫자0'};s.cells.E12={raw:'숫자1'};s.cells.F10={raw:'=XLOOKUP(,D10:D12,E10:E12)',style:{}};s.cells.C10={raw:'=MISSINGFUNC(1)',cachedFormula:'=MISSINGFUNC(1)',cachedValue:42,xlsxFormula:{raw:'=MISSINGFUNC(1)',formula:'MISSINGFUNC(1)'},style:{}};
  s.charts=[{id:crypto.randomUUID(),type:pick(['bar','line','pie','doughnut','radar','scatter']),range:'A1:D7',firstRowAsHeader:false,firstColumnAsLabels:false,seriesBy:'columns',blankCells:pick(['gap','zero','span']),title:'한글 '+index,width:600,height:360}];refs.push(Object.keys(s.cells));
 }
 const history=createHistory(book),original=JSON.stringify(history.raw),plan=replaceWorkbookPlan(history.raw,{sheetId:book.sheets[0].id,query:'marker',replacement:'changed',scope:'workbook'}),patch=replacementPatch(history.raw,plan.book);history.begin();applyReplacementPatch(history.book,patch);history.commit();assert.equal(history.book.sheets[0].cells.J1.raw,'changed');history.undo();assert.equal(JSON.stringify(history.raw),original);history.undo(true);let current=history.raw,bytes;
 const expected=snapshot(current,refs);
 try{
  for(let round=0;round<2;round++){
   bytes=(await writeSpreadsheet(current)).data;current=await readSpreadsheet(bytes,'random.xlsx');const actual=snapshot(current,refs);
   for(let si=0;si<expected.length;si++){
    assert.equal(actual[si].name,expected[si].name);for(let n=0;n<expected[si].cells.length;n++){const a=actual[si].cells[n],e=expected[si].cells[n];assert.ok(close(a.value,e.value),`seed ${index} sheet ${si} ${e.ref}: ${JSON.stringify(a.value)} != ${JSON.stringify(e.value)}`);assert.equal(a.error,e.error,`seed ${index} ${e.ref} error`);assert.equal(a.display,e.display,`seed ${index} ${e.ref} display`);assert.equal(a.note,e.note,`seed ${index} ${e.ref} note`);}assert.deepEqual(actual[si].charts,expected[si].charts,`seed ${index} sheet ${si} chart`);
   }
  }
  results.push({index,sheets:2,roundtrips:2});
 }catch(error){await fs.writeFile(path.join(output,'property-failure-input.sheet.json'),JSON.stringify(book));if(bytes)await fs.writeFile(path.join(output,'property-failure.xlsx'),bytes);await fs.writeFile(path.join(output,'property-failure.json'),JSON.stringify({index,error:String(error),expected,actual:snapshot(current,refs)},null,2));throw error;}
}
const report={seed:'0x51ee738',cases:results.length,sheets:results.length*2,xlsxRoundtrips:results.length*2,ms:performance.now()-start,scope:'Seeded synthetic typed values, formulas, dates, formatting, notes, chart layouts/blank policies, atomic replace undo/redo and two XLSX roundtrips. Self-consistency, not new Excel execution.'};await fs.writeFile(path.join(output,'property-report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report));
