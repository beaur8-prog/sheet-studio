import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator,evaluateCell,parseWorkbook,serializeWorkbook} from '../src/engine.mjs';
import {paneRects,visibleAxis,axisPick} from '../src/worksheet-view.mjs';
import {pastePlan,scalarRaw} from '../src/cell-paste.mjs';
import {createConditionalEvaluator,shiftWorksheetFeatures,conditionalFormula} from '../src/conditional.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {attachNativeSource,buildNativePlan,nativeSourceChanged} from '../src/native-plan.mjs';
import {createRequire} from 'node:module';
import {worksheetFixture} from './fixtures/worksheet.mjs';
const vba=createRequire(import.meta.url)('../electron/vba.cjs');
test('frozen panes keep corner, rows and columns visible with scroll and hidden dimensions',()=>{
  const axis=[0,40,100,100,180,240,300];
  assert.deepEqual(visibleAxis(axis,100,200,1),[0,3,4,5]);assert.equal(axisPick(axis,20,100,1),0);assert.equal(axisPick(axis,60,100,1),3);
  const rects=paneRects({r1:0,c1:0,r2:5,c2:5},{cols:axis,rows:axis},{cols:1,rows:1},{left:100,top:100},200,200);
  assert.equal(rects.length,4);assert.deepEqual(rects.map(r=>[r.x,r.y,r.w,r.h]),[[0,0,40,40],[40,0,160,40],[0,40,40,160],[40,40,160,160]]);
});
test('paste special preserves text types, target style and relative formulas through transpose/repeat',()=>{
  const book=createWorkbook(),s=book.sheets[0];s.cells.D4={raw:'99',style:{bold:true},note:'keep'};
  const clip={r:0,c:0,data:[[{raw:'=B1',style:{background:'#ff0000'}},{raw:"'001"}],[{raw:'3'},{raw:''}]],values:[[8,'001'],[3,'']]};
  let p=pastePlan(s,clip,{r:3,c:3},{mode:'values',skipBlanks:true});assert.equal(p.updates.length,3);assert.equal(p.updates[0].cell.raw,'8');assert.equal(p.updates[0].cell.style.bold,true);assert.equal(p.updates[1].cell.raw,"'001");
  p=pastePlan(s,clip,{r:3,c:3},{mode:'formulas',transpose:true});assert.equal(p.updates[0].cell.raw,'=D5');assert.equal(p.updates[1].cell.raw,'3');assert.equal(p.updates[0].cell.note,'keep');
  p=pastePlan(s,clip,{r:3,c:3},{mode:'formats'});assert.equal(p.updates[0].cell.raw,'99');assert.equal(p.updates[0].cell.style.background,'#ff0000');
  p=pastePlan(s,clip,{r:3,c:3},{selection:{r1:3,c1:3,r2:6,c2:6}});assert.equal(p.updates.length,16);assert.equal(p.updates.find(v=>v.key==='F6').cell.raw,'=G6');
  assert.equal(scalarRaw('=literal'),"'=literal");assert.equal(scalarRaw(true),'TRUE');
  assert.throws(()=>pastePlan(s,clip,{r:1048575,c:0}),/초과/);assert.throws(()=>pastePlan(s,{...clip,cut:true},{r:0,c:0},{transpose:true}),/복사/);
});
test('arithmetic paste is atomic and catches invalid operands before edits',()=>{
  const s=createWorkbook().sheets[0];s.cells.A1={raw:'10'};const clip={r:0,c:0,data:[[{raw:'2'}]],values:[[2]]};
  assert.equal(pastePlan(s,clip,{r:0,c:0},{operation:'multiply'}).updates[0].cell.raw,'20');assert.equal(s.cells.A1.raw,'10');
  assert.equal(pastePlan(s,{...clip,values:[[0]]},{r:0,c:0},{operation:'divide'}).updates[0].cell.raw,'#DIV/0!');
  assert.throws(()=>pastePlan(s,{...clip,values:[['bad']]},{r:0,c:0},{operation:'add'}),/숫자/);
});
test('conditional rules recalculate dependencies, preserve base styles and honor priority',()=>{
  const b=worksheetFixture(),s=b.sheets[0];let e=createEvaluator(b,s.id),c=createConditionalEvaluator(b,s,e);
  assert.equal(c(3,1,e('B4')).background,'#ffc7ce');assert.equal(c(3,2,e('C4')).color,'#9c0006');assert.equal(c(2,3,e('D3')).background,'#c6efce');assert.equal(c(1,0,e('A2')).background,'#ffeb9c');assert.deepEqual(c(2,0,e('A3')),{});
  s.cells.B3.raw='1';s.cells.B4.raw='5';e=createEvaluator(b,s.id);c=createConditionalEvaluator(b,s,e);assert.deepEqual(c(2,3,e('D3')),{});assert.deepEqual(c(3,1,e('B4')),{});assert.deepEqual(s.cells.B4.style,{});
  s.conditionalRules=[{id:'a',range:'B2:B6',type:'expression',formulae:['TRUE'],style:{color:'#ff0000'},stopIfTrue:true},{id:'b',range:'B2:B6',type:'expression',formulae:['TRUE'],style:{background:'#00ff00',color:'#000000'}}];assert.deepEqual(createConditionalEvaluator(b,s)(1,1,2),{color:'#ff0000'});
  s.conditionalRules[0].stopIfTrue=false;assert.deepEqual(createConditionalEvaluator(b,s)(1,1,2),{color:'#ff0000',background:'#00ff00'});
});
test('worksheet features survive JSON/XLSX and native plans without flattening conditional colors',async()=>{
  const b=worksheetFixture(),s=b.sheets[0],json=parseWorkbook(serializeWorkbook(b));assert.deepEqual(json.sheets[0].freeze,s.freeze);
  const out=await writeSpreadsheet(b),round=await readSpreadsheet(out.data,'features.xlsx'),r=round.sheets[0];assert.deepEqual(r.freeze,s.freeze);assert.equal(r.columnWidths[5],0);assert.equal(r.rowHeights[8],0);assert.equal(r.conditionalRules.length,4);assert.equal(r.conditionalRules[0].stopIfTrue,true);assert.ok(!r.cells.B4.style.background);
  const e=createEvaluator(round,r.id),c=createConditionalEvaluator(round,r,e);assert.equal(c(3,1,e('B4')).background,'#ffc7ce');assert.equal(c(1,0,e('A2')).background,'#ffeb9c');
  attachNativeSource(b,out.data,'test.xlsm');assert.equal(nativeSourceChanged(b),false);s.freeze={rows:2,cols:2};s.conditionalRules[0].formulae=['1'];const plan=buildNativePlan(b);assert.equal(plan.sheets[0].conditionalFormats.length,4);assert.deepEqual(plan.sheets[0].freeze,{rows:2,cols:2});vba.validatePayload({mode:'save',source:b.nativeSource,plan});
  plan.sheets[0].conditionalFormats[0].style.background='injection';assert.throws(()=>vba.validatePayload({mode:'save',source:b.nativeSource,plan}),/색상/);
});
test('structural edits shift rule ranges and freeze boundary, malformed persisted rules fail early',()=>{
  const b=worksheetFixture(),s=b.sheets[0];shiftWorksheetFeatures(s,'r',0,1);assert.equal(s.freeze.rows,2);assert.equal(s.conditionalRules[0].range,'B3:B7');assert.equal(s.conditionalRules[3].formulae[0],'$B3<0');
  s.freeze.rows=-1;assert.throws(()=>serializeWorkbook(b),/틀 고정/);s.freeze.rows=2;s.conditionalRules[0].style.color='red;bad';assert.throws(()=>serializeWorkbook(b),/색상/);
  assert.equal(conditionalFormula({range:'A2:A6',type:'duplicateValues'}),'AND(NOT(ISBLANK(A2)),COUNTIF($A$2:$A$6,A2)>1)');
});
test('top/bottom ties, percentages, averages, wildcard text and blank duplicate rules remain dynamic',async()=>{
  const b=createWorkbook(),s=b.sheets[0];[10,20,20,40].forEach((n,i)=>s.cells['A'+(i+1)]={raw:String(n)});s.cells.B1={raw:'Alpha'};s.cells.B2={raw:'Beta'};
  const rule={id:'r',range:'A1:A5',type:'top10',rank:50,percent:true,formulae:[],style:{background:'#ff0000'}};s.conditionalRules=[rule];
  const selected=()=>{const e=createEvaluator(b,s.id),c=createConditionalEvaluator(b,s,e);return [0,1,2,3,4].filter(r=>c(r,0,e('A'+(r+1))).background);};
  assert.deepEqual(selected(),[1,2,3]);rule.bottom=true;assert.deepEqual(selected(),[0,1,2]);rule.type='aboveAverage';rule.aboveAverage=false;assert.deepEqual(selected(),[0,1,2]);
  s.conditionalRules=[{...rule,type:'containsText',range:'B1:B3',text:'a*ha'}];const c=createConditionalEvaluator(b,s);assert.ok(c(0,1,'Alpha').background);assert.deepEqual(c(1,1,'Beta'),{});assert.deepEqual(c(2,1,''),{});
  s.conditionalRules=[{...rule,type:'duplicateValues'}];assert.deepEqual(selected(),[1,2]);const round=await readSpreadsheet((await writeSpreadsheet(b)).data,'rules.xlsx'),e=createEvaluator(round,round.sheets[0].id),cond=createConditionalEvaluator(round,round.sheets[0],e);assert.deepEqual(cond(4,0,''),{});assert.ok(cond(1,0,20).background);
});
