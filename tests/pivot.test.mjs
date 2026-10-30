import test from 'node:test';
import assert from 'node:assert/strict';
import {address,parseWorkbook,serializeWorkbook,createWorkbook} from '../src/engine.mjs';
import {sourceData,buildPivot,applyPivot,shiftPivotSources} from '../src/pivot.mjs';

import {fixture,spec} from './fixtures/pivot.mjs';

test('pivot cross tab, grand totals and formulas use original records',()=>{
  const book=fixture();book.sheets[0].cells.C3.raw='=C2*2';
  assert.deepEqual(buildPivot(book,spec()).matrix,[['지역','A · 합계: 매출','B · 합계: 매출','총합계 · 합계: 매출'],['동부',30,30,60],['서부',100,0,100],['총합계',130,30,160]]);
  const s=spec();s.values[0].aggregate='average';const m=buildPivot(book,s).matrix;
  assert.deepEqual(m[1],['동부',15,30,20]);assert.equal(m.at(-1).at(-1),40);assert.equal(m[2][2],'');
});
test('multiple measures, row fields, sort and typed filters',()=>{
  const s=spec();s.rows=['지역','상품'];s.columns=[];s.values=[{field:'매출',aggregate:'count'},{field:'수량',aggregate:'max'},{field:'수량',aggregate:'min'}];s.filters=[{field:'상품',values:['A']}];s.sort='desc';
  assert.deepEqual(buildPivot(fixture(),s).matrix,[['지역','상품','개수: 매출','최댓값: 수량','최솟값: 수량'],['서부','A',1,4,4],['동부','A',2,2,1],['총합계','',3,4,1]]);
  const b=fixture();b.sheets[0].cells.C2.raw="'10";s.rows=[];s.filters=[];s.values=[{field:'매출',aggregate:'count'},{field:'매출',aggregate:'countNumbers'}];assert.deepEqual(buildPivot(b,s).matrix[1],['전체',4,3]);
});
test('blank groups, no matches, numeric text, boolean and formula literal remain distinct',()=>{
  const b=fixture();b.sheets[0].cells.A2.raw='';b.sheets[0].cells.A3.raw='1';b.sheets[0].cells.A4.raw="'1";b.sheets[0].cells.A5.raw="'=SUM(A1)";b.sheets[0].cells.A6.raw='FALSE';
  const s=spec();s.columns=[];s.filters=[{field:'지역',values:[1]}];assert.equal(buildPivot(b,s).matched,1);
  s.filters=[];const out=buildPivot(b,s);assert.equal(out.rows.length,5);assert.ok(out.matrix.some(r=>r[0]==='(비어 있음)'));
  const target={id:'pivot',name:'피벗'};applyPivot(b,target,s);assert.ok(Object.values(target.cells).some(c=>c.raw==="'=SUM(A1)"));
  s.filters=[{field:'상품',values:[]}];assert.equal(buildPivot(b,s).matched,0);assert.equal(buildPivot(b,s).matrix.at(-1).at(-1),0);
});
test('invalid source, schema and calculation errors fail without modifying result',()=>{
  const b=fixture(),s=spec(),target={id:'pivot',name:'피벗'};applyPivot(b,target,s);const before=structuredClone(target);
  b.sheets[0].cells.C2.raw='=1/0';assert.throws(()=>applyPivot(b,target,s),/계산 오류/);assert.deepEqual(target,before);
  assert.throws(()=>sourceData(b,'source','A1:XFD1048576'),/100,000/);assert.throws(()=>sourceData(b,'source','A1'),/머리글/);
  b.sheets[0].cells.B1.raw='지역';assert.throws(()=>buildPivot(b,s),/고유한/);
  assert.throws(()=>buildPivot(fixture(),{...s,rows:['없음']}),/필드/);
  assert.throws(()=>buildPivot(fixture(),{...s,columns:['지역']}),/중복/);
});
test('pivot definitions survive JSON, refresh, source rename and source structural edits',()=>{
  const b=fixture(),s=spec(),target={id:'pivot',name:'피벗',charts:[]};applyPivot(b,target,s);b.sheets.push(target);
  const read=parseWorkbook(serializeWorkbook(b));assert.deepEqual(read.sheets[1].pivot,s);
  read.sheets[0].name='이름 변경';read.sheets[0].cells.C2.raw='40';applyPivot(read,read.sheets[1],s);assert.equal(read.sheets[1].cells.D4.raw,'190');
  shiftPivotSources(read,'source','r',3,1);assert.equal(read.sheets[1].pivot.sourceRange,'A1:D7');
  read.sheets.shift();assert.throws(()=>applyPivot(read,read.sheets[0],read.sheets[0].pivot),/원본 시트/);
  assert.throws(()=>parseWorkbook(JSON.stringify({...b,sheets:[b.sheets[0],{...target,pivot:{...s,values:[{field:'매출',aggregate:'eval'}]}}]})),/설정/);
});

test('dynamic array headers and followers supply complete typed pivot records',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'={"지역","매출";"동부",10;"서부",20;"동부",30}'};
 const p={version:1,sourceSheetId:s.id,sourceRange:'A1:B4',rows:['지역'],columns:[],values:[{field:'매출',aggregate:'sum'}],filters:[],totals:true,sort:'asc'};
 assert.deepEqual(sourceData(b,s.id,'A1:B4').records,[['동부',10],['서부',20],['동부',30]]);assert.deepEqual(buildPivot(b,p).matrix,[['지역','합계: 매출'],['동부',40],['서부',20],['총합계',60]]);assert.equal(Object.keys(s.cells).length,1);
 s.cells.A1.raw='={"지역","매출";"#N/A",10;"#CALC!",20;"#SPILL!",30}';assert.equal(buildPivot(b,p).matched,3);
 s.cells.A1.raw='={"지역","매출";"동부",#N/A;"서부",20;"동부",30}';assert.throws(()=>buildPivot(b,p),/B2.*#N\/A/);
 s.cells.A1.raw='={"지역","매출";"동부",#CALC!;"서부",20;"동부",30}';assert.throws(()=>buildPivot(b,p),/B2.*#CALC/);
});

test('formula strings resembling errors stay text while real errors and unsupported cached values fail explicitly',()=>{
 const b=fixture(),p=spec();b.sheets[0].cells.A2.raw='="#N/A"';assert.equal(buildPivot(b,p).matched,5);
 b.sheets[0].cells.C2={raw:'=FutureFoo()',cachedFormula:'=FutureFoo()',cachedValue:99};assert.throws(()=>buildPivot(b,p),/C2.*#NAME/);
});
