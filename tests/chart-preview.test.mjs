import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,address} from '../src/engine.mjs';
import {buildChartPreview} from '../src/chart-preview.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
function fixture(rows){const book=createWorkbook(),sheet=book.sheets[0];rows.forEach((row,r)=>row.forEach((v,c)=>{if(v!==null)sheet.cells[address(r,c)]={raw:String(v),style:{}};}));return {book,sheet};}

test('four horizontal numeric cells remain four preview points for all compatible chart types without changing the workbook',()=>{
 const {book,sheet}=fixture([[1,2,3,4]]),before=JSON.stringify(book),ids=[];
 for(const type of ['bar','bar-horizontal','line','area','pie','doughnut','scatter','radar','combo']){
  const result=buildChartPreview(book,sheet.id,{patch:{type,range:'A1:D1',title:'제목 <태그>',legend:true}});assert.equal(result.seriesCount,1,type);assert.deepEqual(result.rows.map(r=>r.value),['1','2','3','4'],type);assert.equal(result.points[0].total,4);assert.ok(!result.svg.includes('<태그>'));assert.ok(!result.svg.includes('chart-plot-'));ids.push(result.svg.match(/chart-preview-[a-z0-9-]+/)?.[0]);
 }
 assert.equal(JSON.stringify(book),before);assert.equal(new Set(ids.filter(Boolean)).size,ids.filter(Boolean).length);
});

test('named categories, missing and zero values, hidden rows and original references are explicit in the preview',()=>{
 const {book,sheet}=fixture([['항목','매출','비용'],['봄',10,0],['여름',null,5],['가을',30,9]]);sheet.rowHeights[3]=0;
 const patch={type:'bar',range:'A1:C4',seriesBy:'columns',firstRowAsHeader:true,firstColumnAsLabels:true};const result=buildChartPreview(book,sheet.id,{patch});assert.equal(result.seriesCount,2);assert.deepEqual(result.rows.slice(0,2).map(r=>[r.item,r.value]),[['봄','10'],['여름','빈 값']]);assert.equal(result.rows[2].value,'0');assert.match(result.references[0].values,/\$B\$2:\$B\$4/);assert.match(result.references[0].items,/\$A\$2:\$A\$4/);
 assert.equal(buildChartPreview(book,sheet.id,{patch:{...patch,plotVisibleOnly:false}}).rows.length,6);assert.equal(buildChartPreview(book,sheet.id,{patch:{...patch,type:'pie'}}).rows.filter(r=>r.used).length,2);
});

test('bubble previews pair X, Y and size and reject incomplete ranges and invalid axes',()=>{
 const {book,sheet}=fixture([['X','Y','크기'],[1,10,5],[2,20,8]]);const result=buildChartPreview(book,sheet.id,{patch:{type:'bubble',range:'A1:C3',firstRowAsHeader:true}});assert.deepEqual(result.rows.map(r=>[r.x,r.value,r.size]),[[1,'10',5],[2,'20',8]]);assert.match(result.references[0].x,/\$A\$2/);assert.match(result.references[0].size,/\$C\$2/);
 assert.throws(()=>buildChartPreview(book,sheet.id,{patch:{type:'bubble',range:'A1:B3'}}),/Y값·크기/);assert.throws(()=>buildChartPreview(book,sheet.id,{patch:{type:'bar',range:'bad'}}),/범위/);assert.throws(()=>buildChartPreview(book,sheet.id,{patch:{type:'bar',range:'A1:C3',valueAxis:{min:5,max:2}}}),/차트 설정/);
});

test('preview tables remain bounded and identify chart truncation and unsupported cached values',()=>{
 const {book,sheet}=fixture([['Value']]);for(let r=1;r<=1100;r++)sheet.cells['A'+(r+1)]={raw:String(r),style:{}};let result=buildChartPreview(book,sheet.id,{patch:{type:'bar',range:'A1:A1101',firstRowAsHeader:true}});assert.equal(result.rows.length,8);assert.equal(result.points[0].total,1100);assert.equal(result.points[0].shown,1000);assert.equal(result.truncated,true);
 sheet.cells.A2={raw:'=UnimplementedFunction()',cachedFormula:'=UnimplementedFunction()',cachedValue:7,style:{}};result=buildChartPreview(book,sheet.id,{patch:{type:'bar',range:'A1:A3',firstRowAsHeader:true}});assert.equal(result.usesCachedValues,true);assert.equal(result.rows[0].value,'7');
});

test('an imported chart preview keeps the actual series and never substitutes its fallback rectangle',async()=>{
 const {book,sheet}=fixture([['항목','수량'],['갑',11],['을',22]]);sheet.charts=[{id:'chart',type:'bar',range:'A1:B3',title:'원본',width:600,height:360,firstRowAsHeader:true,firstColumnAsLabels:true}];const read=await readSpreadsheet((await writeSpreadsheet(book)).data,'chart.xlsx'),s=read.sheets[0],chart=s.charts[0],before=JSON.stringify(read);
 const result=buildChartPreview(read,s.id,{chartId:chart.id,retain:true,patch:{title:'미리 보기'}});assert.equal(result.seriesCount,1);assert.deepEqual(result.rows.map(r=>r.value),['11','22']);assert.match(result.references[0].values,/\$B\$2:\$B\$3/);assert.equal(JSON.stringify(read),before);assert.throws(()=>buildChartPreview(read,s.id,{chartId:'missing',retain:true}),/원본 차트/);
});
