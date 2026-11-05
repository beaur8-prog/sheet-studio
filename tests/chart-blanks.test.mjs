import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createWorkbook,address,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {chartData} from '../src/chart-model.mjs';
import {chartBlankPolicy} from '../src/chart-blanks.mjs';
import {renderChartSVG} from '../src/chart-render.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,all,first,children,attr,val} from '../src/chart-xml.mjs';
import {attachNativeSource,buildNativePlan} from '../src/native-plan.mjs';

function fixture(){const book=createWorkbook(),sheet=book.sheets[0];['a','b','c','d','e'].forEach((v,r)=>sheet.cells['A'+(r+1)]={raw:v,style:{}});[10,null,30,'=NA()',50].forEach((v,r)=>{if(v!==null)sheet.cells['B'+(r+1)]={raw:String(v),style:{}};});const chart={id:'chart',type:'line',range:'A1:B5',width:600,height:360,title:'빈 값',firstRowAsHeader:false,firstColumnAsLabels:true};sheet.charts=[chart];return {book,sheet,chart};}
const lineLengths=(chart,sheet,book)=>{const doc=xml('<svg xmlns="http://www.w3.org/2000/svg">'+renderChartSVG(chart,sheet,book).svg+'</svg>'),plot=all(doc,'g').find(n=>attr(n,'class')==='chart-data');return all(plot,'polyline').map(n=>attr(n,'points').trim().split(/\s+/).length);};

test('line gaps, zero and span only affect physical blanks; errors and literal empty text remain gaps',()=>{
 const {book,sheet,chart}=fixture(),before=JSON.stringify(book);assert.deepEqual(chartData(chart,sheet,book)[0].values,[10,null,30,null,50]);assert.deepEqual(lineLengths(chart,sheet,book),[1,1,1]);
 assert.deepEqual(chartData({...chart,blankCells:'zero'},sheet,book)[0].values,[10,0,30,null,50]);assert.deepEqual(lineLengths({...chart,blankCells:'zero'},sheet,book),[3,1]);assert.deepEqual(lineLengths({...chart,blankCells:'span'},sheet,book),[2,1]);assert.equal(JSON.stringify(book),before);
 sheet.cells.B2={raw:'=""',style:{}};assert.equal(chartData({...chart,blankCells:'zero'},sheet,book)[0].values[1],null);assert.deepEqual(lineLengths({...chart,blankCells:'span'},sheet,book),[1,1,1]);
});

test('date sorting and hidden masks keep blank identity aligned with the plotted categories',()=>{
 const {book,sheet,chart}=fixture();for(let r=0;r<5;r++)sheet.cells[address(r,0)]={raw:String([45005,45001,45004,45002,45003][r]),style:{format:'date',numberFormat:'yyyy-mm-dd'}};
 let result=chartData({...chart,blankCells:'span'},sheet,book)[0];assert.deepEqual(result.values,[null,null,50,30,10]);assert.deepEqual(result.blankPoints,[true,false,false,false,false]);sheet.rowHeights[1]=0;result=chartData({...chart,blankCells:'span'},sheet,book)[0];assert.deepEqual(result.blankPoints,[false,false,false,false]);assert.deepEqual(result.values,[null,50,30,10]);
});

test('radar gaps do not silently pass through zero, while span closes the remaining physical-blank path',()=>{
 const {book,sheet,chart}=fixture();chart.type='radar';delete sheet.cells.B4;assert.deepEqual(lineLengths(chart,sheet,book),[1,2]);assert.deepEqual(lineLengths({...chart,blankCells:'span'},sheet,book),[4]);assert.equal(chartData({...chart,blankCells:'zero'},sheet,book)[0].values.filter(v=>v===0).length,2);
});

test('scatter lines honor Y and X gaps and keep markers-only series unconnected',()=>{
 const {book,sheet}=fixture(),chart={id:'xy',type:'scatter',width:600,height:360,series:[{type:'scatter',name:'XY',lines:true,xRef:{sheetId:sheet.id,range:'C1:C5'},valRef:{sheetId:sheet.id,range:'B1:B5'}}]};for(let r=0;r<5;r++)sheet.cells['C'+(r+1)]={raw:String(r+1),style:{}};
 assert.deepEqual(lineLengths(chart,sheet,book),[1,1,1]);assert.deepEqual(lineLengths({...chart,blankCells:'span'},sheet,book),[2,1]);assert.deepEqual(lineLengths({...chart,blankCells:'zero'},sheet,book),[3,1]);assert.deepEqual(lineLengths({...chart,series:[{...chart.series[0],lines:false}],blankCells:'span'},sheet,book),[]);
 delete sheet.cells.C3;assert.equal(chartData({...chart,blankCells:'zero'},sheet,book)[0].xs[2],0);
});

test('XLSX blank policies survive native import, partial editing, legacy app JSON and three roundtrips',async()=>{
 const {book,sheet,chart}=fixture();chart.blankCells='span';let current=book;
 for(let i=0;i<3;i++){
  const out=await writeSpreadsheet(current),zip=await JSZip.loadAsync(out.data),file=zip.file(/(?:Chart|chart)\d+\.xml$/).find(p=>p.name.includes('chart')),doc=xml(await file.async('string'));assert.equal(val(first(doc,'chart'),'dispBlanksAs'),'span');
  current=await readSpreadsheet(out.data,'blank.xlsx');assert.equal(current.sheets[0].charts[0].blankCells,'span');assert.deepEqual(lineLengths(current.sheets[0].charts[0],current.sheets[0],current),[2,1]);
 }
 const native=current.sheets[0].charts[0];delete native.blankCells;delete native.source.original.blankCells;current=parseWorkbook(serializeWorkbook(current));assert.equal(chartBlankPolicy(current.sheets[0].charts[0]),'span');assert.deepEqual(lineLengths(current.sheets[0].charts[0],current.sheets[0],current),[2,1]);
 const preserved=current.sheets[0].charts[0];preserved.blankCells='zero';preserved.title='편집';const saved=await JSZip.loadAsync((await writeSpreadsheet(current)).data),file=saved.file(/(?:Chart|chart)\d+\.xml$/).find(p=>p.name.includes('chart')),doc=xml(await file.async('string'));assert.equal(val(first(doc,'chart'),'dispBlanksAs'),'zero');assert.ok(first(doc,'lineChart'));assert.equal(all(doc,'ser').length,1);assert.match(all(doc,'f').map(n=>n.textContent).join(' '),/\$B\$1:\$B\$5/);
 assert.throws(()=>parseWorkbook(JSON.stringify({...book,sheets:[{...sheet,charts:[{...chart,blankCells:'bad'}]}]})),/차트 설정/);
});

test('XLSM keeps its original blank policy and explicitly rejects unsupported blank-policy edits',async()=>{
 const {book,chart}=fixture();chart.blankCells='span';attachNativeSource(book,(await writeSpreadsheet(book)).data,'native.xlsm');assert.equal(buildNativePlan(book).sheets[0].charts.length,0);chart.title='제목 변경';assert.equal(buildNativePlan(book).sheets[0].charts[0].title,'제목 변경');chart.blankCells='zero';assert.throws(()=>buildNativePlan(book),/빈 셀 표시/);
});
