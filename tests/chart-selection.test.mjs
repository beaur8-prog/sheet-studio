import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,address} from '../src/engine.mjs';
import {chartData} from '../src/chart-model.mjs';
const fixture=rows=>{const b=createWorkbook();rows.forEach((row,r)=>row.forEach((v,c)=>b.sheets[0].cells[address(r,c)]={raw:String(v)}));return b;};
test('column charts retain the first numeric column and first numeric row',()=>{
 const b=fixture([[10,20],[30,40],[50,60]]),s=b.sheets[0],data=chartData({type:'bar',range:'A1:B3'},s,b);
 assert.equal(data.length,2);assert.deepEqual(data.map(s=>s.values),[[10,30,50],[20,40,60]]);
});

test('text headers retain numeric first-column series, while text categories stay on the axis',()=>{
 let b=fixture([['매출','비용','수량'],[40,22,8],[25,18,12]]),s=b.sheets[0],data=chartData({type:'bar',range:'A1:C3'},s,b);
 assert.deepEqual(data.map(s=>s.name),['매출','비용','수량']);assert.deepEqual(data.map(s=>s.values),[[40,25],[22,18],[8,12]]);
 b=fixture([['상품','매출','비용'],['커피',40,22],['차',25,18]]);s=b.sheets[0];data=chartData({type:'bar',range:'A1:C3'},s,b);
 assert.equal(data.length,2);assert.deepEqual(data[0].labels,['커피','차']);assert.deepEqual(data[0].values,[40,25]);
 b=fixture([['커피',40,22],['차',25,18]]);s=b.sheets[0];data=chartData({type:'bar',range:'A1:C2'},s,b);assert.deepEqual(data[0].values,[40,25]);assert.deepEqual(data[0].labels,['커피','차']);
});

test('single columns, single numeric rows, zeroes and formula results remain plottable',()=>{
 for(const [rows,range,expected] of [
  [[['매출'],[0],['=10+20']],'A1:A3',[[0,30]]],
  [[[10],[20],[30]],'A1:A3',[[10,20,30]]],
  [[[10,20,30]],'A1:C1',[[10,20,30]]],
  [[[0]],'A1:A1',[[0]]],
  [[['매출','비용']],'A1:B1',[]],
 ]){const b=fixture(rows);assert.deepEqual(chartData({type:'bar',range},b.sheets[0],b).map(s=>s.values),expected);}
});

test('explicit first-row/column choices persist and numeric categories can be chosen deliberately',async()=>{
 const {chartLayout,chartRangeError}=await import('../src/chart-model.mjs'),{serializeWorkbook,parseWorkbook}=await import('../src/engine.mjs');
 const b=fixture([[2024,2025,2026],[1,30,50],[2,40,60]]),s=b.sheets[0],chart={id:'labels',type:'bar',range:'A1:C3',firstRowAsHeader:true,firstColumnAsLabels:true};s.charts=[chart];
 assert.deepEqual(chartData(chart,s,b).map(s=>s.values),[[30,40],[50,60]]);assert.deepEqual(chartData(chart,s,b)[0].labels,['1','2']);
 assert.deepEqual(parseWorkbook(serializeWorkbook(b)).sheets[0].charts[0],chart);
 assert.deepEqual(chartLayout({...chart,firstColumnAsLabels:false},s,b),{firstRowAsHeader:true,firstColumnAsLabels:false});
 assert.match(chartRangeError({...chart,range:'A1:A3'},s,b),/값/);assert.match(chartRangeError({...chart,range:'A1:C1'},s,b),/값/);
 assert.match(chartRangeError({type:'bar',range:'A1:AG3',firstColumnAsLabels:false},s,b),/32/);
 assert.match(chartRangeError({type:'bar',range:'A1:B10001',firstRowAsHeader:false},s,b),/10,000/);
 assert.equal(chartRangeError({type:'bar',range:'A1:B10001',firstRowAsHeader:true},s,b),null);
 for(const field of ['firstRowAsHeader','firstColumnAsLabels']){const bad=structuredClone(b);bad.sheets[0].charts[0][field]='true';assert.throws(()=>serializeWorkbook(bad),/차트/);}
});

test('JSON and XLSX roundtrips retain every selected numeric column and row',async()=>{
 const {readSpreadsheet,writeSpreadsheet}=await import('../src/io.mjs');
 for(const rows of [[[10,20],[30,40],[50,60]],[['매출','비용'],[10,20],[30,40]]]){
  const b=fixture(rows),s=b.sheets[0],chart={id:'numeric',type:'bar',range:`A1:B${rows.length}`};s.charts=[chart];
  const expected=chartData(chart,s,b).map(({name,labels,values})=>({name,labels,values})),loaded=await readSpreadsheet((await writeSpreadsheet(b)).data,'numeric-chart.xlsx');
  assert.deepEqual(chartData(loaded.sheets[0].charts[0],loaded.sheets[0],loaded).map(({name,labels,values})=>({name,labels,values})),expected);
 }
});

test('native Excel plans bind explicit series ranges and rebuild when label choices change',async()=>{
 const {attachNativeSource,buildNativePlan}=await import('../src/native-plan.mjs'),{writeSpreadsheet}=await import('../src/io.mjs'),{createRequire}=await import('node:module'),{validatePayload}=createRequire(import.meta.url)('../electron/vba.cjs');
 const b=fixture([['매출','비용'],[10,20],[30,40]]),s=b.sheets[0];attachNativeSource(b,(await writeSpreadsheet(b)).data,'charts.xlsm');
 s.charts=[{id:'new',type:'bar',range:'A1:B3',firstRowAsHeader:true,firstColumnAsLabels:false}];let plan=buildNativePlan(b);
 assert.deepEqual(plan.sheets[0].charts[0].series,[{valuesRange:'A2:A3',nameRange:'A1:A1'},{valuesRange:'B2:B3',nameRange:'B1:B1'}]);validatePayload({mode:'save',source:b.nativeSource,plan});
 b.nativeSource.baseline=structuredClone(b.sheets);s.charts[0].firstColumnAsLabels=true;plan=buildNativePlan(b);assert.deepEqual(plan.sheets[0].charts[0].series,[{valuesRange:'B2:B3',nameRange:'B1:B1',categoriesRange:'A2:A3'}]);validatePayload({mode:'save',source:b.nativeSource,plan});
 for(const series of [[{valuesRange:'A2:B3'}],[{valuesRange:'A2:A3',categoriesRange:'B2:B2'}],[{valuesRange:'A2:A3',nameRange:'A1:A2'}],[{valuesRange:'A2:A3',Formula:'evil'}]]){const bad=structuredClone(plan);bad.sheets[0].charts[0].series=series;assert.throws(()=>validatePayload({mode:'save',source:b.nativeSource,plan:bad}),/차트/);}
});


test('an empty selection gives a data hint while unsupported imported charts keep the preservation notice',async()=>{
 const {renderChartSVG}=await import('../src/chart-render.mjs'),b=fixture([]),s=b.sheets[0];
 assert.match(renderChartSVG({type:'bar',range:'A1:B2'},s,b).svg,/표시할 숫자 데이터/);
 assert.match(renderChartSVG({type:'unsupported',range:'A1:B2'},s,b).svg,/Excel 원본 차트 보존/);
});
