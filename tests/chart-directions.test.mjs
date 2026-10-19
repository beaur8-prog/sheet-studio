import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,address,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {chartData,chartSeries,chartSeriesBy,chartLayout,chartRangeError} from '../src/chart-model.mjs';
import {renderChartSVG} from '../src/chart-render.mjs';
import {chartCases} from './fixtures/charts.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';
import {attachNativeSource,buildNativePlan} from '../src/native-plan.mjs';
import {createRequire} from 'node:module';
const {validatePayload}=createRequire(import.meta.url)('../electron/vba.cjs');
const fixture=rows=>{const b=createWorkbook();rows.forEach((row,r)=>row.forEach((v,c)=>{if(v!==null)b.sheets[0].cells[address(r,c)]={raw:String(v)};}));return b;};
const transpose=rows=>rows[0].map((_,c)=>rows.map(row=>row[c]));
const values=(b,c)=>chartData(c,b.sheets[0],b).map(s=>s.values);
const pointCount=svg=>(svg.match(/class="chart-point"/g)||[]).length;
const projected=(b,c)=>chartData(c,b.sheets[0],b).map(({name,labels,values,xs,sizes,type})=>({name,labels,values,xs,sizes,type}));

test('all compatible types keep four horizontal or vertical values in one series',()=>{
 for(const type of ['bar','bar-horizontal','line','area','pie','doughnut','radar','combo','scatter'])for(const rows of [[[10,20,30,40]],[[10],[20],[30],[40]]]){
  const b=fixture(rows),c={type,range:`A1:${address(rows.length-1,rows[0].length-1)}`,legend:true};
  assert.deepEqual(values(b,c),[[10,20,30,40]],type);assert.equal(pointCount(renderChartSVG(c,b.sheets[0],b).svg),4,type);
  if(type==='scatter')assert.deepEqual(chartData(c,b.sheets[0],b)[0].xs,[1,2,3,4]);
 }
});
test('row names and column headings work symmetrically for both series directions',()=>{
 for(const rows of [[['매출',10,20,30,40]],transpose([['매출',10,20,30,40]])]){
  const b=fixture(rows),c={type:'pie',range:`A1:${address(rows.length-1,rows[0].length-1)}`},data=chartData(c,b.sheets[0],b);
  assert.equal(data[0].name,'매출');assert.deepEqual(data[0].values,[10,20,30,40]);
 }
 const b=fixture([['항목','가','나','다','라'],['매출',10,20,30,40],['비용',5,8,12,15]]),c={type:'doughnut',range:'A1:E3',seriesBy:'rows'};
 assert.deepEqual(chartLayout(c,b.sheets[0],b),{firstRowAsHeader:true,firstColumnAsLabels:true});assert.deepEqual(values(b,c),[[10,20,30,40],[5,8,12,15]]);
 assert.deepEqual(chartData(c,b.sheets[0],b).map(s=>s.name),['매출','비용']);assert.deepEqual(chartData(c,b.sheets[0],b)[0].labels,['가','나','다','라']);
 assert.equal(pointCount(renderChartSVG(c,b.sheets[0],b).svg),8);
});
test('explicit directions preserve independent series and native row references',()=>{
 const b=fixture([[10,20,30,40],[5,8,12,15]]),c={type:'line',range:'A1:D2',seriesBy:'rows',firstRowAsHeader:false,firstColumnAsLabels:false};
 assert.deepEqual(chartSeries(c,b.sheets[0],b).map(s=>s.valRef.range),['A1:D1','A2:D2']);assert.equal(chartSeriesBy(c),'rows');
 assert.deepEqual(values(b,{...c,seriesBy:'columns'}),[[10,5],[20,8],[30,12],[40,15]]);
 b.sheets[0].charts=[{id:'direction',...c}];assert.equal(parseWorkbook(serializeWorkbook(b)).sheets[0].charts[0].seriesBy,'rows');
 b.sheets[0].charts[0].seriesBy='invalid';assert.throws(()=>serializeWorkbook(b),/차트/);
});
test('scatter rows and columns bind every explicit X and Y coordinate',()=>{
 const source=[['X',1,2,3,4],['Y',10,20,30,40],['Z',15,25,35,45]];
 for(const [rows,seriesBy] of [[source,'rows'],[transpose(source),'columns']]){
  const b=fixture(rows),c={type:'scatter',seriesBy,range:`A1:${address(rows.length-1,rows[0].length-1)}`},data=chartData(c,b.sheets[0],b);
  assert.equal(chartRangeError(c,b.sheets[0],b),null);assert.deepEqual(data.map(s=>s.xs),[[1,2,3,4],[1,2,3,4]]);assert.deepEqual(data.map(s=>s.values),[[10,20,30,40],[15,25,35,45]]);assert.equal(pointCount(renderChartSVG(c,b.sheets[0],b).svg),8);
 }
});
test('bubble requires complete X/Y/size sets and never drops an unmatched series',()=>{
 const source=[['X',1,2,3,4],['Y',10,20,30,40],['크기',5,10,15,20]];
 for(const [rows,seriesBy] of [[source,'rows'],[transpose(source),'columns']]){
  const b=fixture(rows),c={type:'bubble',seriesBy,range:`A1:${address(rows.length-1,rows[0].length-1)}`},data=chartData(c,b.sheets[0],b);
  assert.equal(chartRangeError(c,b.sheets[0],b),null);assert.deepEqual(data[0].values,[10,20,30,40]);assert.deepEqual(data[0].sizes,[5,10,15,20]);assert.equal(pointCount(renderChartSVG(c,b.sheets[0],b).svg),4);
 }
 for(const rows of [[[10,20,30,40]],[[10],[20],[30],[40]],[[1,10,5,99],[2,20,8,88]]]){
  const b=fixture(rows),c={type:'bubble',range:`A1:${address(rows.length-1,rows[0].length-1)}`};assert.match(chartRangeError(c,b.sheets[0],b),/X값·Y값·크기/);assert.deepEqual(values(b,c),[]);
 }
});
test('blank explicit X and zero bubble sizes do not create invented points',()=>{
 const b=fixture([[1,10,5],[null,20,8],[3,30,0],[4,40,10]]),s=b.sheets[0];
 assert.equal(pointCount(renderChartSVG({type:'scatter',range:'A1:B4'},s,b).svg),3);
 assert.equal(pointCount(renderChartSVG({type:'bubble',range:'A1:C4'},s,b).svg),2);
});
test('pie uses a single full series, doughnut uses all rings, zeroes make no slice',()=>{
 const b=fixture([[10,20,30,40],[20,30,40,50]]),s=b.sheets[0],c={type:'pie',range:'A1:D2',seriesBy:'rows',legend:true};
 assert.equal(pointCount(renderChartSVG(c,s,b).svg),4);assert.equal(pointCount(renderChartSVG({...c,type:'doughnut'},s,b).svg),8);
 for(const type of ['pie','doughnut']){const single=fixture([[0,0,30,0]]),svg=renderChartSVG({type,range:'A1:D1'},single.sheets[0],single).svg;assert.equal(pointCount(svg),1);assert.match(svg,/<title>3: 30<\/title>/);}
});
test('all 18 chart variants preserve horizontal and vertical data in JSON and native XLSX',async()=>{
 for(const spec of chartCases)for(const seriesBy of ['rows','columns']){
  const xy=['scatter','bubble'].includes(spec.type),rows=xy?[[1,2,3,4],[10,20,30,40],...(spec.type==='bubble'?[[5,10,15,20]]:[])]:[[10,20,30,40],[5,8,12,15]],source=seriesBy==='rows'?rows:transpose(rows),b=fixture(source),s=b.sheets[0],c={id:'audit',...spec,seriesBy,firstRowAsHeader:false,firstColumnAsLabels:false,range:`A1:${address(source.length-1,source[0].length-1)}`,title:'선택 데이터'};s.charts=[c];
  assert.deepEqual(projected(parseWorkbook(serializeWorkbook(b)),c),projected(b,c));
  const loaded=await readSpreadsheet((await writeSpreadsheet(b)).data,'all-types.xlsx'),restored=loaded.sheets[0].charts[0],expected=projected(b,c).slice(0,spec.type==='pie'?1:32);
  assert.deepEqual(projected(loaded,restored),expected,JSON.stringify({spec,seriesBy}));assert.ok(!/NaN|Infinity/.test(renderChartSVG(restored,loaded.sheets[0],loaded).svg));
 }
});
test('Windows native plans accept row vectors and rebuild direction changes',async()=>{
 const b=fixture([['항목','가','나','다','라'],['매출',10,20,30,40],['費用',5,8,12,15]]),s=b.sheets[0];attachNativeSource(b,(await writeSpreadsheet(b)).data,'charts.xlsm');
 s.charts=[{id:'native-rows',type:'bar',range:'A1:E3',seriesBy:'rows',firstRowAsHeader:true,firstColumnAsLabels:true}];let plan=buildNativePlan(b);
 assert.deepEqual(plan.sheets[0].charts[0].series,[{valuesRange:'B2:E2',nameRange:'A2:A2',categoriesRange:'B1:E1'},{valuesRange:'B3:E3',nameRange:'A3:A3',categoriesRange:'B1:E1'}]);validatePayload({mode:'save',source:b.nativeSource,plan});
 b.nativeSource.baseline=structuredClone(b.sheets);s.charts[0].seriesBy='columns';plan=buildNativePlan(b);assert.equal(plan.sheets[0].charts[0].series.length,4);validatePayload({mode:'save',source:b.nativeSource,plan});
 for(const valuesRange of ['B2:E3','E2:B2','A1:XFD1']){const bad=structuredClone(plan);bad.sheets[0].charts[0].series=[{valuesRange}];assert.throws(()=>validatePayload({mode:'save',source:b.nativeSource,plan:bad}));}
});
test('limits apply to values per series in either direction, including wide vectors',()=>{
 const b=fixture([[1]]),s=b.sheets[0];assert.equal(chartRangeError({type:'bar',range:'A1:AG1'},s,b),null);assert.equal(chartRangeError({type:'bar',range:'A1:A10000'},s,b),null);
 assert.match(chartRangeError({type:'bar',range:'A1:A10001',firstRowAsHeader:false},s,b),/10,000/);assert.match(chartRangeError({type:'bar',range:'A1:AG2',seriesBy:'columns',firstColumnAsLabels:false},s,b),/32/);
 assert.match(renderChartSVG({type:'line',range:'B1:E1'},s,b).svg,/표시할 숫자 데이터/);
});
