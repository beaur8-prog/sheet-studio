import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createWorkbook,address,evaluateCell} from '../src/engine.mjs';
import {chartData} from '../src/chart-model.mjs';
import {renderChartSVG} from '../src/chart-render.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,all,attr} from '../src/chart-xml.mjs';

function fixture(rows){
  const book=createWorkbook(),sheet=book.sheets[0];
  rows.forEach((row,r)=>row.forEach((value,c)=>{if(value!==null)sheet.cells[address(r,c)]={raw:String(value)};}));
  return {book,sheet};
}
const chart=(patch={})=>({id:'parity-chart',type:'bar',range:'A1:B4',width:600,height:360,legend:true,...patch});
const outputChart=async book=>{
  const zip=await JSZip.loadAsync((await writeSpreadsheet(book)).data);
  const part=zip.file(/(?:Chart1|chart1)\.xml$/).find(p=>p.name.includes('chart'));
  return xml(await part.async('string'));
};

test('chart visibility removes filtered rows and hidden value columns without shifting categories',async()=>{
  const {book,sheet}=fixture([['항목','매출','비용'],['가',10,1],['나',999,2],['다',30,3]]);
  sheet.autoFilter={range:{r1:0,c1:0,r2:3,c2:2},columns:{},hiddenRows:[2]};
  sheet.columnWidths[2]=0;
  const c=chart({range:'A1:C4'}),data=chartData(c,sheet,book);
  assert.equal(data.length,1);
  assert.deepEqual(data[0].values,[10,30]);
  assert.deepEqual(data[0].labels,['가','다']);
  assert.equal((renderChartSVG(c,sheet,book).svg.match(/class="chart-point"/g)||[]).length,2);
  assert.deepEqual(chartData({...c,plotVisibleOnly:false},sheet,book).map(s=>s.values),[[10,999,30],[1,2,3]]);
  sheet.charts.push({...c,plotVisibleOnly:false});
  const doc=await outputChart(book);
  assert.equal(attr(all(doc,'plotVisOnly')[0],'val'),'0');
  const loaded=await readSpreadsheet((await writeSpreadsheet(book)).data,'visible.xlsx');
  assert.equal(loaded.sheets[0].charts[0].plotVisibleOnly,false);
});

test('manual hidden rows and horizontal hidden columns use the same point mask',()=>{
  const {book,sheet}=fixture([['종류','가','나','다'],['매출',10,999,30]]);
  sheet.columnWidths[2]=0;
  let c=chart({range:'A1:D2',seriesBy:'rows'});
  assert.deepEqual(chartData(c,sheet,book)[0].values,[10,30]);
  assert.deepEqual(chartData(c,sheet,book)[0].labels,['가','다']);
  sheet.columnWidths={};sheet.rowHeights[1]=0;
  assert.deepEqual(chartData(c,sheet,book),[]);
});

test('current formula errors never turn back into historical successful chart values',async()=>{
  const {book,sheet}=fixture([['항목','매출'],['가',40],['나',10]]);
  sheet.cells.B2={raw:'=40',cachedValue:40,cachedFormula:'=40'};
  const c=chart({range:'A1:B3'});sheet.charts.push(c);
  const loaded=await readSpreadsheet((await writeSpreadsheet(book)).data,'formulas.xlsx'),s=loaded.sheets[0];
  s.cells.B2.raw='=1/0';
  assert.equal(evaluateCell(loaded,s.id,'B2'),'#DIV/0!');
  assert.deepEqual(chartData(s.charts[0],s,loaded)[0].values,[null,10]);
  const doc=await outputChart(loaded);
  const values=all(all(doc,'val')[0],'v').map(n=>n.textContent);
  assert.ok(!values.includes('40'));
  s.cells.B2.raw='=NA()';
  assert.deepEqual(chartData(s.charts[0],s,loaded)[0].values,[null,10]);
  s.cells.B2={raw:'=UNSUPPORTED()',cachedValue:7,cachedFormula:'=UNSUPPORTED()'};
  assert.deepEqual(chartData(s.charts[0],s,loaded)[0].values,[7,10]);
  s.cells.B2.raw='=ANOTHER_UNSUPPORTED()';
  assert.deepEqual(chartData(s.charts[0],s,loaded)[0].values,[null,10]);
});

test('stacked combination charts do not add column heights to the secondary line',()=>{
  const {book,sheet}=fixture([['항목','매출','비율'],['가',100,10],['나',50,5]]);
  for(const grouping of ['stacked','percentStacked']){
    const result=renderChartSVG(chart({type:'combo',grouping,range:'A1:C3'}),sheet,book);
    const height=Number(result.viewBox.split(' ')[3]);
    const ys=[...result.svg.matchAll(/<circle class="chart-point"[^>]*cy="([^"]+)"/g)].map(m=>Number(m[1]));
    assert.equal(ys.length,2);
    assert.ok(ys.every(y=>y>=0&&y<=height),JSON.stringify({grouping,ys,height}));
  }
});

test('bubble size is comparable across series in the same chart',()=>{
  const {book,sheet}=fixture([['X','Y1','크기1','Y2','크기2'],[1,10,25,20,100]]);
  const svg=renderChartSVG(chart({type:'bubble',range:'A1:E2'}),sheet,book).svg;
  const radii=[...svg.matchAll(/<circle class="chart-point"[^>]* r="([^"]+)"/g)].map(m=>Number(m[1]));
  assert.equal(radii.length,2);
  assert.ok(Math.abs(radii[1]/radii[0]-2)<1e-10);
});

test('pie legend colors match slices and percentage labels survive XLSX',async()=>{
  const {book,sheet}=fixture([[1,3]]),c=chart({type:'pie',range:'A1:B1',labels:true,series:[{type:'pie',valRef:{cache:[1,3]},catRef:{cache:['가','나']},pointColors:{0:'#ff0000',1:'#00ff00'}}]});
  const svg=renderChartSVG(c,sheet,book).svg;
  assert.match(svg,/>25%<\/text>/);assert.match(svg,/>75%<\/text>/);
  const legend=svg.slice(svg.indexOf('class="chart-legend-row"'));
  assert.match(legend,/fill="#ff0000"/);assert.match(legend,/fill="#00ff00"/);
  sheet.charts.push(c);
  const doc=await outputChart(book);
  assert.equal(attr(all(doc,'showVal')[0],'val'),'0');
  assert.equal(attr(all(doc,'showPercent')[0],'val'),'1');
  const loaded=await readSpreadsheet((await writeSpreadsheet(book)).data,'pie.xlsx');
  assert.equal(loaded.sheets[0].charts[0].labels,true);
  const restoredSVG=renderChartSVG(loaded.sheets[0].charts[0],loaded.sheets[0],loaded).svg;
  assert.match(restoredSVG.slice(restoredSVG.indexOf('class="chart-legend-row"')),/fill="#ff0000"/);
});

test('explicit single-series line legend is respected while column legend stays omitted',()=>{
  const {book,sheet}=fixture([[10,20,30]]);
  assert.match(renderChartSVG(chart({type:'line',range:'A1:C1'}),sheet,book).svg,/chart-legend-item/);
  assert.doesNotMatch(renderChartSVG(chart({type:'bar',range:'A1:C1'}),sheet,book).svg,/chart-legend-item/);
});

test('date categories and percentage axis labels use the cell display formats',async()=>{
  const {book,sheet}=fixture([['날짜','비율'],[46022,.25],[46032,.5]]);
  for(const ref of ['A2','A3'])sheet.cells[ref].style={format:'date',numberFormat:'yyyy-mm-dd'};
  for(const ref of ['B2','B3'])sheet.cells[ref].style={format:'percent'};
  const c=chart({type:'line',range:'A1:B3',firstColumnAsLabels:true,labels:true});
  assert.ok(chartData(c,sheet,book)[0].labels.every(v=>/^\d{4}-\d{2}-\d{2}$/.test(v)));
  const svg=renderChartSVG(c,sheet,book).svg;
  assert.match(svg,/>25\.00%<\/text>/);assert.match(svg,/>50\.00%<\/text>/);
  sheet.charts.push(c);const doc=await outputChart(book);
  assert.equal(all(doc,'dateAx').length,1);
  assert.ok(all(doc,'numFmt').some(n=>attr(n,'formatCode')==='0.00%'));
  assert.ok(all(doc,'formatCode').some(n=>n.textContent==='yyyy-mm-dd'));
});

test('radar values below zero have a distinct radius from zero',()=>{
  const {book,sheet}=fixture([[-10,0,10]]),svg=renderChartSVG(chart({type:'radar',range:'A1:C1'}),sheet,book).svg;
  const points=[...svg.matchAll(/<circle class="chart-point"[^>]*cx="([^"]+)" cy="([^"]+)"/g)].map(m=>m.slice(1).map(Number));
  assert.equal(points.length,3);
  assert.notDeepEqual(points[0],points[1]);
});


test('date axes preserve unequal intervals and chronological data alignment',()=>{
  const {book,sheet}=fixture([['날짜','매출'],[46032,20],[46022,10],[46033,30]]);
  for(const ref of ['A2','A3','A4'])sheet.cells[ref].style={format:'date'};
  const c=chart({type:'line',range:'A1:B4',firstColumnAsLabels:true}),data=chartData(c,sheet,book);
  assert.deepEqual(data[0].values,[10,20,30]);
  const svg=renderChartSVG(c,sheet,book).svg;
  const xs=[...svg.matchAll(/<circle class="chart-point"[^>]*cx="([^"]+)"/g)].map(m=>Number(m[1]));
  assert.equal(xs.length,3);
  assert.ok(Math.abs((xs[1]-xs[0])/(xs[2]-xs[1])-10)<1e-9);
});
