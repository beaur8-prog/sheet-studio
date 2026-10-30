import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {axisScale,axisOptionsError,readAxisOptions} from '../src/chart-axis.mjs';
import {createWorkbook,validateWorkbook} from '../src/engine.mjs';
import {renderChartSVG} from '../src/chart-render.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,stringify,all,child,attr} from '../src/chart-xml.mjs';
import {attachNativeSource,buildNativePlan} from '../src/native-plan.mjs';
const fixture=()=>{const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'X'},B1:{raw:'Y'},C1:{raw:'Other'},A2:{raw:'1'},B2:{raw:'10'},C2:{raw:'1000'},A3:{raw:'2'},B3:{raw:'20'},C3:{raw:'2000'},A4:{raw:'3'},B4:{raw:'30'},C4:{raw:'3000'}};const chart={id:'axis-chart',type:'bar',range:'A1:B4',title:'Axes',width:600,height:400,firstRowAsHeader:true,firstColumnAsLabels:true};s.charts=[chart];return {b,s,chart};};
const chartXML=async b=>{const out=await writeSpreadsheet(b),zip=await JSZip.loadAsync(out.data),part=zip.file(/(?:Chart1|chart1)\.xml$/).find(p=>p.name.includes('chart'));return {out,doc:xml(await part.async('string'))};};
test('axis domains use readable intervals, explicit bounds and fractional percent units, without excessive tick counts',()=>{
 let a=axisScale([0,13]);assert.deepEqual(a.ticks,[0,5,10,15]);assert.equal(a.max,15);a=axisScale([10,20,30],{min:5,max:25,majorUnit:5});assert.deepEqual(a.ticks,[5,10,15,20,25]);assert.equal(a.ratio(5),0);assert.equal(a.ratio(25),1);assert.equal(a.contains(30),false);
 a=axisScale([0,75],{min:0,max:1,majorUnit:.25},{percent:true});assert.deepEqual(a.ticks,[0,25,50,75,100]);assert.equal(a.ratio(50),.5);assert.ok(axisScale([0,1e9],{majorUnit:1}).ticks.length<=21);assert.ok(axisScale([0,.00000004]).ticks.some(n=>n>0&&n<1e-7));
});
test('logarithmic/reversed axes omit nonpositive coordinates and retain numeric powers',()=>{
 const a=axisScale([-1,0,1,10,100],{logBase:10,reverse:true});assert.deepEqual(a.ticks,[1,10,100]);assert.equal(a.ratio(1),1);assert.equal(a.ratio(10),.5);assert.equal(a.ratio(100),0);assert.equal(a.ratio(0),null);assert.equal(a.contains(-1),false);assert.equal(axisScale([0,0],{logBase:10}).ratio(1),0);
});
test('axis schema rejects invalid bounds, increments, logarithms and persisted settings',()=>{
 for(const options of [{min:10,max:0},{majorUnit:0},{min:Infinity},{logBase:1},{logBase:10,min:0},{reverse:'yes'},{gridlines:1},{numberFormat:'a'.repeat(513)}])assert.ok(axisOptionsError(options));const {b,chart}=fixture();chart.valueAxis={majorUnit:-1};assert.throws(()=>validateWorkbook(b),/차트/);
});
test('custom axes and separate secondary axes render visible formatted ticks and clip data to the plot',()=>{
 const {b,s,chart}=fixture();chart.valueAxis={min:5,max:25,majorUnit:5,numberFormat:'0.0" kg"',gridlines:false};const doc=xml('<svg>'+renderChartSVG(chart,s,b).svg+'</svg>'),labels=all(doc,'text').filter(n=>attr(n,'class')==='chart-axis-value').map(n=>n.textContent);assert.deepEqual(labels,['5.0 kg','10.0 kg','15.0 kg','20.0 kg','25.0 kg']);assert.ok(all(doc,'clipPath').length);assert.equal(all(doc,'line').length,0);
 chart.type='combo';chart.range='A1:C4';chart.valueAxis={min:0,max:40,majorUnit:10};chart.secondaryValueAxis={min:0,max:4000,majorUnit:1000,numberFormat:'0"원"'};const second=xml('<svg>'+renderChartSVG(chart,s,b).svg+'</svg>');assert.deepEqual(all(second,'text').filter(n=>attr(n,'class')==='chart-axis-secondary-value').map(n=>n.textContent),['0원','1000원','2000원','3000원','4000원']);
 chart.type='scatter';chart.range='A1:B4';chart.xAxis={min:1,max:3,majorUnit:1,reverse:true};chart.valueAxis={min:1,max:100,logBase:10};const scatter=xml('<svg>'+renderChartSVG(chart,s,b).svg+'</svg>');assert.deepEqual(all(scatter,'text').filter(n=>attr(n,'class')==='chart-axis-x-value').map(n=>n.textContent),['1','2','3']);assert.doesNotMatch(renderChartSVG(chart,s,b).svg,/NaN|Infinity/);
});
test('new and imported axis settings survive three XLSX saves while other advanced axis XML remains intact',async()=>{
 let {b,chart}=fixture();chart.valueAxis={min:1,max:100,logBase:10,reverse:true,numberFormat:'0.0',gridlines:false};for(let i=0;i<3;i++){const {out,doc}=await chartXML(b),axis=all(doc,'valAx')[0];assert.deepEqual(readAxisOptions(axis),chart.valueAxis);b=await readSpreadsheet(out.data,'axes.xlsx');chart=b.sheets[0].charts[0];assert.deepEqual(chart.valueAxis,{min:1,max:100,logBase:10,reverse:true,numberFormat:'0.0',gridlines:false});}
 const part=chart.source.parts[chart.source.root],original=xml(Buffer.from(part.data,'base64').toString()),axis=all(original,'valAx')[0],minor=original.createElementNS(axis.namespaceURI,'c:minorUnit');minor.setAttribute('val','2');axis.appendChild(minor);part.data=Buffer.from(stringify(original)).toString('base64');
 chart.valueAxis={min:0,max:40,majorUnit:10,numberFormat:'0" units"',gridlines:true};const {doc}=await chartXML(b),changed=all(doc,'valAx')[0];assert.deepEqual(readAxisOptions(changed),{...chart.valueAxis,reverse:false});assert.equal(attr(child(changed,'minorUnit'),'val'),'2');assert.equal(child(child(changed,'scaling'),'logBase'),undefined);
});
test('XLSM keeps original axes and explicitly rejects unsupported axis/placement edits',async()=>{
 const {b,chart}=fixture(),{out}=await chartXML(b);attachNativeSource(b,out.data,'native.xlsm');assert.equal(buildNativePlan(b).sheets[0].charts.length,0);chart.valueAxis={min:2};assert.throws(()=>buildNativePlan(b),/축 범위/);delete chart.valueAxis;chart.placement='move';assert.throws(()=>buildNativePlan(b),/배치/);
});
