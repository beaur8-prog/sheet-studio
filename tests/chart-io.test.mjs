import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {chartFixture,chartCases,addChart} from './fixtures/charts.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {chartData,updateChartReferences} from '../src/chart-model.mjs';
import {renderChartSVG} from '../src/chart-render.mjs';
import {xml,all,attr} from '../src/chart-xml.mjs';

test('18 chart variants export native chart parts, import multiple series, and render finite SVG',async()=>{
  for(const spec of chartCases){const b=chartFixture(),s=b.sheets[0],ch=addChart(b,spec);
    const before=renderChartSVG(ch,s,b);assert.ok(!/NaN|Infinity/.test(before.svg),spec.type);
    const bytes=(await writeSpreadsheet(b)).data,zip=await JSZip.loadAsync(bytes);assert.ok(zip.file('xl/charts/sheetStudioChart1.xml'));
    const loaded=await readSpreadsheet(bytes,'charts.xlsx'),c=loaded.sheets[0].charts[0];assert.ok(c,JSON.stringify(spec));assert.equal(c.type,spec.type);assert.equal(c.series.length,['pie','bubble'].includes(spec.type)?1:2);assert.equal(c.title,ch.title);assert.equal(c.width,ch.width);
    assert.deepEqual(chartData(c,loaded.sheets[0],loaded)[0].values,[40,25,-10,35]);
    assert.ok(!/NaN|Infinity/.test(renderChartSVG(c,loaded.sheets[0],loaded).svg));
    assert.equal(parseWorkbook(serializeWorkbook(loaded)).sheets[0].charts[0].type,spec.type);
  }
});

test('imported chart caches and references follow edits, sheet rename and repeated save without path growth',async()=>{
  const b=chartFixture();addChart(b);let loaded=await readSpreadsheet((await writeSpreadsheet(b)).data,'charts.xlsx');
  loaded.sheets[0].name='다른 이름';loaded.sheets[0].cells.B2.raw='99';Object.assign(loaded.sheets[0].charts[0],{title:'수정한 제목',legend:false,labels:true,xTitle:'제품명',yTitle:'판매액'});
  let longest=0;
  for(let i=0;i<3;i++){
    const out=await writeSpreadsheet(loaded),zip=await JSZip.loadAsync(out.data),parts=zip.file(/sheetStudioCharts.*Chart1.xml$/);assert.equal(parts.length,1);
    const chart=xml(await parts[0].async('string'));assert.ok(all(chart,'f').every(n=>n.textContent.includes("'다른 이름'!")));
    if(i)assert.equal(parts[0].name.length,longest);longest=parts[0].name.length;
    loaded=await readSpreadsheet(out.data,'charts.xlsx');const restored=loaded.sheets[0].charts[0];assert.equal(restored.title,'수정한 제목');assert.equal(restored.legend,false);assert.equal(restored.labels,true);assert.equal(restored.xTitle,'제품명');assert.equal(restored.yTitle,'판매액');assert.equal(chartData(restored,loaded.sheets[0],loaded)[0].values[0],99);
  }
  const c=loaded.sheets[0].charts[0];updateChartReferences(loaded,loaded.sheets[0].id,'r',1,1);assert.equal(c.series[0].valRef.range,'B3:B6');assert.equal(c.source.formulas.find(r=>r.range==='B3:B6').sheetId,loaded.sheets[0].id);
});

test('unknown Excel chart keeps its XML and dependent image parts through export',async()=>{
  const b=chartFixture();addChart(b);const zip=await JSZip.loadAsync((await writeSpreadsheet(b)).data),part='xl/charts/sheetStudioChart1.xml';
  const original=(await zip.file(part).async('string')).replaceAll('barChart','stockChart');zip.file(part,original);
  zip.file('xl/charts/_rels/sheetStudioChart1.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/chart.png"/></Relationships>');
  zip.file('xl/media/chart.png',new Uint8Array([1,2,3,4]));
  const loaded=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'unknown.xlsx');assert.equal(loaded.sheets[0].charts[0].type,'unsupported');
  const output=await JSZip.loadAsync((await writeSpreadsheet(loaded)).data);assert.deepEqual(await output.file('xl/sheetStudioCharts/chart1/xl/media/chart.png').async('uint8array'),new Uint8Array([1,2,3,4]));
  assert.ok((await output.file('xl/sheetStudioCharts/chart1/'+part).async('string')).includes('stockChart'));
  loaded.sheets[0].charts=[];assert.equal((await JSZip.loadAsync((await writeSpreadsheet(loaded)).data)).file(/xl\/charts\/.*\.xml$/).length,0);
});

test('chart JSON rejects unsafe style, preservation path and oversized cache',()=>{
  const b=chartFixture(),chart=addChart(b);chart.series=[{color:'red" onload="alert(1)',valRef:{cache:[1]}}];assert.throws(()=>parseWorkbook(JSON.stringify(b)));
  chart.series=[{valRef:{cache:Array(10001).fill(1)}}];assert.throws(()=>parseWorkbook(JSON.stringify(b)));
});

test('chart named ranges retain their definitions and resolve current cell values',async()=>{
  const b=chartFixture();addChart(b);const zip=await JSZip.loadAsync((await writeSpreadsheet(b)).data);
  const file='xl/charts/sheetStudioChart1.xml',source=await zip.file(file).async('string'),doc=xml(source),formula=all(doc,'f').find(n=>n.textContent.includes('$B$2:$B$5'));
  const old=formula.textContent;formula.textContent='ChartRevenue';
  const {stringify,xmlEscape}=await import('../src/chart-xml.mjs');zip.file(file,stringify(doc));
  const bookXML=await zip.file('xl/workbook.xml').async('string');zip.file('xl/workbook.xml',bookXML.replace('<calcPr',`<definedNames><definedName name="ChartRevenue">${xmlEscape(old)}</definedName></definedNames><calcPr`));
  const loaded=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'named.xlsx');loaded.sheets[0].cells.B2.raw='81';
  assert.equal(chartData(loaded.sheets[0].charts[0],loaded.sheets[0],loaded)[0].values[0],81);
  loaded.sheets[0].name='이름 변경';const out=await JSZip.loadAsync((await writeSpreadsheet(loaded)).data),names=all(xml(await out.file('xl/workbook.xml').async('string')),'definedName');assert.equal(attr(names[0],'name'),'ChartRevenue');assert.ok(names[0].textContent.includes("'이름 변경'!"));
  const chartPart=out.file(/sheetStudioCharts.*Chart1.xml$/)[0];
  assert.ok(all(xml(await chartPart.async('string')),'f').some(n=>n.textContent==='ChartRevenue'),'the chart must keep its link to the defined name');
});

test('modern ChartEx carriers remain opaque and retain their relationship and chart XML',async()=>{
  const b=chartFixture();addChart(b);const zip=await JSZip.loadAsync((await writeSpreadsheet(b)).data),part='xl/charts/sheetStudioChart1.xml',cx='http://schemas.microsoft.com/office/drawing/2014/chartex';
  const modern=`<cx:chartSpace xmlns:cx="${cx}"><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"/></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;
  zip.file(part,modern);
  const drawing='xl/drawings/sheetStudioDrawing1.xml',rels='xl/drawings/_rels/sheetStudioDrawing1.xml.rels';
  zip.file(drawing,(await zip.file(drawing).async('string')).replaceAll('http://schemas.openxmlformats.org/drawingml/2006/chart',cx));
  zip.file(rels,(await zip.file(rels).async('string')).replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart','http://schemas.microsoft.com/office/2014/relationships/chartEx'));
  const book=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'modern.xlsx');assert.equal(book.sheets[0].charts[0].type,'unsupported');
  const output=await JSZip.loadAsync((await writeSpreadsheet(book)).data);assert.ok((await output.file('xl/sheetStudioCharts/chart1/'+part).async('string')).includes('layoutId="waterfall"'));assert.ok((await output.file(rels).async('string')).includes('/chartEx'));
});
