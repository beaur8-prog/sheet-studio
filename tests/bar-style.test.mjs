import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createWorkbook,address} from '../src/engine.mjs';
import {chartData,chartSeriesBy} from '../src/chart-model.mjs';
import {renderChartSVG} from '../src/chart-render.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';
import {attachNativeSource,buildNativePlan} from '../src/native-plan.mjs';
const fixture=rows=>{const b=createWorkbook();rows.forEach((row,r)=>row.forEach((v,c)=>b.sheets[0].cells[address(r,c)]={raw:String(v)}));return b;};

test('one observation beneath named headings becomes a single series with category labels',()=>{
 for(const rows of [[['커피','차','주스','물'],[120,90,150,80]],[['항목','커피','차','주스','물'],['매출',120,90,150,80]]]){
  const b=fixture(rows),s=b.sheets[0],c={type:'bar',range:`A1:${address(1,rows[0].length-1)}`,legend:true};
  assert.equal(chartSeriesBy(c,s,b),'rows');const data=chartData(c,s,b);assert.equal(data.length,1);assert.deepEqual(data[0].values,[120,90,150,80]);assert.deepEqual(data[0].labels,['커피','차','주스','물']);
  const svg=renderChartSVG(c,s,b).svg;assert.doesNotMatch(svg,/chart-legend-item/);assert.equal((svg.match(/class="chart-category"/g)||[]).length,4);
 }
});
test('multiple observations and deliberate column series retain distinct series and legends',()=>{
 const b=fixture([['매출','비용'],[120,90],[150,80]]),s=b.sheets[0],c={type:'bar',range:'A1:B3',legend:true};assert.equal(chartSeriesBy(c,s,b),'columns');assert.equal(chartData(c,s,b).length,2);assert.match(renderChartSVG(c,s,b).svg,/chart-legend-item/);
 const small=fixture([['매출','비용'],[120,90]]),explicit={type:'bar',range:'A1:B2',seriesBy:'columns',legend:true};assert.equal(chartData(explicit,small.sheets[0],small).length,2);assert.match(renderChartSVG(explicit,small.sheets[0],small).svg,/chart-legend-item/);
});
test('single-series column exports omit redundant legends and preserve category references',async()=>{
 const b=fixture([['커피','차','주스','물'],[120,90,150,80]]),s=b.sheets[0],c={id:'clean',type:'bar',range:'A1:D2',legend:true};s.charts=[c];
 const bytes=(await writeSpreadsheet(b)).data,zip=await JSZip.loadAsync(bytes),xml=await zip.file('xl/charts/sheetStudioChart1.xml').async('string');assert.doesNotMatch(xml,/<c:legend>/);assert.match(xml,/<c:varyColors val="0"/);assert.match(xml,/\$A\$1:\$D\$1/);assert.match(xml,/\$A\$2:\$D\$2/);
 const loaded=await readSpreadsheet(bytes,'bar.xlsx'),chart=loaded.sheets[0].charts[0];assert.equal(chart.legend,false);assert.deepEqual(chartData(chart,loaded.sheets[0],loaded)[0].labels,['커피','차','주스','물']);assert.deepEqual(chartData(chart,loaded.sheets[0],loaded)[0].values,[120,90,150,80]);
 const native=fixture([['커피','차','주스','물'],[120,90,150,80]]);attachNativeSource(native,(await writeSpreadsheet(native)).data,'native.xlsm');native.sheets[0].charts=[c];const plan=buildNativePlan(native);assert.equal(plan.sheets[0].charts[0].legend,false);assert.equal(plan.sheets[0].charts[0].series.length,1);
});
