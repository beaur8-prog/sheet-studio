import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createWorkbook,validateWorkbook} from '../src/engine.mjs';
import {chartSeriesPlan,chartSeriesRows,chartSeriesEditorSeed} from '../src/chart-series-editor.mjs';
import {chartData,chartRangeError} from '../src/chart-model.mjs';
import {buildChartPreview} from '../src/chart-preview.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';
import {structuralWorkbookPlan} from '../src/workbook-structure.mjs';
import {createTablePlan} from '../src/tables.mjs';
import {xml,all,attr,child,children,stringify} from '../src/chart-xml.mjs';
function fixture(){const book=createWorkbook(),sheet=book.sheets[0];sheet.name="매출 '24";sheet.cells={A1:{raw:'항목'},B1:{raw:'실적'},D1:{raw:'목표'}};for(let i=0;i<4;i++){sheet.cells['A'+(i+2)]={raw:['가','나','다','라'][i]};sheet.cells['B'+(i+2)]={raw:String((i+1)*10)};sheet.cells['D'+(i+2)]={raw:String((i+1)*12)};sheet.cells['F'+(i+2)]={raw:String(i+1)};}const row=(values,name)=>({name,values,categories:'A2:A5',color:'#4472c4',type:'bar',secondary:false,lines:false});return {book,sheet,row};}
test('explicit non-adjacent series preserve names, references, colors, ordering and the four selected values',()=>{
  const {book,sheet,row}=fixture(),plan=chartSeriesPlan([row('D2:D5','목표'),row('B2:B5','=B1')],'bar',book,sheet.id),chart={id:'chart',type:'bar',...plan};const data=chartData(chart,sheet,book);assert.deepEqual(data.map(s=>s.name),['목표','실적']);assert.deepEqual(data.map(s=>s.values),[[12,24,36,48],[10,20,30,40]]);assert.deepEqual(data[0].labels,['가','나','다','라']);assert.equal(chartRangeError(chart,sheet,book),null);assert.equal(chartSeriesRows(chart,sheet,book)[1].name,"='매출 ''24'!$B$1:$B$1");
  const before=JSON.stringify(book),preview=buildChartPreview(book,sheet.id,{patch:chart});assert.equal(preview.seriesCount,2);assert.deepEqual(preview.rows.slice(0,4).map(r=>r.value),['12','24','36','48']);assert.equal(JSON.stringify(book),before);
});
test('literal series names beginning with equals or apostrophe remain literal when edited',()=>{
  const {book,sheet,row}=fixture();for(const name of ['=실적',"'목표"]){const chart={type:'bar',series:[{...chartSeriesPlan([row('B2:B5','placeholder')],'bar',book,sheet.id).series[0],name}]},rows=chartSeriesRows(chart,sheet,book),plan=chartSeriesPlan(rows,'bar',book,sheet.id);assert.equal(plan.series[0].name,name);assert.equal(plan.series[0].nameRef,undefined);}
});
test('invalid series ranges, lengths, category sources, names and colors reject before changing the document',()=>{
  const {book,sheet,row}=fixture(),before=JSON.stringify(book);for(const rows of [[],[row('B2:C5','n')],[row('B2:B4','n')],[row('B2:B5','=B1:B2')],[{...row('B2:B5','n'),color:'url(bad)'}],[row('B2:B5','n'),{...row('D2:D5','m'),categories:'F2:F5'}],[row('B2:B10002','n')]])assert.throws(()=>chartSeriesPlan(rows,'bar',book,sheet.id));assert.equal(JSON.stringify(book),before);
});
test('names and cross-sheet vectors are retained as dynamic formulas instead of silently replacing them with a rectangle',()=>{
  const {book,sheet,row}=fixture(),other={...createWorkbook().sheets[0],id:'other',name:'자료 2',cells:{A1:{raw:'선택 계열'},B2:{raw:'1'},B3:{raw:'2'},B4:{raw:'3'},B5:{raw:'4'}}};book.sheets.push(other);book.chartNames=[{name:'Values',formula:"'자료 2'!$B$2:$B$5"}];
  const plan=chartSeriesPlan([{...row('Values',"='자료 2'!A1"),categories:''}],'line',book,sheet.id);assert.equal(plan.series[0].valRef.formula,'Values');assert.deepEqual(chartData({type:'line',...plan},sheet,book)[0].values,[1,2,3,4]);assert.equal(chartData({type:'line',...plan},sheet,book)[0].name,'선택 계열');assert.deepEqual(chartData({type:'line',...plan},sheet,book)[0].labels,['1','2','3','4']);
});
test('a growing named source reports the current preview limit instead of its old resolved rectangle',()=>{
  const {book,sheet,row}=fixture();book.chartNames=[{name:'Values',formula:"'매출 ''24'!$B$2:$B$5"}];const chart={type:'bar',...chartSeriesPlan([{...row('Values','실적'),categories:''}],'bar',book,sheet.id)};assert.equal(chartData(chart,sheet,book)[0].truncated,false);for(let r=2;r<=1002;r++)sheet.cells['B'+r]={raw:String(r)};book.chartNames[0].formula="'매출 ''24'!$B$2:$B$1002";const data=chartData(chart,sheet,book)[0];assert.equal(data.values.length,1000);assert.equal(data.truncated,true);assert.equal(buildChartPreview(book,sheet.id,{patch:chart}).truncated,true);
});
test('table columns can feed a chart through defined names while direct structured chart formulas require the named route',()=>{
  const {book,sheet,row}=fixture(),plan=createTablePlan(book,sheet,{name:'Sales',range:'A1:B5'});sheet.tables=[plan.table];for(const update of plan.updates)sheet.cells[update.key]=update.cell;book.chartNames=[{name:'PlotValues',formula:'Sales[실적]'},{name:'PlotItems',formula:'Sales[항목]'}];
  assert.throws(()=>chartSeriesPlan([row('Sales[실적]','실적')],'bar',book,sheet.id),/이름으로/);const chart={type:'bar',...chartSeriesPlan([{...row('PlotValues','실적'),categories:'PlotItems'}],'bar',book,sheet.id)};assert.deepEqual(chartData(chart,sheet,book)[0].values,[10,20,30,40]);
});
test('scatter implicit X values and mixed point/line styles survive XLSX without changing their connection policy',async()=>{
  let {book,sheet,row}=fixture();const plan=chartSeriesPlan([{...row('B2:B5','연결'),categories:'',lines:true},{...row('D2:D5','점'),categories:'',lines:false}],'scatter',book,sheet.id);sheet.charts=[{id:'scatter',type:'scatter',title:'XY',x:100,y:100,width:500,height:300,blankCells:'span',...plan}];
  for(let round=0;round<3;round++){
    const data=(await writeSpreadsheet(book)).data,zip=await JSZip.loadAsync(data),types=xml(await zip.file('[Content_Types].xml').async('string')),part=all(types,'Override').find(n=>attr(n,'ContentType').endsWith('drawingml.chart+xml')),entry=zip.file(attr(part,'PartName').slice(1)),doc=xml(await entry.async('string'));assert.deepEqual(all(doc,'scatterStyle').map(n=>attr(n,'val')),['lineMarker','marker']);assert.equal(all(doc,'ser').length,2);
    book=validateWorkbook(await readSpreadsheet(data,'scatter.xlsx'));sheet=book.sheets[0];assert.equal(sheet.charts[0].type,'scatter');const points=chartData(sheet.charts[0],sheet,book);assert.deepEqual(points.map(s=>s.lines),[true,false]);assert.deepEqual(points[0].xs,[1,2,3,4]);assert.deepEqual(points[1].values,[12,24,36,48]);
  }
});
test('bubble requires matching X/Y/size vectors, and a combo keeps per-series type and secondary-axis choices',()=>{
  const {book,sheet,row}=fixture();assert.throws(()=>chartSeriesPlan([{...row('B2:B5','bubble'),size:'F2:F4'}],'bubble',book,sheet.id),/셀 수/);
  const bubble=chartSeriesPlan([{...row('B2:B5','bubble'),categories:'F2:F5',size:'D2:D5'}],'bubble',book,sheet.id);const b=chartData({type:'bubble',...bubble},sheet,book)[0];assert.deepEqual(b.xs,[1,2,3,4]);assert.deepEqual(b.sizes,[12,24,36,48]);
  const combo=chartSeriesPlan([row('B2:B5','막대'),{...row('D2:D5','선'),type:'line',secondary:true}],'combo',book,sheet.id);assert.deepEqual(combo.series.map(s=>[s.type,s.secondary]),[['bar',false],['line',true]]);
});
test('one external scatter group can override its connection style per series without becoming a combo',async()=>{
  const {book,sheet,row}=fixture();sheet.charts=[{id:'external',type:'scatter',x:100,y:100,width:500,height:300,...chartSeriesPlan([{...row('B2:B5','연결'),categories:'',lines:true},{...row('D2:D5','점'),categories:'',lines:false}],'scatter',book,sheet.id)}];
  const zip=await JSZip.loadAsync((await writeSpreadsheet(book)).data),types=xml(await zip.file('[Content_Types].xml').async('string')),part=attr(all(types,'Override').find(n=>attr(n,'ContentType').endsWith('drawingml.chart+xml')),'PartName').slice(1),doc=xml(await zip.file(part).async('string')),groups=all(doc,'scatterChart');assert.equal(groups.length,2);groups[0].insertBefore(child(groups[1],'ser'),child(groups[0],'axId'));groups[1].parentNode.removeChild(groups[1]);child(groups[0],'scatterStyle').setAttribute('val','smoothMarker');assert.ok(child(child(child(children(groups[0],'ser')[1],'spPr'),'ln'),'noFill'));zip.file(part,stringify(doc));
  let loaded=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'external-scatter.xlsx'),chart=loaded.sheets[0].charts[0];assert.equal(chart.type,'scatter');assert.deepEqual(chart.series.map(s=>s.lines),[true,false]);assert.deepEqual(chartData(chart,loaded.sheets[0],loaded).map(s=>s.xs),[[1,2,3,4],[1,2,3,4]]);
  for(const series of children(groups[0],'ser'))series.removeChild(child(series,'xVal'));zip.file(part,stringify(doc));loaded=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'implicit-scatter.xlsx');chart=loaded.sheets[0].charts[0];assert.deepEqual(chartData(chart,loaded.sheets[0],loaded).map(s=>s.xs),[[1,2,3,4],[1,2,3,4]],'data preview uses the same implicit X values as the renderer');
});
test('explicit vectors and name references follow real structural edits and survive XLSX reopen',async()=>{
  let {book,sheet,row}=fixture();sheet.charts=[{id:'bar',type:'bar',x:100,y:100,width:500,height:300,...chartSeriesPlan([row('B2:B5','=B1')],'bar',book,sheet.id)}];book=structuralWorkbookPlan(book,sheet.id,{axis:'r',index:0,count:1,delta:1});sheet=book.sheets[0];assert.equal(sheet.charts[0].series[0].valRef.range,'B3:B6');assert.equal(sheet.charts[0].series[0].nameRef.range,'B2:B2');book=await readSpreadsheet((await writeSpreadsheet(book)).data,'shifted.xlsx');assert.deepEqual(chartData(book.sheets[0].charts[0],book.sheets[0],book)[0].values,[10,20,30,40]);
});
test('original charts with more series than their limited preview cannot silently convert that partial preview',async()=>{
  let {book,sheet,row}=fixture();const prototype=chartSeriesPlan([row('B2:B5','계열')],'bar',book,sheet.id).series[0];sheet.charts=[{id:'many',type:'bar',x:100,y:100,width:500,height:300,series:Array.from({length:33},(_,i)=>({...structuredClone(prototype),name:'계열 '+(i+1)}))}];
  for(let round=0;round<3;round++){
    book=await readSpreadsheet((await writeSpreadsheet(book)).data,'many.xlsx');sheet=book.sheets[0];assert.equal(sheet.charts[0].series.length,32,'existing preview limit');const seed=chartSeriesEditorSeed(sheet.charts[0],sheet,book);assert.deepEqual(seed.rows,[]);assert.match(seed.warning,/33개 계열/);assert.throws(()=>chartSeriesPlan(seed.rows,'bar',book,sheet.id),/1~32/);
  }
  const chart=sheet.charts[0];chart.source.parts[chart.source.root].data=btoa('<broken');assert.match(chartSeriesEditorSeed(chart,sheet,book).warning,/확인할 수 없/);
});
