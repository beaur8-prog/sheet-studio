import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {planValidationChange,applyValidationChange} from '../src/data-validation.mjs';
import {createWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {bindPivotTable} from '../src/pivot-source.mjs';
import {applyPivot} from '../src/pivot.mjs';
import {createTablePlan,prepareTableHeaderEdits,applyTableHeaderEdits} from '../src/tables.mjs';
import {prepareTableDataEdits} from '../src/table-data-edits.mjs';
import {fixture as pivotFixture,spec} from './fixtures/pivot.mjs';
import {chartFixture,addChart,chartCases} from './fixtures/charts.mjs';
import {chartSeriesPlan} from '../src/chart-series-editor.mjs';
import {worksheetFixture} from './fixtures/worksheet.mjs';

const output=path.resolve(process.argv[2]||'test-results/ooxml-sdk');await fs.mkdir(output,{recursive:true});
async function save(name,book){for(let round=0;round<3;round++){const {data}=await writeSpreadsheet(book);await fs.writeFile(path.join(output,`${name}-${round}.xlsx`),data);book=await readSpreadsheet(data,`${name}.xlsx`);}}
for(const date1904 of [false,true]){
 const book=createWorkbook();book.date1904=date1904;book.sheets[0].cells={A1:{raw:'12.5%',style:{format:'percent',numberFormat:'0.0%'}},B1:{raw:'',style:{format:'date',numberFormat:'yyyy-mm-dd'},note:'빈 날짜 유지'},C1:{raw:'0',style:{format:'date',numberFormat:'yyyy-mm-dd'}},D1:{raw:'=ISBLANK(B1)'},E1:{raw:'=1/0'},F1:{raw:"'#N/A"}};await save(date1904?'values-1904':'values-1900',book);
}
{
 const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'Qty'},B1:{raw:'Price'},A2:{raw:'2'},B2:{raw:'3'},A3:{raw:'4'},B3:{raw:'5'}};
 const created=createTablePlan(book,sheet,{name:'Sales',range:'A1:B3'});sheet.tables=[created.table];for(const u of created.updates)sheet.cells[u.key]=u.cell;
 const plan=prepareTableDataEdits(book,sheet,prepareTableHeaderEdits(book,sheet,[{key:'C2',cell:{raw:'=[@Qty]*[@Price]'}}]));applyTableHeaderEdits(book,sheet,plan);for(const u of plan.updates)sheet.cells[u.key]=u.cell;await save('expanded-table',book);
}
for(const filtered of [false,true]){
 const book=pivotFixture(),target={id:'pivot',name:'피벗',charts:[]};applyPivot(book,target,filtered?{...spec(),rows:['상품'],columns:[],filters:[{field:'지역',values:['동부']}]}:spec());book.sheets.push(target);target.cells.B2.style={format:'number',numberFormat:'0.000',color:'#224466'};await save(filtered?'filtered-pivot':'app-pivot',book);
 if(!filtered){const zip=await JSZip.loadAsync((await writeSpreadsheet(book)).data);zip.remove('customXml/sheetStudioPivots.xml');const part='xl/pivotTables/pivotTable1.xml';zip.file(part,(await zip.file(part).async('string')).replace('subtotal="sum"','subtotal="average"'));const cells='xl/worksheets/sheet2.xml';zip.file(cells,(await zip.file(cells).async('string')).replace(/(<c r="D4"[^>]*>[\s\S]*?<v>)160(<\/v>)/,(_,a,b)=>a+'40'+b));const native=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'external.xlsx');assert.equal(native.sheets[1].cells.D4.raw,'40');await save('native-average-pivot',native);const newPivot={id:'pivot2',name:'다른 피벗',charts:[]};applyPivot(native,newPivot,{...spec(),sourceSheetId:native.sheets[0].id});native.sheets.push(newPivot);await save('mixed-native-app-pivot',native);}
}
{
 const book=chartFixture();chartCases.forEach((options,i)=>addChart(book,{...options,id:'chart-'+i,range:['bubble','scatter'].includes(options.type)?'B1:D5':'A1:C5',blankCells:options.type==='line'?'span':'gap'}));await save('chart-18-variants',book);
}
for(const policy of ['gap','zero','span']){
 const book=chartFixture();delete book.sheets[0].cells.B3;book.sheets[0].cells.B4={raw:'=NA()'};addChart(book,{type:'line',blankCells:policy});await save('chart-blanks-'+policy,book);
}
await save('worksheet-formats',worksheetFixture());console.log('Generated synthetic OOXML validation fixtures: '+output);

{
 const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={H1:{raw:'5'}};
 const definitions=[['A1:A10',{type:'list',source:'예,아니오',showInput:true,promptTitle:'승인 여부',prompt:'목록에서 선택하세요.',severity:'warning'}],['B1:B10',{type:'whole',formulae:['1','10']}],['C1:C10',{type:'decimal',operator:'greaterThan',formulae:['=$H$1']}],['D1:D10',{type:'date',formulae:['2026-01-01','2026-12-31']}],['E1:E10',{type:'time',formulae:['09:00','18:00']}],['F1:F10',{type:'textLength',operator:'lessThanOrEqual',formulae:['10']}],['G1:G10',{type:'custom',formulae:['G1>$H$1'],severity:'information'}]];
 for(const [range,input] of definitions)applyValidationChange(sheet,planValidationChange(book,sheet,range,input));applyValidationChange(sheet,planValidationChange(book,sheet,'B2:B3',null));await save('authored-validation',book);
}

{
 const book=createWorkbook(),sheet=book.sheets[0];for(const col of ['A','B'])[-10,0,10,'',"'10",'TRUE','#N/A','=""','=5+0'].forEach((raw,i)=>sheet.cells[col+(i+1)]={raw:String(raw)});sheet.conditionalRules=[{id:'scale',range:'A1:A9',type:'colorScale',formulae:[],style:{},cfvo:[{type:'min'},{type:'percent',value:50},{type:'max'}],colors:['#ff0000','#ffffff','#00ff00']},{id:'bar',range:'B1:B9',type:'dataBar',formulae:[],style:{},cfvo:[{type:'min'},{type:'max'}],barColor:'#638ec6'}];await save('visual-conditional',book);
}

{
 const source=new ExcelJS.Workbook(),s=source.addWorksheet('Advanced');s.getCell('A1').value=-10;s.getCell('A2').value=20;s.addConditionalFormatting({ref:'A1:A2',rules:[{type:'dataBar',priority:1,cfvo:[{type:'min'},{type:'max'}],color:{argb:'FF638EC6'},gradient:false,negativeFillColor:{argb:'FFFF0000'},axisColor:{argb:'FF000000'}}]});await save('preserved-advanced-bar',await readSpreadsheet(await source.xlsx.writeBuffer(),'advanced.xlsx'));
}

{
 const book=pivotFixture(),source=book.sheets[0],created=createTablePlan(book,source,{name:'Sales',range:'A1:D6'});source.tables=[created.table];for(const u of created.updates)source.cells[u.key]=u.cell;const target={id:'linked-report',name:'표 피벗',charts:[]};applyPivot(book,target,bindPivotTable(book,spec(),created.table.id));book.sheets.push(target);await save('table-linked-pivot',book);
}

{
 const book=chartFixture(),sheet=book.sheets[0],rows=[{name:'=B1',values:'B2:B5',categories:'',color:'#4472c4',lines:true},{name:'점 계열',values:'C2:C5',categories:'',color:'#e15759',lines:false}];
 addChart(book,{id:'explicit-scatter',type:'scatter',blankCells:'span',...chartSeriesPlan(rows,'scatter',book,sheet.id)});await save('explicit-scatter-groups',book);
 const combo=chartFixture(),s=combo.sheets[0];addChart(combo,{id:'explicit-combo',type:'combo',...chartSeriesPlan(rows.map((row,i)=>({...row,categories:'A2:A5',type:i?'line':'bar',secondary:!!i})),'combo',combo,s.id)});await save('explicit-combo-series',combo);
}

{
 const book=chartFixture(),sheet=book.sheets[0];addChart(book,{id:'original-33',type:'bar',series:Array.from({length:33},(_,i)=>({name:'계열 '+(i+1),type:'bar',color:'#4472c4',valRef:{sheetId:sheet.id,range:'B2:B5'},catRef:{sheetId:sheet.id,range:'A2:A5'}}))});await save('preserved-33-series-chart',book);
}

{
 const book=createWorkbook(),sheet=book.sheets[0];sheet.cells.A1={raw:'자유 메모'};applyValidationChange(sheet,planValidationChange(book,sheet,'A1:A10',{type:'any',showInput:true,promptTitle:'입력 안내',prompt:'형식 제한 없이 입력하세요.'}));sheet.xlsxFeatures.dataValidations=sheet.xlsxFeatures.dataValidations.replace('type="none"','type="any"');await save('prompt-only-validation',book);
}
