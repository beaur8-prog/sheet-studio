import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {applyPivot} from '../src/pivot.mjs';
import {createTablePlan,prepareTableHeaderEdits,applyTableHeaderEdits} from '../src/tables.mjs';
import {prepareTableDataEdits} from '../src/table-data-edits.mjs';
import {fixture as pivotFixture,spec} from './fixtures/pivot.mjs';
import {chartFixture,addChart,chartCases} from './fixtures/charts.mjs';
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
