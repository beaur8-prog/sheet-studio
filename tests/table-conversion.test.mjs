import test from 'node:test';
import assert from 'node:assert/strict';
import {createEvaluator,validateWorkbook} from '../src/engine.mjs';
import {tableConversionPlan} from '../src/table-conversion.mjs';
import {ensureTablePreset} from '../src/table-style.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {tableFixture} from './tables-fixture.mjs';
function apply(book,plan){for(const {sheetId,key,cell} of plan.updates)book.sheets.find(s=>s.id===sheetId).cells[key]=cell;for(const {sheetId,...metadata} of plan.metadata)Object.assign(book.sheets.find(s=>s.id===sheetId),metadata);if(plan.chartNames)book.chartNames=plan.chartNames;book.referenceCachesValid=false;}
test('table conversion materializes in-row, whole-column, named and cross-sheet references without changing the source while planning',async()=>{
  let book=await readSpreadsheet(await tableFixture(),'convert.xlsx'),s=book.sheets[0],t=s.tables[0];t.style.name=ensureTablePreset(book,'green');book.chartNames=[{name:'Revenue',formula:'SUM(Sales[Amount])'}];s.cells.E2={raw:'=SUM(Revenue)'};s.cells.C2.style={color:'#ff0000'};
  const second={id:'second',name:'Another',cells:{A1:{raw:"=SUM('Sales Data'!Sales[Amount])"},A2:{raw:'=SUM(Sales)'}},columnWidths:{},rowHeights:{},charts:[]};book.sheets.push(second);const before=JSON.stringify(book),plan=tableConversionPlan(book,s,t);assert.equal(JSON.stringify(book),before);apply(book,plan);validateWorkbook(book);
  assert.equal(s.tables.length,0);assert.equal(s.cells.C2.raw,'=A2*B2');assert.equal(s.cells.C3.raw,'=A3*B3');assert.equal(createEvaluator(book,s.id)('E1'),26);assert.equal(createEvaluator(book,s.id)('E2'),26);assert.equal(createEvaluator(book,s.id)('C4'),26);assert.equal(createEvaluator(book,second.id)('A1'),26);assert.equal(createEvaluator(book,second.id)('A2'),40);
  assert.equal(s.cells.A1.style.background,'#107c41');assert.equal(s.cells.C2.style.color,'#ff0000');assert.equal(s.cells.C2.style.background,'#e2efda');assert.match(book.chartNames[0].formula,/\$C\$2:\$C\$3/);
  book=await readSpreadsheet((await writeSpreadsheet(book)).data,'converted-roundtrip.xlsx');s=book.sheets[0];assert.equal(s.tables?.length||0,0);assert.equal(createEvaluator(book,s.id)('E1'),26);assert.equal(s.cells.A1.style.background.toLowerCase(),'#107c41');assert.equal(s.cells.C2.style.color.toLowerCase(),'#ff0000');
});
test('conversion retains filtered visibility and array/future formula source attributes while rewriting their references',async()=>{
  const book=await readSpreadsheet(await tableFixture(),'visibility.xlsx'),s=book.sheets[0],t=s.tables[0];t.filter.hiddenRows=[2];s.cells.E1.xlsxFormula={raw:'=SUM(Sales[Amount])',formula:'SUM(Sales[Amount])',shareType:'array',ref:'E1:E1'};const plan=tableConversionPlan(book,s,t);apply(book,plan);assert.equal(s.rowHeights[2],0);assert.match(s.cells.E1.xlsxFormula.raw,/\$C\$2/);assert.equal(s.cells.E1.xlsxFormula.shareType,'array');assert.equal(createEvaluator(book,s.id)('C4'),6);
});
