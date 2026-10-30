import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createWorkbook,createEvaluator,evaluateFormula,validateWorkbook} from '../src/engine.mjs';
import {createHistory} from '../src/history.mjs';
import {createTablePlan,updateTablePlan,calculatedColumnPlan,renameTableReferences,prepareTableHeaderEdits,applyTableHeaderEdits,finalizeTablePlan} from '../src/tables.mjs';
import {tableCellStyle,ensureTablePreset} from '../src/table-style.mjs';
import {tableFixture} from './tables-fixture.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,all,attr,child,children} from '../src/chart-xml.mjs';

function apply(sheet,plan){sheet.tables||=[];if(plan.create)sheet.tables.push(plan.table);else sheet.tables[sheet.tables.findIndex(t=>t.id===plan.table.id)]=plan.table;for(const update of plan.updates)sheet.cells[update.key]=update.cell;}
test('table creation normalizes text headers and rejects conflicting ranges, names and merged cells',()=>{
  const book=createWorkbook(),s=book.sheets[0];s.cells={A1:{raw:'0'},B1:{raw:'FALSE'},C1:{raw:'0'},A2:{raw:'1'},B2:{raw:'2'},C2:{raw:'3'}};
  const plan=createTablePlan(book,s,{range:'A1:C2',name:'Sales',headerRow:true});assert.deepEqual(plan.table.columns.map(c=>c.name),['0','FALSE','02']);apply(s,plan);assert.equal(s.cells.A1.raw,"'0");assert.equal(s.cells.B1.raw,"'FALSE");validateWorkbook(book);
  assert.throws(()=>createTablePlan(book,s,{range:'B1:D2',name:'Another'}),/겹칩/);assert.throws(()=>createTablePlan(book,s,{range:'F1:H2',name:'sales'}),/사용/);s.merges=['F1:G1'];assert.throws(()=>createTablePlan(book,s,{range:'F1:H2',name:'Other'}),/병합/);
  const plain=createTablePlan(book,s,{range:'J1:K2',name:'Plain',headerRow:false});assert.equal(plain.updates.length,0);assert.equal(plain.table.filter,undefined);assert.deepEqual(plain.table.columns.map(c=>c.name),['열1','열2']);
});
test('table rename and header edits update cells, calculated definitions, names and chart references in one history entry',async()=>{
  const source=await readSpreadsheet(await tableFixture(),'edit.xlsx');source.chartNames=[{name:'TotalSales',formula:'SUM(Sales[Amount])'}];source.sheets[0].charts=[{id:'chart',type:'line',range:'A1:C3',width:500,height:300,series:[{valRef:{formula:'Sales[Amount]',sheetId:source.sheets[0].id,range:'C2:C3'}}]}];
  const history=createHistory(source),book=history.book,s=book.sheets[0],t=s.tables[0];history.begin();renameTableReferences(book,s,t,{name:'Revenue'});t.name='Revenue';history.commit();assert.equal(s.cells.E1.raw,'=SUM(Revenue[Amount])');assert.equal(book.chartNames[0].formula,'SUM(Revenue[Amount])');assert.equal(s.charts[0].series[0].valRef.formula,'Revenue[Amount]');
  history.begin();const header=prepareTableHeaderEdits(book,s,[{key:'A1',cell:{...s.cells.A1,raw:'Count'}}]);applyTableHeaderEdits(book,s,header);for(const update of header.updates)s.cells[update.key]=update.cell;history.commit();assert.equal(s.tables[0].columns[0].name,'Count');assert.match(s.cells.C2.raw,/@\[Count\]/);assert.equal(createEvaluator(book,s.id)('E1'),26);
  assert.ok(history.undo());assert.equal(book.sheets[0].tables[0].columns[0].name,'Qty');assert.equal(createEvaluator(book,s.id)('E1'),26);assert.ok(history.undo(true));assert.equal(book.sheets[0].tables[0].columns[0].name,'Count');
  const duplicate=prepareTableHeaderEdits(book,s,[{key:'A1',cell:{raw:'Price'}}]);assert.equal(duplicate.tables[0].table.columns[0].name,'Price2');
  const loaded=await readSpreadsheet((await writeSpreadsheet(book)).data,'renamed.xlsx');assert.equal(loaded.sheets[0].tables[0].name,'Revenue');assert.equal(createEvaluator(loaded,loaded.activeSheetId)('C3'),20);
});
test('calculated columns, totals toggles and row resizing preserve results outside a shrunken table',async()=>{
  const book=await readSpreadsheet(await tableFixture(),'calc.xlsx'),s=book.sheets[0];let t=s.tables[0];
  apply(s,calculatedColumnPlan(book,s,t,{column:2,formula:'=[@Qty]*[@Price]+1'}));t=s.tables[0];assert.equal(createEvaluator(book,s.id)('C2'),7);assert.equal(createEvaluator(book,s.id)('C4'),28);
  apply(s,updateTablePlan(book,s,t,{range:'A1:C3',totalsRow:false}));t=s.tables[0];assert.equal(s.cells.C4.raw,'');assert.equal(createEvaluator(book,s.id)('E1'),28);
  apply(s,updateTablePlan(book,s,t,{range:'A1:C2',totalsRow:false}));t=s.tables[0];assert.equal(s.cells.C3.raw,'=A3*B3+1');assert.equal(createEvaluator(book,s.id)('C3'),21);assert.equal(createEvaluator(book,s.id)('E1'),7);
  s.cells.A4={raw:'3'};s.cells.B4={raw:'8'};apply(s,updateTablePlan(book,s,t,{range:'A1:C4',totalsRow:false}));t=s.tables[0];assert.equal(createEvaluator(book,s.id)('C4'),25);assert.equal(createEvaluator(book,s.id)('E1'),53);
  apply(s,updateTablePlan(book,s,t,{range:'A1:C5',totalsRow:true}));assert.equal(createEvaluator(book,s.id)('C5'),53);assert.equal(s.cells.B5.raw,'');assert.equal(s.cells.A5.raw,'합계');validateWorkbook(book);
});
test('custom table presets retain appearance, explicit formatting precedence and stable style counts on repeated saves',async()=>{
  let book=createWorkbook(),s=book.sheets[0];s.cells={A1:{raw:'Item'},B1:{raw:'Value'},A2:{raw:'a'},B2:{raw:'3'},A3:{raw:'b'},B3:{raw:'4'}};const plan=createTablePlan(book,s,{range:'A1:B3',name:'Styled'});plan.table.style.name=ensureTablePreset(book,'green');apply(s,plan);let counts;
  for(let round=0;round<4;round++){
    assert.equal(tableCellStyle(book,s,0,0).background,'#107c41');assert.equal(tableCellStyle(book,s,0,0).color,'#ffffff');assert.equal(tableCellStyle(book,s,1,0).background,'#e2efda');assert.equal(ensureTablePreset(book,'green'),s.tables[0].style.name);
    const bytes=(await writeSpreadsheet(book)).data,zip=await JSZip.loadAsync(bytes),doc=xml(await zip.file('xl/styles.xml').async('string')),current=[children(child(doc.documentElement,'dxfs'),'dxf').length,all(doc,'tableStyle').length];if(!counts)counts=current;else assert.deepEqual(current,counts);
    book=validateWorkbook(await readSpreadsheet(bytes,'styled.xlsx'));s=book.sheets[0];assert.equal(s.cells.A1.style.color,undefined);assert.equal(s.cells.A1.style.bold,undefined);
  }
  s.cells.A1.style={color:'#ff0000'};assert.equal({...tableCellStyle(book,s,0,0),...s.cells.A1.style}.color,'#ff0000');
});
test('table planning does not mutate the current workbook and reapplies filters against pending calculated cells',async()=>{
  const book=await readSpreadsheet(await tableFixture(),'plan.xlsx'),s=book.sheets[0],t=s.tables[0];t.filter.columns[2]={type:'custom',kind:'number',join:'and',conditions:[{operator:'greaterThan',value:'10'}]};
  const before=JSON.stringify(book),plan=finalizeTablePlan(book,s,calculatedColumnPlan(book,s,t,{column:2,formula:'=[@Qty]*[@Price]+10'}));assert.equal(JSON.stringify(book),before);assert.deepEqual(plan.table.filter.hiddenRows,[]);apply(s,plan);assert.equal(createEvaluator(book,s.id)('C4'),46);
  assert.throws(()=>calculatedColumnPlan(book,s,s.tables[0],{column:2,formula:'1+2'}),/=/);
});

test('changing the preset reapplies table color filters against the new visible style',()=>{
  const book=createWorkbook(),s=book.sheets[0];s.cells={A1:{raw:'Name'},A2:{raw:'a'},A3:{raw:'b'}};let plan=createTablePlan(book,s,{range:'A1:A3',name:'Colors'});plan.table.style.name=ensureTablePreset(book,'green');apply(s,plan);const table=s.tables[0];table.filter.columns[0]={type:'color',target:'background',color:'#e2efda'};plan=finalizeTablePlan(book,s,updateTablePlan(book,s,table,{}));assert.deepEqual(plan.table.filter.hiddenRows,[2]);
  const changed=finalizeTablePlan(book,s,updateTablePlan(book,s,table,{}),{preset:'blue'});assert.deepEqual(changed.table.filter.hiddenRows,[1,2]);assert.equal(book.xlsxTableStyles.length,1);
});
