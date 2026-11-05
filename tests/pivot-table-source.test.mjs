import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createWorkbook,validateWorkbook} from '../src/engine.mjs';
import {bindPivotTable,resolvePivotSpec} from '../src/pivot-source.mjs';
import {applyPivot,buildPivot} from '../src/pivot.mjs';
import {validatePivot} from '../src/pivot-schema.mjs';
import {createTablePlan,prepareTableHeaderEdits,applyTableHeaderEdits,appendTableColumnsPlan} from '../src/tables.mjs';
import {prepareTableDataEdits} from '../src/table-data-edits.mjs';
import {structuralWorkbookPlan} from '../src/workbook-structure.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,all,attr} from '../src/chart-xml.mjs';
function fixture(){
  const book=createWorkbook(),sheet=book.sheets[0];sheet.cells={A1:{raw:'지역'},B1:{raw:'매출'},C1:{raw:'수량'},A2:{raw:'동부'},B2:{raw:'10'},C2:{raw:'1'},A3:{raw:'서부'},B3:{raw:'20'},C3:{raw:'2'}};
  const plan=createTablePlan(book,sheet,{name:'Sales',range:'A1:C3'});sheet.tables=[plan.table];for(const u of plan.updates)sheet.cells[u.key]=u.cell;
  const spec=bindPivotTable(book,{version:1,sourceSheetId:sheet.id,sourceRange:'A1:C3',rows:['지역'],columns:[],values:[{field:'매출',aggregate:'sum'}],filters:[],totals:true,sort:'asc',preserveFormatting:true},plan.table.id),target={id:'report',name:'피벗',charts:[]};applyPivot(book,target,spec);book.sheets.push(target);return {book,sheet,target,spec,table:sheet.tables[0]};
}
test('a table-linked pivot follows row and column growth while an ordinary range stays fixed',()=>{
  const {book,sheet,spec,table}=fixture(),fixed={...spec};for(const key of ['sourceTableId','sourceTableName','sourceFields'])delete fixed[key];
  const plan=prepareTableDataEdits(book,sheet,prepareTableHeaderEdits(book,sheet,[{key:'A4',cell:{raw:'동부'}},{key:'B4',cell:{raw:'30'}},{key:'C4',cell:{raw:'3'}}]));applyTableHeaderEdits(book,sheet,plan);for(const u of plan.updates)sheet.cells[u.key]=u.cell;
  assert.equal(sheet.tables[0].range,'A1:C4');assert.equal(buildPivot(book,spec).matrix.at(-1)[1],60);assert.equal(buildPivot(book,fixed).matrix.at(-1)[1],30);assert.equal(buildPivot(book,spec).spec.sourceRange,'A1:C4');
  const right=appendTableColumnsPlan(book,sheet,sheet.tables[0],3,new Map([['D1',{raw:'새 항목'}],['D2',{raw:'4'}]]));sheet.tables[0]=right.table;for(const u of right.updates)sheet.cells[u.key]=u.cell;assert.equal(buildPivot(book,spec).data.headers.length,4);assert.equal(table.id,spec.sourceTableId);
});
test('column and table renames follow stable IDs and preserve report formats with renamed value headers',()=>{
  const {book,sheet,target,spec}=fixture();target.cells.B2.style={numberFormat:'0.000',format:'number',color:'#aa2255'};target.columnWidths[1]=230;
  const plan=prepareTableHeaderEdits(book,sheet,[{key:'B1',cell:{raw:'판매 금액'}}]);applyTableHeaderEdits(book,sheet,plan);for(const u of plan.updates)sheet.cells[u.key]=u.cell;sheet.tables[0].name='Revenue';
  const next=resolvePivotSpec(book,spec);assert.equal(next.values[0].field,'판매 금액');assert.equal(next.sourceTableName,'Revenue');assert.equal(spec.values[0].field,'매출','resolving cannot change the saved spec');
  applyPivot(book,target,spec);assert.equal(target.pivot.values[0].field,'판매 금액');assert.equal(target.cells.B2.style.numberFormat,'0.000');assert.equal(target.cells.B2.style.color,'#aa2255');assert.equal(target.columnWidths[1],230);
});
test('totals rows are excluded, headerless or empty source tables are rejected, and deleted tables never retarget to same-named replacements',()=>{
  const {book,sheet,spec}=fixture();sheet.tables[0].totalsRow=true;sheet.tables[0].range='A1:C4';sheet.cells.A4={raw:'총합계'};sheet.cells.B4={raw:'999'};assert.equal(buildPivot(book,spec).matrix.at(-1)[1],30);
  sheet.tables[0].headerRow=false;assert.throws(()=>buildPivot(book,spec),/머리글/);sheet.tables[0].headerRow=true;sheet.tables[0].range='A1:C2';sheet.tables[0].totalsRow=true;assert.throws(()=>buildPivot(book,spec),/데이터 행/);
  sheet.tables[0].id='new-table';sheet.tables[0].range='A1:C3';sheet.tables[0].totalsRow=false;assert.throws(()=>buildPivot(book,spec),/원본 표가 없습니다/);
});
test('deleting a used final column and adding another cannot recycle its ID; failed refresh/export preserve the last good result',async()=>{
  let {book,sheet,target,table}=fixture();const spec=bindPivotTable(book,{...target.pivot,values:[{field:'수량',aggregate:'sum'}]},table.id);applyPivot(book,target,spec);const before=JSON.stringify(target);
  book=structuralWorkbookPlan(book,sheet.id,{axis:'c',index:2,count:1,delta:-1});sheet=book.sheets[0];target=book.sheets[1];const added=appendTableColumnsPlan(book,sheet,sheet.tables[0],2,new Map([['C1',{raw:'수량'}],['C2',{raw:'100'}]]));sheet.tables[0]=added.table;for(const u of added.updates)sheet.cells[u.key]=u.cell;
  assert.equal(sheet.tables[0].columns[2].id,4);const beforeRefresh=JSON.stringify(target);assert.throws(()=>applyPivot(book,target,target.pivot),/수량.*削|수량.*삭제/);assert.equal(JSON.stringify(target),beforeRefresh);await assert.rejects(writeSpreadsheet(book),/수량.*삭제/);assert.equal(JSON.stringify(target),beforeRefresh);
});
test('linked source schema requires bounded unique field IDs and a complete mapping',()=>{
  const {spec}=fixture();for(const change of [{sourceFields:[]},{sourceFields:[{id:1,name:'지역'},{id:1,name:'매출'}]},{sourceFields:[{id:-1,name:'지역'}]},{sourceTableName:''}])assert.throws(()=>validatePivot({...spec,...change}));
});
test('native XLSX uses a named table source and safely rebinds new in-memory table IDs through three roundtrips',async()=>{
  let {book}=fixture();for(let round=0;round<3;round++){
    const data=(await writeSpreadsheet(book)).data,zip=await JSZip.loadAsync(data),definition=xml(await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string')),source=all(definition,'worksheetSource')[0];assert.equal(attr(source,'name'),'Sales');assert.equal(attr(source,'ref'),'');
    book=validateWorkbook(await readSpreadsheet(data,'linked.xlsx'));const report=book.sheets[1],table=book.sheets[0].tables[0];assert.ok(report.pivot);assert.equal(report.pivot.sourceTableId,table.id);assert.equal(report.cells.B4.raw,'30');assert.equal(book.xlsxPivotArchive,undefined);assert.equal(buildPivot(book,report.pivot).matrix.at(-1)[1],30);
  }
});
