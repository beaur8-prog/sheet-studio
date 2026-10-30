import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {tableFixture} from './tables-fixture.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {createWorkbook,createEvaluator,evaluateFormula,formulaReference,validateWorkbook,shiftFormula,formulaForXlsx} from '../src/engine.mjs';
import {tableBounds,filteredRowsForSheet} from '../src/table-schema.mjs';
import {applyFilters} from '../src/filters.mjs';
import {xml,stringify,child,children,all,attr,readRelationships,resolvePart,relPath,relationship,relationships,R} from '../src/chart-xml.mjs';

test('structured references calculate table columns, current rows, totals, item unions and reference metadata',async()=>{
  const book=validateWorkbook(await readSpreadsheet(await tableFixture(),'tables.xlsx')),s=book.sheets[0],calc=createEvaluator(book,s.id),formula=f=>evaluateFormula(book,s.id,f);
  assert.equal(calc('C2'),6);assert.equal(calc('C3'),20);assert.equal(calc('C4'),26);assert.equal(calc('E1'),26);
  for(const [f,value] of [['=SUM(Sales[Qty])',6],['=SUM(Sales[[Qty]:[Price]])',14],['=SUM(Sales)',40],['=ROWS(Sales[#All])',4],['=SUM(Sales[[#Totals],[Amount]])',26],['=SUM(Sales[[#Headers],[#Data],[Qty]])',6],['=SUM(Sales[[Qty],[Amount]])',32],['=INDEX(Sales[Amount],2)',20],['=ROW(Sales[Amount])',2],['=COLUMNS(Sales[[Qty]:[Amount]])',3],['=SUM(INDIRECT("Sales[Amount]"))',26]])assert.equal(formula(f),value,f);
  assert.deepEqual(formulaReference(book,s.id,'Sales[Amount]'),{sheetId:s.id,range:{r1:1,r2:2,c1:2,c2:2}});
  assert.equal(formula('=Sales[@Qty]'),'#VALUE!');assert.equal(formula('=SUM(Sales[Missing])'),'#REF!');
  assert.equal(shiftFormula('=SUM(Sales[Amount])+A1',1,1),'=SUM(Sales[Amount])+B2');assert.equal(formulaForXlsx('=SUM(Sales[Amount])'),'=SUM(Sales[Amount])');
});
test('structured selectors accept spaces and escaped brackets, hash, apostrophe and at-sign headers',()=>{
  const book=createWorkbook(),s=book.sheets[0];s.tables=[{id:'t',name:'Special',range:'A1:F3',headerRow:true,totalsRow:false,columns:['Net Sales','#Count','A]B',"O'Clock",'@Flag','Part[A]'].map(name=>({name}))}];s.cells={A2:{raw:'5'},A3:{raw:'7'},B2:{raw:'1'},B3:{raw:'2'},C2:{raw:'3'},C3:{raw:'4'},D2:{raw:'8'},D3:{raw:'9'},E2:{raw:'TRUE'},E3:{raw:'FALSE'},F2:{raw:'2'},F3:{raw:'6'}};
  const calc=f=>evaluateFormula(book,s.id,f);for(const [f,value] of [['=SUM(Special[[Net Sales]])',12],["=SUM(Special['#Count])",3],["=SUM(Special[A']B])",7],["=SUM(Special[O''Clock])",17],["=COUNTIF(Special['@Flag],TRUE)",1],["=SUM(Special[Part'[A']])",8],['=SUM(Special[ [Net Sales]: [A\']B] ])',22]])assert.equal(calc(f),value,f);
});
test('table identity, columns, formulas, totals and changed data survive repeated XLSX roundtrips',async()=>{
  let bytes=await tableFixture();
  for(let round=0;round<3;round++){
    const book=validateWorkbook(await readSpreadsheet(bytes,'tables.xlsx')),s=book.sheets[0],t=s.tables[0];assert.equal(t.name,'Sales');assert.equal(t.range,'A1:C4');assert.equal(t.columns[2].formula,round?'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]':'[@Qty]*[@Price]');assert.equal(t.columns[2].totalsFunction,'sum');assert.equal(t.style.showRowStripes,true);assert.equal(t.filter.unsupported,undefined);
    s.cells.A3.raw='6';bytes=(await writeSpreadsheet(book)).data;const excel=new ExcelJS.Workbook();await excel.xlsx.load(bytes);const out=excel.worksheets[0];assert.equal(out.getTable('Sales').table.tableRef,'A1:C4');assert.equal(out.getCell('C3').result,30);assert.equal(out.getCell('C4').result,36);assert.equal(out.getCell('E1').result,36);assert.equal(out.getCell('C2').formula,'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]');
  }
});
test('filtered tables reopen without the ExcelJS table-filter parser and retain manual row hiding',async()=>{
  let book=await readSpreadsheet(await tableFixture(),'tables.xlsx'),s=book.sheets[0],t=s.tables[0];t.filter.columns[1]={type:'custom',kind:'number',join:'and',conditions:[{operator:'greaterThan',value:'3'}]};applyFilters(book,{...s,autoFilter:t.filter});s.rowHeights[2]=0;
  for(let round=0;round<3;round++){
    book=validateWorkbook(await readSpreadsheet((await writeSpreadsheet(book)).data,'filtered.xlsx'));s=book.sheets[0];t=s.tables[0];assert.deepEqual(t.filter.hiddenRows,[1]);assert.equal(s.rowHeights[1],undefined);assert.equal(s.rowHeights[2],0);assert.deepEqual([...filteredRowsForSheet(s)].sort(),[1]);
    assert.equal(evaluateFormula(book,s.id,'=SUBTOTAL(9,Sales[Price])'),5);assert.equal(evaluateFormula(book,s.id,'=SUBTOTAL(109,Sales[Price])'),0);
  }
  t.filter.columns={};applyFilters(book,{...s,autoFilter:t.filter});assert.deepEqual([...filteredRowsForSheet(s)],[]);
});
test('custom table styles and differential format IDs remain consistent without growth',async()=>{
  const zip=await JSZip.loadAsync(await tableFixture()),styles=xml(await zip.file('xl/styles.xml').async('string')),NS=styles.documentElement.namespaceURI;let dxfs=child(styles.documentElement,'dxfs');if(!dxfs){dxfs=styles.createElementNS(NS,'dxfs');styles.documentElement.insertBefore(dxfs,child(styles.documentElement,'tableStyles'));}const red=xml(`<dxf xmlns="${NS}"><font><color rgb="FFFF0000"/></font></dxf>`).documentElement;dxfs.appendChild(styles.importNode(red,true));dxfs.setAttribute('count','1');const custom=xml(`<tableStyle xmlns="${NS}" name="StudioCustom" pivot="0" table="1" count="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle>`).documentElement,tableStyles=child(styles.documentElement,'tableStyles');tableStyles.appendChild(styles.importNode(custom,true));tableStyles.setAttribute('count','1');zip.file('xl/styles.xml',stringify(styles));
  const table=xml(await zip.file('xl/tables/table1.xml').async('string'));table.documentElement.setAttribute('headerRowDxfId','0');child(table.documentElement,'tableStyleInfo').setAttribute('name','StudioCustom');zip.file('xl/tables/table1.xml',stringify(table));let bytes=await zip.generateAsync({type:'uint8array'}),count;
  for(let round=0;round<4;round++){
    const book=validateWorkbook(await readSpreadsheet(bytes,'custom.xlsx')),s=book.sheets[0];s.conditionalRules=[{id:'green',range:'E1:E1',type:'cellIs',operator:'greaterThan',formulae:['0'],style:{color:'#008800'}}];bytes=(await writeSpreadsheet(book)).data;
    const output=await JSZip.loadAsync(bytes),outStyles=xml(await output.file('xl/styles.xml').async('string')),outTable=xml(await output.file('xl/tables/table1.xml').async('string')),id=Number(attr(outTable.documentElement,'headerRowDxfId')),items=children(child(outStyles.documentElement,'dxfs'),'dxf');assert.equal(attr(all(items[id],'color')[0],'rgb'),'FFFF0000');const style=all(outStyles,'tableStyle').find(n=>attr(n,'name')==='StudioCustom');assert.equal(Number(attr(all(style,'tableStyleElement')[0],'dxfId')),id);if(count===undefined)count=items.length;else assert.equal(items.length,count);
  }
});
test('query-table relationships and workbook connections are retained with stable package paths',async()=>{
  const zip=await JSZip.loadAsync(await tableFixture()),NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';zip.file('xl/tables/_rels/table1.xml.rels',relationships(relationship('rIdQuery',R+'/queryTable','../queryTables/queryTable1.xml')));zip.file('xl/queryTables/queryTable1.xml',`<queryTable xmlns="${NS}" name="ExternalSales" connectionId="1"><queryTableRefresh nextId="4"><queryTableFields count="3"><queryTableField id="1" name="Qty"/><queryTableField id="2" name="Price"/><queryTableField id="3" name="Amount"/></queryTableFields></queryTableRefresh></queryTable>`);zip.file('xl/connections.xml',`<connections xmlns="${NS}"><connection id="1" name="SourceSales" type="5" refreshedVersion="7" saveData="1"><dbPr connection="Provider=Example" command="SELECT Qty,Price FROM Sales"/></connection></connections>`);
  const rels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));rels.documentElement.appendChild(rels.importNode(xml(relationships(relationship('rIdConnection',R+'/connections','connections.xml'))).documentElement.firstChild,true));zip.file('xl/_rels/workbook.xml.rels',stringify(rels));const types=xml(await zip.file('[Content_Types].xml').async('string'));for(const [path,type] of [['xl/queryTables/queryTable1.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml'],['xl/connections.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml']]){const n=types.createElementNS(types.documentElement.namespaceURI,'Override');n.setAttribute('PartName','/'+path);n.setAttribute('ContentType',type);types.documentElement.appendChild(n);}zip.file('[Content_Types].xml',stringify(types));let bytes=await zip.generateAsync({type:'uint8array'}),files;
  for(let round=0;round<3;round++){
    const book=validateWorkbook(await readSpreadsheet(bytes,'query.xlsx'));bytes=(await writeSpreadsheet(book)).data;const out=await JSZip.loadAsync(bytes),queryRel=(await readRelationships(out,'xl/tables/table1.xml'))[0],query=resolvePart('xl/tables/table1.xml',queryRel.target);assert.equal(attr(xml(await out.file(query).async('string')).documentElement,'connectionId'),'1');const conn=(await readRelationships(out,'xl/workbook.xml')).find(r=>r.type.endsWith('/connections')),path=resolvePart('xl/workbook.xml',conn.target);assert.equal(attr(all(xml(await out.file(path).async('string')),'dbPr')[0],'command'),'SELECT Qty,Price FROM Sales');const paths=Object.keys(out.files).filter(p=>!out.files[p].dir);if(files===undefined)files=paths.length;else assert.equal(paths.length,files);
  }
});
test('overlapping tables and duplicate workbook names are rejected',async()=>{
  const book=await readSpreadsheet(await tableFixture(),'tables.xlsx'),s=book.sheets[0],copy=structuredClone(s.tables[0]);copy.id='other';copy.name='Other';s.tables.push(copy);assert.throws(()=>validateWorkbook(book),/겹칠/);s.tables.pop();book.sheets.push({...structuredClone(s),id:'copy',name:'Copy'});assert.throws(()=>validateWorkbook(book),/표 이름/);assert.deepEqual(tableBounds('A1:C4'),{r1:0,r2:3,c1:0,c2:2});
});
