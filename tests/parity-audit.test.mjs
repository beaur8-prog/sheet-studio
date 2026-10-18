import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {xml as parseXml,all,attr} from '../src/chart-xml.mjs';
import {createWorkbook,evaluateFormula,evaluateCell,shiftFormula,transposeFormula,moveFormulaReferences,rewriteWorkbookSheetReferences,formulaForEditing} from '../src/engine.mjs';
import {formulaReferences} from '../src/formula-tools.mjs';
import {shiftWorksheetFeatures,createConditionalEvaluator,importConditional} from '../src/conditional.mjs';
import {pastePlan} from '../src/cell-paste.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {applyFilters,shiftAutoFilter} from '../src/filters.mjs';
import {chartSeries} from '../src/chart-model.mjs';
import {applyPivot} from '../src/pivot.mjs';
const book=(cells={})=>{const b=createWorkbook();b.sheets[0].cells=Object.fromEntries(Object.entries(cells).map(([k,raw])=>[k,{raw,style:{}}]));return b;};
const calc=(f,b=book())=>evaluateFormula(b,b.sheets[0].id,f);
test('audit: coercion, missing arguments, error aggregation and old dates',()=>{
 for(const [f,v] of [
 ['=""+1','#VALUE!'],['=IF("","T","F")','#VALUE!'],['=NOT("")','#VALUE!'],['=SUM("",5)','#VALUE!'],['=+"abc"','abc'],['=+TRUE',true],['=+""',''],
 ['=COUNTA(1,,2)',3],['=AVERAGE(10,,20)',10],['=MIN(10,,20)',0],['=IF(TRUE,,5)&"x"','0x'],['=COUNT(1/0,5)',1],['=COUNTA(1/0,5)',2],
 ['=TEXT(1,"dddd")','Sunday'],['=TEXT(59,"dddd")','Tuesday'],['=TEXT(60,"mmmm dd, yyyy")','February 29, 1900'],['=DATE(1900,2,30)',61],['=DATE(1900,3,0)',60],['=DATE(1900,1,60)',60],
 ['=SUM(IFERROR({1,2},1/0))',3],['=XLOOKUP(99,{1,2},{10,20},,0)','#N/A'],['=MINA({1,"abc"})',1]])assert.equal(calc(f),v,f);
});
test('audit: criteria and lookups distinguish physical blank, empty text and numeric text',()=>{
 const b=book({A1:'123',A2:"'123",A3:'abc',A4:'TRUE',A5:'=""'});
 for(const [criteria,count] of [['"="',1],['"<>"',5],['0',0],['FALSE',0],['">100"',1],['"<200"',1],['"<z"',2],['">a"',1]])assert.equal(calc(`=COUNTIF(A1:A6,${criteria})`,b),count,criteria);
 const l=book({A1:'1',A2:'3',B1:'10',B2:'30',B3:'99',A7:'=""',B6:'32',B7:'64'});
 for(const f of ['=MATCH(0,A6:A7,0)','=VLOOKUP(0,A6:B7,2,FALSE)','=XLOOKUP(0,A6:A7,B6:B7)'])assert.equal(calc(f,l),'#N/A',f);
 assert.equal(calc('=XLOOKUP("",A6:A7,B6:B7)',l),64);assert.equal(calc('=XMATCH("",A6:A7)',l),2);assert.equal(calc('=VLOOKUP(5,A1:B3,2,TRUE)',l),30);
 assert.equal(calc('=VLOOKUP("abc",{100,"num";"def","str"},2,TRUE)'),'#N/A');
 assert.equal(calc('=SUMIF(A1:A2,">0",B1)',l),40);
});
test('audit: reference operators, invalid endpoints, numeric sheet names and scanner casing',()=>{
 const b=book({A1:'1',A2:'2',A3:'3',B1:'10',B2:'20',B3:'30'});b.sheets[0].name='2026년';
 assert.equal(calc('=2026년!A1',b),1);assert.equal(formulaReferences('=2026년!A1')[0].sheet,'2026년');
 assert.equal(formulaReferences('=Sheet1!A1:sheet1!B2').length,1);
 assert.equal(formulaForEditing('_xlfn.XLOOKUP("_xlfn.XLOOKUP",{1},{2})'),'XLOOKUP("_xlfn.XLOOKUP",{1},{2})');
 assert.equal(calc('=SUM(INDEX(A1:B2,1,2):B3)',b),60);assert.equal(calc('=SUM(A1:INDEX(A1:B3,3,2))',b),66);
 assert.equal(calc('=SUM(A1:B2 A2:B3)',b),22);assert.equal(calc('=SUM((A1:A2,B1:B2))',b),33);
 for(const f of ['=SUM(#REF!:B5)',"=SUM('2026년'!#REF!:B5)",shiftFormula('=SUM(A1:B2)',-1,-1)])assert.equal(calc(f,b),'#REF!',f);
 b.chartNames=[{name:'test',formula:"'2026년'!$A$1:$B$3"}];rewriteWorkbookSheetReferences(b,'2026년','Renamed');assert.match(b.chartNames[0].formula,/Renamed/);
});
test('audit: conditional anchors survive deletion and disjoint sqref is rebased',()=>{
 const b=book({A2:'20',A3:'30'}),s=b.sheets[0];s.conditionalRules=[{id:'x',type:'expression',range:'B2:B4',formulae:['=A2>10'],style:{color:'#ff0000'}}];shiftWorksheetFeatures(s,'r',1,-1);assert.equal(s.conditionalRules[0].formulae[0],'=A2>10');assert.equal(s.conditionalRules[0].range,'B2:B3');
 const other={...structuredClone(s),id:'other',name:'Other'};other.conditionalRules[0].formulae=['=Sheet1!$A$5>100'];shiftWorksheetFeatures(other,'r',0,1,'Sheet1');assert.equal(other.conditionalRules[0].formulae[0],'=Sheet1!$A$6>100');
 importConditional({conditionalFormattings:[{ref:'A1:A5 C1:C5',rules:[{type:'expression',formulae:['A1=1'],style:{font:{bold:true}}}]}]},s);assert.equal(s.conditionalRules[1].formulae[0],'C1=1');
});
test('audit: paste skips hidden destinations, transposes relative and locked references, protects merges',()=>{
 const s=book().sheets[0],clip={r:0,c:0,data:[[{raw:'10'},{raw:'=A1'}]]};
 const p=pastePlan(s,clip,{r:0,c:0},{transpose:true});assert.equal(p.updates[1].cell.raw,'=A1');
 assert.equal(transposeFormula('=SUM(A1:B1)',{r:0,c:2},{r:2,c:4}),'=SUM(E1:E2)');
 assert.equal(transposeFormula('=$B$3',{r:1,c:1},{r:4,c:4}),'=$B$3');
 assert.equal(transposeFormula('=$A1+A$1',{r:1,c:1},{r:4,c:4}),'=D$1+$A4');
 const v=pastePlan(s,{r:0,c:0,data:[[{raw:'1'}],[{raw:'2'}]]},{r:0,c:0},{hiddenRow:r=>r===1});assert.deepEqual(v.updates.map(v=>v.key),['A1','A3']);
 s.merges=['A1:B1'];assert.throws(()=>pastePlan(s,clip,{r:0,c:1}),/병합/);
 const m=pastePlan(s,{...clip,merges:['A1:B1']},{r:2,c:0});assert.ok(m.merges.includes('A3:B3'));
 assert.equal(moveFormulaReferences('=SUM(A1:A3)',{sourceSheet:'Sheet1',targetSheet:'Sheet1',ownerSheet:'Sheet1',range:{r1:0,r2:0,c1:0,c2:0},dr:4,dc:4}),'=SUM((A2:A3,E5:E5))');
 assert.equal(moveFormulaReferences('=SUM(A1:A3)',{sourceSheet:'Sheet1',targetSheet:'Sheet1',ownerSheet:'Sheet1',range:{r1:0,r2:0,c1:0,c2:0},dr:1,dc:0}),'=SUM((A3:A3,A2:A2))');
 assert.equal(moveFormulaReferences('=A1+$A$1',{sourceSheet:'Sheet1',targetSheet:'Sheet1',ownerSheet:'Sheet1',range:{r1:0,r2:0,c1:0,c2:0},dr:4,dc:4}),'=E5+$E$5');
});
test('audit: empty charts and header-only filters have no phantom data',()=>{
 const b=book(),s=b.sheets[0];assert.deepEqual(chartSeries({range:'A1:B1'},s,b),[]);s.autoFilter={range:{r1:0,c1:0,r2:1,c2:1},columns:{},hiddenRows:[]};shiftAutoFilter(s,'r',1,-1);assert.equal(s.autoFilter,undefined);
});
test('audit: CSV escaping remains stable across repeated roundtrips',async()=>{
 let b=book({A1:"'=SUM(A1)",A2:"'+123",A3:"'-hello",A4:"'@user",A5:'-10'});
 for(let i=0;i<3;i++){b=await readSpreadsheet((await writeSpreadsheet(b,'csv')).data,'audit.csv');assert.equal(evaluateCell(b,b.sheets[0].id,'A1'),'=SUM(A1)');assert.equal(evaluateCell(b,b.sheets[0].id,'A2'),'+123');assert.equal(evaluateCell(b,b.sheets[0].id,'A5'),-10);}
});
test('audit: XLSX ISO dates, numeric old dates, future functions and manual hidden rows roundtrip',async()=>{
 const b=book({A1:'Date',A2:'1',A3:'60',A4:'61',B1:'=XLOOKUP(1,{1},{2})'}),s=b.sheets[0];for(const r of [2,3,4])s.cells['A'+r].style={format:'date'};
 s.rowHeights={3:0};s.autoFilter={range:{r1:0,c1:0,r2:3,c2:0},columns:{0:{type:'custom',kind:'number',join:'and',conditions:[{operator:'greaterThan',value:'1'}]}},hiddenRows:[]};applyFilters(b,s);
 const out=await writeSpreadsheet(b),z=await JSZip.loadAsync(out.data);let xml=await z.file('xl/worksheets/sheet1.xml').async('string');xml=xml.replace(/<c r="A2"[^>]*>.*?<\/c>/,'<c r="A2" t="d"><v>1900-01-01T12:00:00.000Z</v></c>');z.file('xl/worksheets/sheet1.xml',xml);
 const re=await readSpreadsheet(await z.generateAsync({type:'uint8array'}),'dates.xlsx');assert.equal(re.sheets[0].cells.A2.raw,'1.5');assert.equal(re.sheets[0].cells.A3.raw,'60');assert.equal(re.sheets[0].cells.A4.raw,'61');assert.equal(re.sheets[0].rowHeights[3],0);assert.ok(!re.sheets[0].autoFilter.hiddenRows.includes(3));assert.equal(re.sheets[0].cells.B1.raw,'=XLOOKUP(1,{1},{2})');
});
test('audit: broken pivot sources preserve cached sheet and allow workbook export',async()=>{
 const b=book({A1:'Item',B1:'Amount',A2:'a',B2:'2'});const s=structuredClone(createWorkbook().sheets[0]);s.id='pivot';s.name='Pivot';b.sheets.push(s);const spec={version:1,totals:true,sort:'source',sourceSheetId:b.sheets[0].id,sourceRange:'A1:B2',rows:['Item'],columns:[],values:[{field:'Amount',aggregate:'sum'}],filters:[],rowTotals:true,columnTotals:true};applyPivot(b,s,spec);const cached=structuredClone(s.cells);b.sheets.shift();const out=await writeSpreadsheet(b);assert.ok(out.data.length);assert.deepEqual(s.cells,cached);assert.ok(s.pivot);const re=await readSpreadsheet(out.data,'broken-pivot.xlsx');assert.deepEqual(Object.keys(re.sheets[0].cells),Object.keys(cached));
});

test('audit: pivot header-only source and page-filter row offsets export safely',async()=>{
 const b=book({A1:'Item',B1:'Amount',C1:'Region',A2:'a',B2:'2',C2:'East'});const s=structuredClone(createWorkbook().sheets[0]);s.id='pivot';s.name='Pivot';b.sheets.push(s);const spec={version:1,totals:true,sort:'source',sourceSheetId:b.sheets[0].id,sourceRange:'A1:C2',rows:['Item'],columns:[],values:[{field:'Amount',aggregate:'sum'}],filters:[{field:'Region',values:['East']}]};applyPivot(b,s,spec);
 s.merges=['A1:B1'];s.rowHeights[1]=0;
 const out=await writeSpreadsheet(b),z=await JSZip.loadAsync(out.data),xml=await z.file('xl/worksheets/sheet2.xml').async('string');
 const row=all(parseXml(xml),'row').find(n=>attr(n,'r')==='3');assert.equal(attr(row,'ht'),'33');
 assert.match(await z.file('xl/sharedStrings.xml').async('string'),/Region/);
 s.pivot.sourceRange='A1:C1';assert.ok((await writeSpreadsheet(b)).data.length);assert.ok(s.pivot);
});
