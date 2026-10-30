import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {createWorkbook,createEvaluator,evaluateFormula,formulaReference,absoluteFormula,validateWorkbook} from '../src/engine.mjs';
import {nameDefinition,nameError} from '../src/named-ranges.mjs';
import {refValues,chartData} from '../src/chart-model.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,all,attr} from '../src/chart-xml.mjs';

function fixture(){const book=createWorkbook(),s=book.sheets[0];s.name='Input';s.cells={A1:{raw:'2'},A2:{raw:'3'},A3:{raw:'4'},B1:{raw:'=SUM(Values)*(1+Rate)'},B2:{raw:'=ROWS(Values)'},B3:{raw:'=ROW(Self)'}};book.sheets.push({id:'other',name:'Other',cells:{A1:{raw:'=Rate'}},columnWidths:{},rowHeights:{},charts:[],merges:[]});book.chartNames=[{name:'Values',formula:"'Input'!$A$1:$A$3"},{name:'Rate',formula:'0.1'},{name:'Rate',formula:'0.2',sheetId:'other'},{name:'Self',formula:"'Input'!$B$3"},{name:'Alias',formula:'Values'}];return book;}
test('defined ranges, constants, aliases, dynamic names and local scope participate in calculation',()=>{
  const book=fixture(),s=book.sheets[0],calc=createEvaluator(book,s.id);assert.equal(calc('B1'),9.9);assert.equal(calc('B2'),3);assert.equal(calc('B3'),3);assert.equal(createEvaluator(book,'other')('A1'),.2);
  assert.ok(Math.abs(evaluateFormula(book,s.id,"='Other'!Rate+Rate")-.3)<1e-15);assert.equal(evaluateFormula(book,s.id,'=SUM(Alias)'),9);assert.equal(evaluateFormula(book,s.id,'=SUM(INDIRECT("Values"))'),9);assert.equal(evaluateFormula(book,s.id,'=IF(FALSE,Missing,Rate)'),.1);
  book.chartNames.push({name:'Dynamic',formula:"OFFSET('Input'!$A$1,0,0,COUNTA('Input'!$A$1:$A$3),1)"});assert.equal(evaluateFormula(book,s.id,'=SUM(Dynamic)'),9);
  assert.deepEqual(formulaReference(book,s.id,'Alias'),{sheetId:s.id,range:{r1:0,r2:2,c1:0,c2:0}});assert.equal(formulaReference(book,s.id,'Rate'),null);
  book.chartNames.push({name:'Loop',formula:'Loop2'},{name:'Loop2',formula:'Loop'});assert.equal(evaluateFormula(book,s.id,'=Loop'),'#CYCLE!');assert.equal(evaluateFormula(book,s.id,'=IFERROR(Loop,7)'),7);
  book.chartNames.push({name:'RelativeImported',formula:'Input!A1'});assert.equal(evaluateFormula(book,s.id,'=RelativeImported'),'#NAME?');assert.equal(book.chartNames.at(-1).formula,'Input!A1');
});
test('new names bind absolute cells on their owner sheet and reject duplicates/reserved addresses',()=>{
  const book=fixture(),s=book.sheets[0];const def=nameDefinition(book,s,{name:'Sales',formula:'=SUM(A1:A3)'});assert.equal(def.formula,"SUM('Input'!$A$1:$A$3)");
  assert.equal(absoluteFormula('=SUM(A1:B2)+\'Other\'!C3+D$4+N("A1")','Input'),"=SUM('Input'!$A$1:$B$2)+'Other'!$C$3+'Input'!$D$4+N(\"A1\")");
  for(const name of ['A1','XFD1048576','R1C1','two words','1start','R','C'])assert.ok(nameError(book,name));assert.ok(nameError(book,'rate'));assert.equal(nameError(book,'Rate',s.id),null);assert.equal(nameError(book,'매출_합계'),null);
});
test('chartless defined names retain source attributes and local scope through XLSX',async()=>{
  const base=await writeSpreadsheet(fixture()),zip=await JSZip.loadAsync(base.data);let text=await zip.file('xl/workbook.xml').async('string');text=text.replace('name="Values"','name="Values" comment="매출 범위" hidden="0"');zip.file('xl/workbook.xml',text);let bytes=await zip.generateAsync({type:'uint8array'});
  for(let round=0;round<3;round++){
    const book=validateWorkbook(await readSpreadsheet(bytes,'names.xlsx'));assert.equal(createEvaluator(book,book.sheets[0].id)('B1'),9.9);assert.equal(createEvaluator(book,book.sheets[1].id)('A1'),.2);bytes=(await writeSpreadsheet(book)).data;
    const out=await JSZip.loadAsync(bytes),doc=xml(await out.file('xl/workbook.xml').async('string')),names=all(doc,'definedName');assert.equal(attr(names.find(n=>attr(n,'name')==='Values'),'comment'),'매출 범위');assert.equal(attr(names.find(n=>attr(n,'name')==='Rate'&&attr(n,'localSheetId')==='1'),'localSheetId'),'1');
  }
});
test('changing a named range updates chart data and cache invalidation prevents deleted-name stale values',async()=>{
  const book=fixture(),s=book.sheets[0];s.charts=[{id:'chart',type:'line',range:'A1:B3',width:500,height:300,title:'Names',series:[{name:'Sales',valRef:{sheetId:s.id,range:'A1:A3',formula:'Values'}}]}];
  let imported=await readSpreadsheet((await writeSpreadsheet(book)).data,'chart.xlsx'),sheet=imported.sheets[0],named=imported.chartNames.find(n=>n.name==='Values');named.formula="'Input'!$A$2:$A$3";imported.referenceCachesValid=false;
  assert.deepEqual(chartData(sheet.charts[0],sheet,imported)[0].values,[3,4]);
  let output=await writeSpreadsheet(imported),outZip=await JSZip.loadAsync(output.data),part=Object.keys(outZip.files).find(p=>p.endsWith('.xml')&&p.includes('sheetStudioChart')&&!p.includes('Drawing'));
  const doc=xml(await outZip.file(part).async('string'));assert.ok(all(doc,'f').some(n=>n.textContent==='Values'));assert.ok(all(doc,'numCache').some(n=>all(n,'v').map(v=>Number(v.textContent)).join(',')==='3,4'));
  imported.chartNames=imported.chartNames.filter(n=>!['Values','Alias'].includes(n.name));sheet.cells.Z1={raw:'=DeletedName',cachedFormula:'=DeletedName',cachedValue:99};
  assert.deepEqual(refValues({sheetId:sheet.id,range:'Z1:Z1'},imported),['#NAME?']);assert.deepEqual(chartData(sheet.charts[0],sheet,imported),[]);
  const excel=new ExcelJS.Workbook();await excel.xlsx.load((await writeSpreadsheet(imported)).data);assert.deepEqual(excel.worksheets[0].getCell('Z1').result,{error:'#NAME?'});
});
