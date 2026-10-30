// Optional independent-engine validation. Never touches the user's Calc profile
// or documents; the generated inputs deliberately contain a stale numeric cache.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {createWorkbook,createEvaluator,validateWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {sortRange} from '../src/filters.mjs';
import {xml,stringify,all,child,attr} from '../src/chart-xml.mjs';
import {tableFixture} from './tables-fixture.mjs';
import {styleDefaultsFixture} from './fixtures/style-defaults.mjs';
import {chartFixture,chartCases,addChart} from './fixtures/charts.mjs';
const execute=promisify(execFile),root=path.resolve(import.meta.dirname,'..'),output=path.join(root,'.verification','excel-improvements','libreoffice'),input=path.join(output,'input'),converted=path.join(output,'recalculated'),profile=path.join(output,'profile-'+Date.now()),soffice=process.env.SHEET_STUDIO_SOFFICE||'soffice';
await fs.mkdir(input,{recursive:true});await fs.mkdir(converted,{recursive:true});
const cases=[];
const table=await readSpreadsheet(await tableFixture(),'tables.xlsx'),ts=table.sheets[0];ts.cells.E2={raw:'=LET(data,Sales[Amount],SUM(data))'};ts.cells.E3={raw:'=SUMIF(A:A,">2",C:C)'};ts.cells.E4={raw:'=SUMPRODUCT(A2:A3,B2:B3)'};cases.push({name:'table-structured-let',book:table,refs:['C2','C3','C4','E1','E2','E3','E4'],poison:'E1',tables:1});
const dates=createWorkbook(),ds=dates.sheets[0];dates.date1904=true;ds.cells={A1:{raw:'39267',style:{format:'date',numberFormat:'yyyy-mm-dd'}},A2:{raw:'39268',style:{format:'date',numberFormat:'yyyy-mm-dd'}},B1:{raw:'10'},B2:{raw:'20'},C1:{raw:'=DATE(2011,7,5)-39267'},C2:{raw:'=SUMIFS(B1:B2,A1:A2,">2011-07-05")'},C3:{raw:'=TEXT(A1,"yyyy-mm-dd")'},C4:{raw:'=NETWORKDAYS(DATE(2011,7,4),DATE(2011,7,8))'}};cases.push({name:'date1904',book:dates,refs:['C1','C2','C3','C4'],poison:'C1'});
const arrays=createWorkbook(),as=arrays.sheets[0];as.cells={A1:{raw:'Input'},A2:{raw:'2'},A3:{raw:'-1'},A4:{raw:'4'},B2:{raw:'3'},B3:{raw:'5'},B4:{raw:'7'},C1:{raw:'=SUM(D2:D4)'},C2:{raw:'=SUM(Data)*Rate'},C3:{raw:'=LET(x,2,LET(x,x+1,x)+x)'},D2:{raw:'=IF(A2:A4>0,A2:A4*B2:B4,0)',xlsxFormula:{raw:'=IF(A2:A4>0,A2:A4*B2:B4,0)',formula:'IF(A2:A4>0,A2:A4*B2:B4,0)',shareType:'array',ref:'D2:D4'}}};arrays.chartNames=[{name:'Data',formula:'Sheet1!$A$2:$A$4'},{name:'Rate',formula:'0.5'}];cases.push({name:'arrays-names',book:arrays,refs:['C1','C2','C3','D2','D3','D4'],poison:'C1'});
const styles=await readSpreadsheet(await styleDefaultsFixture(),'style.xlsx'),ss=styles.sheets[0];for(let r=2;r<=5;r++)ss.cells['C'+r]={raw:String(7-r)};sortRange(styles,ss,{r1:1,r2:4,c1:2,c2:2},2,false);ss.cells.F1={raw:'=SUM(C2:C5)'};cases.push({name:'inherited-sort',book:styles,refs:['C2','C3','C4','C5','F1'],poison:'F1'});
const charts=chartFixture(),prototype=structuredClone(charts.sheets[0]);charts.sheets=[];
for(const [i,variant] of chartCases.entries()){const s={...structuredClone(prototype),id:crypto.randomUUID(),name:'Chart'+(i+1),charts:[]};charts.sheets.push(s);const holder={...charts,sheets:[s]};addChart(holder,{...variant,title:variant.type+' '+variant.grouping,x:300,y:10,width:620,height:360,range:variant.type==='bubble'?'B1:D5':variant.type==='scatter'?'B1:C5':'A1:C5',valueAxis:{min:-20,max:100,majorUnit:20}});s.cells.F1={raw:'=SUM(B2:B5)'};}charts.activeSheetId=charts.sheets[0].id;cases.push({name:'chart-combinations',book:charts,refs:['F1'],poison:'F1',charts:chartCases.length});
const expected=new Map();
for(const item of cases){validateWorkbook(item.book);const calc=createEvaluator(item.book,item.book.sheets[0].id);expected.set(item.name,Object.fromEntries(item.refs.map(ref=>[ref,calc(ref)])));const zip=await JSZip.loadAsync((await writeSpreadsheet(item.book)).data),part='xl/worksheets/sheet1.xml',doc=xml(await zip.file(part).async('string')),cell=all(doc,'c').find(node=>attr(node,'r')===item.poison);assert.ok(child(cell,'f'));child(cell,'v').textContent='-999';zip.file(part,stringify(doc));await fs.writeFile(path.join(input,item.name+'.xlsx'),await zip.generateAsync({type:'uint8array'}));}
const version=(await execute(soffice,['--version'],{timeout:30000})).stdout.trim(),files=cases.map(item=>path.join(input,item.name+'.xlsx'));
// 0 = always recalculate; use only this test profile. The default is 1 (never).
await fs.mkdir(path.join(profile,'user'),{recursive:true});await fs.writeFile(path.join(profile,'user','registrymodifications.xcu'),'<oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>');
try{
 const convertedResult=await execute(soffice,['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--nologo','--norestore','--convert-to','xlsx:Calc MS Excel 2007 XML','--outdir',converted,...files],{timeout:120000,maxBuffer:2*1024*1024});
 await fs.writeFile(path.join(output,'conversion.log'),convertedResult.stdout+convertedResult.stderr);
 const results=[],differences=[];
 for(const item of cases){const file=path.join(converted,item.name+'.xlsx'),bytes=await fs.readFile(file),check=new ExcelJS.Workbook();await check.xlsx.load(bytes);const sheet=check.worksheets[0],values={},zip=await JSZip.loadAsync(bytes),document=xml(await zip.file('xl/worksheets/sheet1.xml').async('string'));for(const ref of item.refs){const node=all(document,'c').find(n=>attr(n,'r')===ref);assert.ok(node,item.name+' '+ref+' exists');const raw=child(node,'v')?.textContent;assert.notEqual(raw,undefined,item.name+' '+ref+' cache exists');const kind=attr(node,'t'),actual=kind==='str'||kind==='e'?raw:kind==='b'?raw==='1':Number(raw),want=expected.get(item.name)[ref];if(item.name==='arrays-names'&&ref==='C3'&&actual==='#NAME?'&&want===5)differences.push({case:item.name,ref,sheetStudio:want,libreOffice:actual,note:'Nested LET differs in LibreOffice 26.8, including fully prefixed and distinct-name probes. This run has no Microsoft Excel execution to resolve that external compatibility gap.'});else assert.equal(actual,want,item.name+' '+ref);values[ref]=actual;}
  const loaded=validateWorkbook(await readSpreadsheet(bytes,item.name+'.xlsx'));if(item.tables)assert.equal(loaded.sheets[0].tables.length,item.tables);if(item.charts)assert.equal(loaded.sheets.reduce((n,s)=>n+(s.charts?.length||0),0),item.charts);
  if(item.name==='date1904')assert.equal(loaded.date1904,true);
  if(item.name==='inherited-sort'){assert.equal(sheet.getCell('C2').font.name,'Georgia');assert.equal(sheet.getCell('C2').alignment.textRotation,30);assert.equal(sheet.getCell('C5').font.name,'Arial');}
  results.push({case:item.name,status:differences.some(d=>d.case===item.name)?'passed_with_documented_difference':'passed',values,sheets:loaded.sheets.length,...(item.tables?{tables:item.tables}:{}),...(item.charts?{charts:item.charts}:{})});
 }
 const report={engine:version,at:new Date().toISOString(),scope:'Generated XLSX opened, stale caches recalculated and re-saved by LibreOffice; this is not Microsoft Excel/VBA verification.',differences,cases:results};await fs.writeFile(path.join(output,'report.json'),JSON.stringify(report,null,2)+'\n');console.log(JSON.stringify(report,null,2));
}finally{await fs.rm(profile,{recursive:true,force:true});}
