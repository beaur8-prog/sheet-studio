// Independent-engine diagnostic, not a Microsoft Excel execution oracle.
// Deliberately poison every output cache so a match proves recalculation.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {pathToFileURL} from 'node:url';
import JSZip from 'jszip';
import {createWorkbook,createEvaluator,address} from '../src/engine.mjs';
import {writeSpreadsheet} from '../src/io.mjs';
import {xml,all,child,attr} from '../src/chart-xml.mjs';
const root=path.resolve(import.meta.dirname,'..'),out=path.join(root,'.verification/dynamic-arrays/libreoffice'),input=path.join(out,'input'),output=path.join(out,'output'),profile=path.join(out,'profile-'+Date.now()),exec=promisify(execFile),soffice=process.env.SHEET_STUDIO_SOFFICE||'soffice';
await fs.mkdir(input,{recursive:true});await fs.mkdir(output,{recursive:true});await fs.mkdir(path.join(profile,'user'),{recursive:true});await fs.writeFile(path.join(profile,'user/registrymodifications.xcu'),'<oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>');
const formulas=['=SEQUENCE(3,2,2,3)','=FILTER({1,2;3,4;5,6},{1;0;1})','=FILTER({1,2,3;4,5,6},{1,0,1})','=SORT({3,5;1,9;2,7})','=SORT({3,1,2;30,10,20},1,1,TRUE)','=UNIQUE({1,2;1,3;1,2})','=UNIQUE({1,1,2;2,2,3},TRUE)','=UNIQUE({1;2;1;3},FALSE,TRUE)','=SORT({"b";"A";"a"})','=UNIQUE({"a";"A";"b"})','=FILTER({1;2},{0;0},"empty")','=XLOOKUP(2,{1;2;3},{10,11;20,21;30,31})','=XLOOKUP(2,{1,2,3},{10,20,30;11,21,31})'];
const files=[],expectations=[];
for(const [i,formula] of formulas.entries()){
 const b=createWorkbook();b.sheets[0].cells.A1={raw:formula};const calc=createEvaluator(b,b.activeSheetId),a=calc.spills()[0],rows=a?a.r2+1:1,cols=a?a.c2+1:1;
 expectations.push({formula,rows,cols,values:Array.from({length:rows*cols},(_,n)=>calc(address(Math.floor(n/cols),n%cols)))});
 const zip=await JSZip.loadAsync((await writeSpreadsheet(b)).data),part='xl/worksheets/sheet1.xml';zip.file(part,(await zip.file(part).async('string')).replace(/<v>[^<]*<\/v>/g,'<v>-999</v>'));const file=path.join(input,'array-'+i+'.xlsx');await fs.writeFile(file,await zip.generateAsync({type:'uint8array'}));files.push(file);
}
try{
 const version=(await exec(soffice,['--version'],{timeout:30000})).stdout.trim();
 const log=await exec(soffice,['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--nologo','--norestore','--convert-to','xlsx:Calc MS Excel 2007 XML','--outdir',output,...files],{timeout:120000,maxBuffer:2*1024*1024});await fs.writeFile(path.join(out,'conversion.log'),log.stdout+log.stderr);
 const results=[];
 for(const [i,item] of expectations.entries()){
  const zip=await JSZip.loadAsync(await fs.readFile(path.join(output,'array-'+i+'.xlsx'))),doc=xml(await zip.file('xl/worksheets/sheet1.xml').async('string')),cells=new Map(all(doc,'c').map(c=>[attr(c,'r'),c])),values=[];
  for(let n=0;n<item.values.length;n++){const ref=address(Math.floor(n/item.cols),n%item.cols),cell=cells.get(ref);assert.ok(cell,ref+' cache exists');const text=child(cell,'v')?.textContent;assert.notEqual(text,undefined);const type=attr(cell,'t'),value=type==='str'||type==='e'?text:type==='b'?text==='1':Number(text);assert.equal(value,item.values[n],item.formula+' '+ref);values.push(value);}
  results.push({...item,actual:values,status:'passed'});
 }
 const report={engine:version,at:new Date().toISOString(),scope:'Generated XLSX stale caches recalculated by LibreOffice; no Microsoft Excel/VBA execution.',cases:results};await fs.writeFile(path.join(out,'report.json'),JSON.stringify(report,null,2)+'\n');console.log(JSON.stringify(report,null,2));
}finally{await fs.rm(profile,{recursive:true,force:true});}
