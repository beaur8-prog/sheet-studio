// Recalculate the actual UI-created dashboard in a separate engine. No real
// user files or Microsoft Excel are involved. Inspect differences in the report.
import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import JSZip from 'jszip';
import {readSpreadsheet} from '../src/io.mjs';
import {createEvaluator,parseAddress} from '../src/engine.mjs';
import {worksheetParts} from '../src/xlsx-preservation.mjs';
import {xml,all,attr,child,stringify} from '../src/chart-xml.mjs';

const root=path.resolve(import.meta.dirname,'..'),file=process.argv.slice(2).find(arg=>!arg.startsWith('--'))||path.join(root,'test-results/dashboard-journey.xlsx'),plain=process.argv.includes('--plain-references'),cached=process.argv.includes('--cached');
// This optional diagnostic control changes only a private input copy. It is
// not an app export behavior and does not replace the user's original formulas.
const folder=path.join(root,'.verification/dashboard-interop'+(plain?'-plain':'')+(cached?'-cached':'')),output=path.join(folder,'converted'),profile=path.join(folder,'profile-'+Date.now());
await fs.mkdir(output,{recursive:true});
const data=await fs.readFile(file),b=await readSpreadsheet(data,'dashboard.xlsx'),z=await JSZip.loadAsync(data);
const checks={Source:{D2:440,D3:80,D4:150,D5:160,D6:120,D7:30},Sheet2:{A1:'Region',D1:'Amount',A2:'Seoul',D2:440,A3:'Jeju',D3:160,A4:'Seoul',D4:150,A5:'Busan',D5:120,F2:870},'피벗 1':{A2:'Busan',A3:'Jeju',A4:'Seoul',B5:870}};
let poisoned=0;
for(const part of await worksheetParts(z)){
 const sheet=b.sheets.find(s=>s.name===part.name),calc=createEvaluator(b,sheet.id),arrays=calc.spills();
 const doc=xml(await z.file(part.path).async('string'));
 for(const cell of all(doc,'c')){
  const ref=attr(cell,'r'),p=parseAddress(ref),derived=arrays.some(a=>p.r>=a.r1&&p.r<=a.r2&&p.c>=a.c1&&p.c<=a.c2);
  if(!child(cell,'f')&&!derived)continue;const value=child(cell,'v');if(value&&!cached){value.textContent='-999';cell.removeAttribute('t');poisoned++;}
  if(plain&&part.name==='Sheet2'){
   const control={A1:"'Source'!$A$1:$D$1",A2:"_xlfn._xlws.SORT(_xlfn._xlws.FILTER('Source'!$A$2:$D$7,'Source'!$D$2:$D$7>=100),4,-1)",F2:'SUM(INDEX(A2:D5,,4))'}[ref];
   if(control&&child(cell,'f'))child(cell,'f').textContent=control;
  }
 }
 z.file(part.path,stringify(doc));
}
if(!poisoned&&!cached)throw Error('Expected calculated cells in the UI-generated workbook.');
const input=path.join(folder,'dashboard.xlsx');await fs.writeFile(input,await z.generateAsync({type:'uint8array'}));
const chartCount=async zip=>(await Promise.all(Object.keys(zip.files).filter(name=>/\/charts\/[^/]+\.xml$/.test(name)).map(async name=>/<(?:\w+:)?chartSpace(?:\s|>)/.test(await zip.file(name).async('string'))))).filter(Boolean).length;
try{
 await fs.mkdir(path.join(profile,'user'),{recursive:true});await fs.writeFile(path.join(profile,'user/registrymodifications.xcu'),'<oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>');
 await fs.rm(path.join(output,'dashboard.xlsx'),{force:true});
 const exec=promisify(execFile),soffice=process.env.SHEET_STUDIO_SOFFICE||'soffice',version=(await exec(soffice,['--version'],{timeout:30000})).stdout.trim();
 const log=await exec(soffice,['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--nologo','--norestore','--convert-to','xlsx:Calc MS Excel 2007 XML','--outdir',output,input],{timeout:90000,maxBuffer:2000000});await fs.writeFile(path.join(folder,'conversion.log'),log.stdout+log.stderr);
 const bytes=await fs.readFile(path.join(output,'dashboard.xlsx')),out=await JSZip.loadAsync(bytes),strings=out.file('xl/sharedStrings.xml')?all(xml(await out.file('xl/sharedStrings.xml').async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent).join('')):[];
 const roundtrip=await readSpreadsheet(bytes,'dashboard.xlsx'),parts=await worksheetParts(out),differences=[];let compared=0;
 for(const [name,values] of Object.entries(checks)){
  const part=parts.find(p=>p.name===name),sheet=roundtrip.sheets.find(s=>s.name===name);if(!part||!sheet)throw Error('Missing output worksheet '+name);
  const doc=xml(await out.file(part.path).async('string')),cells=new Map(all(doc,'c').map(c=>[attr(c,'r'),c])),e=createEvaluator(roundtrip,sheet.id);
  for(const [ref,expected] of Object.entries(values)){
   const cell=cells.get(ref),type=attr(cell,'t'),text=child(cell,'v')?.textContent,actual=type==='s'?strings[Number(text)]:type==='str'||type==='e'?text:type==='b'?text==='1':text===undefined?null:Number(text),reimported=e(ref);compared++;
   if(actual===-999)throw Error('A poisoned output was not recalculated: '+name+'!'+ref);
   if(actual!==expected||reimported!==expected)differences.push({sheet:name,ref,expected,libreoffice:actual,sheetStudioReimport:reimported});
  }
 }
 const report={engine:version,at:new Date().toISOString(),input:'actual UI-created table, calculated column, cross-sheet SORT/FILTER dashboard, chart, pivot and page settings',plainReferenceControl:plain,originalCachesRetained:cached,poisonedCells:poisoned,compared,matched:compared-differences.length,inputCharts:await chartCount(z),outputCharts:await chartCount(out),differences,scope:'Independent forced LibreOffice recalculation and Sheet Studio reimport. LibreOffice may refresh/rearrange the native pivot against its own calculated source, so changed report cells are recorded as differences. No Microsoft Excel/VBA execution.'};
 await fs.writeFile(path.join(folder,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));
}finally{await fs.rm(profile,{recursive:true,force:true});}
