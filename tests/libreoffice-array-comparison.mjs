// Optional independent-engine diagnostic. Success means the comparison ran;
// inspect `differences`, since LibreOffice is not a Microsoft Excel oracle.
// Every generated output cache is poisoned before an isolated recalculation.
import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import JSZip from 'jszip';
import {createWorkbook,createEvaluator,address,parseAddress} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,all,child,attr,stringify} from '../src/chart-xml.mjs';

const root=path.resolve(import.meta.dirname,'..'),output=path.join(root,'.verification/array-comparison');
const converted=path.join(output,'converted'),profile=path.join(output,'profile-'+Date.now());
const soffice=process.env.SHEET_STUDIO_SOFFICE||'soffice',exec=promisify(execFile);
await fs.mkdir(converted,{recursive:true});
let seed=20261031;const random=n=>{seed=(Math.imul(seed,1664525)+1013904223)>>>0;return Math.floor(seed/4294967296*n);};
const formulas=[
 '=SORT(A1:B18,1)', '=SORT(A1:B18,1,-1)', '=SORT(A1:B18,2)', '=SORT(A1:B18,2,-1)',
 '=UNIQUE(A1:B18)', '=UNIQUE(A1:A18)', '=UNIQUE(A1:A18,,TRUE)',
 '=FILTER(A1:B18,MOD(C1:C18,2)=0)', '=FILTER(A1:B18,MOD(C1:C18,3)=0)',
 '=FILTER(A1:B18,A1:A18=2,"none")', '=FILTER(A1:B18,A1:A18=99,"none")',
 '=SORT(UNIQUE(A1:A18))', '=INDEX(A1:B18,2)', '=INDEX(A1:B18,,2)',
 '=XLOOKUP(2,A1:A18,A1:B18,"none")', '=XLOOKUP(99,A1:A18,A1:B18,SEQUENCE(2,2))',
 '=IF(MOD(C1:C18,2)=0,A1:A18,B1:B18)', '=IFERROR(A1:B18,0)',
 '=FILTER(A1:B18,{TRUE,FALSE})', '=UNIQUE(A1:B18,TRUE)'
];
const b=createWorkbook(),template=structuredClone(b.sheets[0]);b.sheets=[];
for(let i=0;i<12;i++){
 const sheet={...structuredClone(template),id:'array-'+i,name:(i<6?'Numeric':'Mixed')+(i%6+1),cells:{}};b.sheets.push(sheet);
 const pool=i<6?[undefined,'0','1','2','2','-3','7']:[undefined,'0','1','2',"'2","'a","'A","'",'TRUE','FALSE','=IF(FALSE,0,"")','=NA()'];
 for(let r=1;r<=18;r++){for(const c of ['A','B']){const raw=pool[random(pool.length)];if(raw!==undefined)sheet.cells[c+r]={raw};}sheet.cells['C'+r]={raw:String(r)};}
 for(let f=0;f<formulas.length;f++)sheet.cells[address(29+Math.floor(f/10)*30,(f%10)*4)]={raw:formulas[f]};
 sheet.cells.ZZ1={raw:'=SUM(1,2)'};
}
b.activeSheetId=b.sheets[0].id;
const expected=b.sheets.map(sheet=>{const e=createEvaluator(b,sheet.id);return formulas.map((formula,f)=>{
 const row=29+Math.floor(f/10)*30,col=f%10*4,ref=address(row,col),region=e.spillAt(ref),rows=region?region.r2-row+1:1,cols=region?region.c2-col+1:1;
 return {ref,formula,rows,cols,values:Array.from({length:rows*cols},(_,n)=>e(address(row+Math.floor(n/cols),col+n%cols)))};
});});
const exported=(await writeSpreadsheet(b)).data,roundtrip=await readSpreadsheet(exported,'array-comparison.xlsx');
await fs.writeFile(path.join(output,'own-roundtrip.xlsx'),exported);
await fs.writeFile(path.join(output,'source.sheet.json'),JSON.stringify(b));
let ownRoundtrip=0;
for(let i=0;i<b.sheets.length;i++){const e=createEvaluator(roundtrip,roundtrip.sheets[i].id);for(const item of expected[i]){const p=parseAddress(item.ref);for(let n=0;n<item.values.length;n++){const ref=address(p.r+Math.floor(n/item.cols),p.c+n%item.cols),actual=e(ref);if(!Object.is(actual,item.values[n]))throw Error('Own XLSX result changed: '+JSON.stringify({sheet:b.sheets[i].name,anchor:item.ref,formula:item.formula,ref,expected:item.values[n],actual}));ownRoundtrip++;}}}
const zip=await JSZip.loadAsync(exported);
for(let i=0;i<b.sheets.length;i++){
 const part='xl/worksheets/sheet'+(i+1)+'.xml',doc=xml(await zip.file(part).async('string'));
 for(const cell of all(doc,'c')){const ref=attr(cell,'r');if(parseAddress(ref).r>=29||ref==='ZZ1'){const value=child(cell,'v');if(value){cell.removeAttribute('t');value.textContent='-999';}}}
 zip.file(part,stringify(doc));
}
const file=path.join(output,'array-comparison.xlsx');await fs.writeFile(file,await zip.generateAsync({type:'uint8array'}));
try{
 await fs.mkdir(path.join(profile,'user'),{recursive:true});
 await fs.writeFile(path.join(profile,'user/registrymodifications.xcu'),'<oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>');
 await fs.rm(path.join(converted,'array-comparison.xlsx'),{force:true});
 const version=(await exec(soffice,['--version'],{timeout:30000})).stdout.trim();
 const log=await exec(soffice,['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--nologo','--norestore','--convert-to','xlsx:Calc MS Excel 2007 XML','--outdir',converted,file],{timeout:120000,maxBuffer:2000000});await fs.writeFile(path.join(output,'conversion.log'),log.stdout+log.stderr);
 const out=await JSZip.loadAsync(await fs.readFile(path.join(converted,'array-comparison.xlsx'))),strings=out.file('xl/sharedStrings.xml')?all(xml(await out.file('xl/sharedStrings.xml').async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent).join('')):[];
 const differences=[],groups={Numeric:{cases:0,matched:0},Mixed:{cases:0,matched:0}};let cases=0,matched=0,comparedValues=0;
 for(let i=0;i<b.sheets.length;i++){
  const doc=xml(await out.file('xl/worksheets/sheet'+(i+1)+'.xml').async('string')),cells=new Map(all(doc,'c').map(c=>[attr(c,'r'),c]));
  const read=ref=>{const cell=cells.get(ref);if(!cell)return null;const text=child(cell,'v')?.textContent,type=attr(cell,'t');return type==='s'?strings[Number(text)]:type==='str'||type==='e'?text:type==='inlineStr'?all(cell,'t').map(t=>t.textContent).join(''):type==='b'?text==='1':text===undefined?null:Number(text);};
  if(read('ZZ1')!==3)throw Error('Stale control was not recalculated: '+b.sheets[i].name);
  for(const item of expected[i]){
   cases++;const group=groups[i<6?'Numeric':'Mixed'];group.cases++;const p=parseAddress(item.ref),array=attr(child(cells.get(item.ref),'f'),'ref'),end=array?parseAddress(array.split(':').at(-1)):p,rows=end.r-p.r+1,cols=end.c-p.c+1;
   const actual=Array.from({length:rows*cols},(_,n)=>read(address(p.r+Math.floor(n/cols),p.c+n%cols)));
   comparedValues+=Math.max(actual.length,item.values.length);
   const equal=rows===item.rows&&cols===item.cols&&actual.every((value,n)=>typeof value==='number'&&typeof item.values[n]==='number'?Math.abs(value-item.values[n])<=1e-12*Math.max(1,Math.abs(item.values[n])):Object.is(value,item.values[n]));
   if(equal){matched++;group.matched++;}else differences.push({sheet:b.sheets[i].name,...item,actualRows:rows,actualCols:cols,actual});
  }
 }
 const report={engine:version,at:new Date().toISOString(),seed:20261031,cases,matched,groups,comparedValues,ownRoundtripValues:ownRoundtrip,staleControlsRecalculated:b.sheets.length,differences,scope:'Numeric/blank and mixed-type source ranges, optional arguments, whole row/column return, composition and per-element errors; independently forced LibreOffice recalculation. No Microsoft Excel/VBA execution.'};
 await fs.writeFile(path.join(output,'report.json'),JSON.stringify(report,null,2)+'\n');console.log(JSON.stringify({engine:version,cases,matched,groups,ownRoundtripValues:ownRoundtrip,differences:differences.length,report:path.join(output,'report.json')},null,2));
}finally{await fs.rm(profile,{recursive:true,force:true});}
