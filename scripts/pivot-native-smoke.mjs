import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {promisify} from 'node:util';
import {execFile} from 'node:child_process';
import {createRequire} from 'node:module';
import JSZip from 'jszip';
import {fixture,spec} from '../tests/fixtures/pivot.mjs';
import {applyPivot} from '../src/pivot.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {buildNativePlan} from '../src/native-plan.mjs';
const exec=promisify(execFile),root=path.resolve(import.meta.dirname,'..'),dir=path.join(root,'.verification');
const ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
const vba=createRequire(import.meta.url)('../electron/vba.cjs');
await fs.mkdir(dir,{recursive:true});
const cases=[
 ['sum',{},160],['average',{values:[{field:'매출',aggregate:'average'}]},40],
 ['multi',{rows:['지역','상품'],columns:[],values:[{field:'매출',aggregate:'average'},{field:'수량',aggregate:'countNumbers'}]},5],
 ['columns',{values:[{field:'매출',aggregate:'sum'},{field:'수량',aggregate:'max'}]},5],
 ['filter',{filters:[{field:'지역',values:['동부']}]},60],
 ['page',{rows:['상품'],columns:[],filters:[{field:'지역',values:['동부']}]},60],
 ['scalar',{rows:[],columns:[]},160],['no-totals',{totals:false},null],
 ['empty',{filters:[{field:'지역',values:[]}]},0]
];
const entries=[];
for(const [name,patch,total] of cases){const b=fixture(),s={...spec(),...patch},target={id:'pivot',name:'Summary',charts:[]};applyPivot(b,target,s);b.sheets.push(target);const out=await writeSpreadsheet(b);const file=path.join(dir,`pivot-check-${name}.xlsx`);await fs.writeFile(file,out.data);entries.push({name,path:file,total,empty:name==='empty'});}
await fs.writeFile(path.join(dir,'pivot-checks.json'),JSON.stringify(entries));
const script=path.join(root,'scripts/check-pivots.ps1');
const {stdout}=await exec(ps,['-NoProfile','-NonInteractive','-STA','-File',script,'-Manifest',path.join(dir,'pivot-checks.json')],{windowsHide:true,timeout:60000});
const report=JSON.parse(stdout.trim());assert.equal(report.length,cases.length);
for(const result of report){const reread=await readSpreadsheet(await fs.readFile(result.saved),'roundtrip.xlsx');assert.ok(reread.sheets[1].pivot,'Excel resave preserves app definitions');assert.equal(reread.sheets[1].pivotError,undefined);}
// Verify the VBA-preserving path retains app pivot settings in an XLSM package.
const source=path.join(dir,`pivot-macro-source-${Date.now()}.xlsm`);
await exec(ps,['-NoProfile','-NonInteractive','-STA','-File',path.join(root,'scripts/create-native-fixture.ps1'),'-Destination',source],{windowsHide:true,timeout:60000});
const original=await fs.readFile(source),book=await readSpreadsheet(original,'pivot-source.xlsm');book.nativeSource.originToken=await vba.rememberFile(source,original);
const sample=fixture();book.sheets.push(sample.sheets[0]);sample.sheets[0].name='PivotSource';const target={id:'pivot',name:'Summary',charts:[]};applyPivot(book,target,spec());book.sheets.push(target);
try{
 const session=await vba.start({mode:'save',source:book.nativeSource,plan:buildNativePlan(book)}),deadline=Date.now()+45000;
 while(true){const status=await vba.status(session.id);if(status.state==='error')throw Error(status.message);if(status.state==='completed')break;if(Date.now()>deadline)throw Error('XLSM pivot save timeout');await new Promise(r=>setTimeout(r,200));}
 const saved=await vba.result(session.id),bytes=new Uint8Array(saved.data),reopened=await readSpreadsheet(bytes,saved.name);
 const pivot=reopened.sheets.find(s=>s.name==='Summary');assert.deepEqual(pivot.pivot.values,spec().values);assert.equal(pivot.cells.D4.raw,'160');assert.equal(pivot.pivot.sourceSheetId,reopened.sheets.find(s=>s.name==='PivotSource').id);
 const oldZip=await JSZip.loadAsync(original),newZip=await JSZip.loadAsync(bytes);if(oldZip.file('xl/vbaProject.bin'))assert.ok(newZip.file('xl/vbaProject.bin'));
 assert.deepEqual(await fs.readFile(source),original);
 await fs.writeFile(path.join(dir,'pivot-native-report.json'),JSON.stringify({xlsx:report,xlsmSettings:true,sourceUnchanged:true},null,2));
 console.log('Native pivot verification passed: 9 XLSX cases, native refresh/totals, Excel resave and XLSM settings preservation.');
}finally{await vba.closeAll();}
