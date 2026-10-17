import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {FUNCTION_CATALOG} from '../src/function-catalog.mjs';
import {createWorkbook,evaluateFormula} from '../src/engine.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';
if(process.platform!=='win32')throw Error('This optional oracle check requires Windows desktop Excel. Portable fixture tests run with npm test on all platforms.');
const root=path.resolve(import.meta.dirname,'..'),dir=path.join(root,'.verification/functions-native');await fs.mkdir(dir,{recursive:true});
const cases=FUNCTION_CATALOG.filter(f=>!['RAND','RANDBETWEEN','NOW','TODAY'].includes(f.name)).map(f=>({name:f.name,formula:f.example}));
const wb=createWorkbook();wb.sheets[0].cells=Object.fromEntries([1,2,3,4,5].map((n,i)=>['A'+(i+1),{raw:String(n)}]));wb.sheets[0].cells.B1={raw:'=SUM(A1:A5)'};
cases.forEach((c,i)=>{wb.sheets[0].cells['D'+(i+1)]={raw:c.formula};});
const input=path.join(dir,'cases.json'),xlsx=path.join(dir,'functions.xlsx');await fs.writeFile(input,JSON.stringify(cases));
const bytes=(await writeSpreadsheet(wb)).data;await fs.writeFile(xlsx,bytes);
const exec=promisify(execFile),ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
async function oracle(output,readWorkbook){
  await exec(ps,['-NoProfile','-NonInteractive','-STA','-File',path.join(root,'scripts/excel-function-oracle.ps1'),'-InputPath',input,'-OutputPath',output,...(readWorkbook?['-ReadWorkbook',readWorkbook]:[])],{windowsHide:true,timeout:120000,maxBuffer:2*1024*1024});
  return JSON.parse((await fs.readFile(output,'utf8')).replace(/^\uFEFF/,''));
}
const direct=await oracle(path.join(dir,'direct.json')),exported=await oracle(path.join(dir,'exported.json'),xlsx);
const imported=await readSpreadsheet(bytes,'functions.xlsx'),unsupported=[];let compared=0;
const near=(a,b)=>typeof a==='number'&&Number.isFinite(a)&&Math.abs(a-b)<=1e-7*Math.max(1,Math.abs(b));
function complex(s){const n='[+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][+-]?\\d+)?';s=s.replace(/([+-])([ij])$/,'$11$2');const m=new RegExp(`^(${n})(${n})[ij]$`).exec(s);return m?[+m[1],+m[2]]:/[ij]$/.test(s)?[0,Number(s.slice(0,-1))]:[Number(s),0];}
function compare(actual,expected,name){
  if(typeof expected==='number')assert.ok(near(actual,expected),`${name}: ${actual} != ${expected}`);
  else if(name.startsWith('IM')&&typeof expected==='string'&&!expected.startsWith('#'))complex(actual).forEach((v,i)=>assert.ok(near(v,complex(expected)[i]),name));
  else assert.equal(actual,expected,name);
}
for(let i=0;i<cases.length;i++){
  const result=direct.cases[i];assert.ok(!result.oracleError,`${result.formula}: ${result.oracleError}`);
  assert.ok(!exported.cases[i].oracleError,JSON.stringify(exported.cases[i]));
  assert.deepEqual(exported.cases[i].expected,result.expected,`Excel reading exported ${result.name}`);
  if(result.expected==='#NAME?'){unsupported.push(result.name);continue;}
  compare(evaluateFormula(wb,'sheet-1',result.formula),result.expected,result.name);
  compare(evaluateFormula(imported,imported.sheets[0].id,result.formula),result.expected,result.name);compared++;
}
assert.deepEqual(await fs.readFile(xlsx),Buffer.from(bytes),'Read-only Excel validation must not modify the source workbook.');
console.log(JSON.stringify({ok:true,excelVersion:direct.excelVersion,excelBuild:direct.excelBuild,compared,exported:cases.length,notRecognizedByInstalledExcel:unsupported,sourceUnchanged:true}));
