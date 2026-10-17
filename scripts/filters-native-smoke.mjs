import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {fixture,cases} from '../tests/fixtures/filters.mjs';
import {applyFilters} from '../src/filters.mjs';
import {writeSpreadsheet} from '../src/io.mjs';
const exec=promisify(execFile),root=path.resolve(import.meta.dirname,'..'),dir=path.join(root,'.verification');
if(process.platform!=='win32')throw Error('Installed Windows Excel is required for this optional comparison.');
await fs.mkdir(dir,{recursive:true});
const oracle=JSON.parse((await fs.readFile(path.join(root,'tests/fixtures/excel-filters.json'),'utf8')).replace(/^\uFEFF/,''));
const input=path.join(dir,'filter-open.json'),output=path.join(dir,'filter-open-result.json');
await fs.writeFile(input,JSON.stringify({cases:[{name:'open'},{name:'reapply',reapply:true}]}));
for(const [i,c] of cases.entries()){
  const book=fixture(),s=book.sheets[0];s.autoFilter.columns=structuredClone(c.columns);applyFilters(book,s);
  const out=await writeSpreadsheet(book),file=path.join(dir,`filter-export-${i}.xlsx`);await fs.writeFile(file,out.data);
  await exec('powershell.exe',['-NoProfile','-NonInteractive','-File',path.join(root,'scripts/excel-filter-oracle.ps1'),'-InputPath',input,'-OutputPath',output,'-ReadWorkbook',file],{windowsHide:true,timeout:45000});
  const actual=JSON.parse((await fs.readFile(output,'utf8')).replace(/^\uFEFF/,'')),expected=oracle.cases.find(v=>v.name===c.name);
  for(const v of actual.cases){assert.deepEqual(v.visible,expected.visible,c.name+' '+v.name);assert.equal(v.subtotal,expected.subtotal,c.name+' '+v.name);assert.equal(v.arrows,true);}
  console.log(`Excel opened and reapplied ${i+1}/${cases.length}: ${c.name}`);
}

for(const [name,rule,expected] of [['values',{type:'values',values:['10','30']},[3,5]],['custom',{type:'custom',kind:'date',join:'and',conditions:[{operator:'greaterThanOrEqual',value:'1900-01-20'},{operator:'lessThan',value:'1900-03-01'}]},[4,5,6,7,8]]]){
const b=fixture(),s=b.sheets[0];for(let r=4;r<=12;r++)if(s.cells['C'+r])s.cells['C'+r].style={format:'date'};s.autoFilter.columns={2:rule};applyFilters(b,s);await fs.writeFile('.verification/filter-date.xlsx',(await writeSpreadsheet(b)).data);
await exec('powershell.exe',['-NoProfile','-File','scripts/excel-filter-oracle.ps1','-InputPath','.verification/filter-open.json','-OutputPath','.verification/filter-date-result.json','-ReadWorkbook','.verification/filter-date.xlsx'],{windowsHide:true});
const result=JSON.parse((await fs.readFile('.verification/filter-date-result.json','utf8')).replace(/^\uFEFF/,''));for(const v of result.cases)assert.deepEqual(v.visible,expected,name+' '+v.name);console.log('Date '+name+' Excel open/reapply passed');}
