import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {readSpreadsheet} from '../src/io.mjs';
import {buildNativePlan} from '../src/native-plan.mjs';
import {chartCases} from '../tests/fixtures/charts.mjs';
const vba=createRequire(import.meta.url)('../electron/vba.cjs'),root=path.resolve(import.meta.dirname,'..');
const file=path.join(root,'.verification/native-source.xlsm'),original=await fs.readFile(file),book=await readSpreadsheet(original,'source.xlsm'),sheet=book.sheets[0];
book.nativeSource.originToken=await vba.rememberFile(file,original);
assert.ok(sheet.charts.length);sheet.charts[0].title='Native chart edit';sheet.charts[0].width=680;sheet.charts[0].height=410;sheet.charts[0].legend=false;
const rows=[['Category','Revenue','Cost'],['A','40','22'],['B','25','18'],['C','10','5']];
rows.forEach((row,r)=>row.forEach((v,c)=>sheet.cells[String.fromCharCode(65+c)+(r+5)]={raw:v}));
for(const [i,spec] of chartCases.entries())sheet.charts.push({id:'new-'+i,title:spec.type+' '+spec.grouping,range:'A5:C8',x:100,y:500+i*350,width:600,height:320,legend:true,labels:true,xTitle:['pie','doughnut'].includes(spec.type)?'':'Category',yTitle:['pie','doughnut'].includes(spec.type)?'':'Value',...spec});
async function save(model){const s=await vba.start({mode:'save',source:model.nativeSource,plan:buildNativePlan(model)}),deadline=Date.now()+60000;while(true){const state=await vba.status(s.id);if(state.state==='error')throw Error(state.message);if(state.state==='completed')break;if(Date.now()>deadline)throw Error('Excel chart save timeout');await new Promise(r=>setTimeout(r,200));}const result=await vba.result(s.id);return readSpreadsheet(new Uint8Array(result.data),result.name);}
try{
 const loaded=await save(book);assert.equal(loaded.sheets[0].charts.length,19);assert.equal(loaded.sheets[0].charts[0].title,'Native chart edit');assert.ok(Math.abs(loaded.sheets[0].charts[0].width-680)<1);assert.equal(loaded.sheets[0].charts[0].legend,false);
 loaded.sheets[0].charts.splice(1,1);const again=await save(loaded);assert.equal(again.sheets[0].charts.length,18);assert.deepEqual(await fs.readFile(file),original);
 console.log('XLSM charts passed: existing title/size/legend edits, 18 new chart variants, deletion, source unchanged.');
}finally{await vba.closeAll();}
