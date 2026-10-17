import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {createRequire} from 'node:module';
import {worksheetFixture} from '../tests/fixtures/worksheet.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {buildNativePlan} from '../src/native-plan.mjs';
import {createConditionalEvaluator,conditionalFormula} from '../src/conditional.mjs';
import {createEvaluator,parseAddress,evaluateFormula} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),dir=path.join(root,'.verification');await fs.mkdir(dir,{recursive:true});
const b=worksheetFixture(),file=path.join(dir,'worksheet-features.xlsx'),report=path.join(dir,'worksheet-excel-oracle.json');await fs.writeFile(file,(await writeSpreadsheet(b)).data);
await promisify(execFile)('powershell.exe',['-NoLogo','-NoProfile','-NonInteractive','-File',path.join(root,'scripts/excel-worksheet-oracle.ps1'),'-WorkbookPath',file,'-OutputPath',report],{windowsHide:true,timeout:60000});
const result=JSON.parse((await fs.readFile(report,'utf8')).replace(/^\uFEFF/,''));assert.deepEqual(result.view,{cols:1,rows:1,frozen:true,hiddenCol:true,hiddenRow:true});
const ole=hex=>parseInt(hex.slice(1,3),16)+256*parseInt(hex.slice(3,5),16)+65536*parseInt(hex.slice(5,7),16);
const compare=(expected,sheet)=>{const evaluate=createEvaluator(b,sheet.id),conditional=createConditionalEvaluator(b,sheet,evaluate);for(const [key,v] of Object.entries(expected)){const p=parseAddress(key),style=conditional(p.r,p.c,evaluate(key));assert.equal(ole(style.background||'#ffffff'),v.background,key+' background');assert.equal(ole(style.color||'#000000'),v.color,key+' color');}};
compare(result.cells,b.sheets[0]);b.sheets[0].cells.B3.raw='1';b.sheets[0].cells.B4.raw='5';compare(result.after,b.sheets[0]);assert.deepEqual(result.paste,[2,1,100,200]);assert.deepEqual(result.fill,['=B2*2','=C3*2']);
// Test the real native save bridge on a synthetic VBA-bearing source; never run macros.
const vba=createRequire(import.meta.url)('../electron/vba.cjs'),source=path.join(dir,'native-source.xlsm'),original=await fs.readFile(source),book=await readSpreadsheet(original,'native-source.xlsm');book.nativeSource.originToken=await vba.rememberFile(source,original);
const s=worksheetFixture().sheets[0];s.id='worksheet-native';s.name='WorksheetFeatures';book.sheets.push(s);
const save=async book=>{const {id}=await vba.start({mode:'save',source:book.nativeSource,plan:buildNativePlan(book)}),end=Date.now()+60000;while(Date.now()<end){const status=await vba.status(id);if(status.state==='error')throw Error(status.message);if(status.state==='completed')return vba.result(id);await new Promise(r=>setTimeout(r,250));}throw Error('Native worksheet save timed out');};
try{
  const out=await save(book),round=await readSpreadsheet(new Uint8Array(out.data),'worksheet.xlsm'),r=round.sheets.find(s=>s.name==='WorksheetFeatures');await fs.writeFile(path.join(dir,'worksheet-saved.xlsm'),Buffer.from(out.data));assert.deepEqual(r.freeze,{rows:1,cols:1});assert.equal(r.columnWidths[5],0);assert.equal(r.rowHeights[8],0);assert.equal(r.conditionalRules.length,4);assert.equal(r.conditionalRules[0].stopIfTrue,true);const c=createConditionalEvaluator(round,r),e=createEvaluator(round,r.id);assert.equal(c(3,1,e('B4')).background,'#ffc7ce');assert.equal(c(1,0,e('A2')).background,'#ffeb9c');assert.equal(c(3,2,e('C4')).color,'#9c0006');
  round.nativeSource.originToken=out.originToken;r.freeze={rows:0,cols:0};r.conditionalRules=[];delete r.columnWidths[5];delete r.rowHeights[8];const clear=await save(round),final=await readSpreadsheet(new Uint8Array(clear.data),'cleared.xlsm'),f=final.sheets.find(s=>s.name==='WorksheetFeatures');assert.ok(!f.freeze?.rows&&!f.freeze?.cols);assert.equal(f.conditionalRules?.length||0,0);assert.notEqual(f.columnWidths[5],0);assert.notEqual(f.rowHeights[8],0);assert.deepEqual(await fs.readFile(source),original);
  await fs.writeFile(path.join(dir,'worksheet-saved.xlsm'),Buffer.from(out.data));console.log('Native Excel passed: freeze panes, hidden rows/columns, live conditional colors, paste transpose, relative fill and XLSM apply/clear. Original file unchanged; no macros executed.');
}finally{await vba.closeAll();}
