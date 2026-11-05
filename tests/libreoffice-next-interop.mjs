// Optional independent-engine check for the second Excel improvement block.
// Generate the SDK inputs first. Only synthetic files and an isolated LO profile
// are used; this check does not claim Microsoft Excel rendering or VBA coverage.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {pathToFileURL} from 'node:url';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {createEvaluator} from '../src/engine.mjs';
import {validationRules,validateCellUpdates} from '../src/cell-constraints.mjs';
import {chartData} from '../src/chart-model.mjs';
import {preservedPivots} from '../src/native-pivots.mjs';

const root=path.resolve(import.meta.dirname,'..'),output=path.join(root,'test-results','libreoffice-next'),source=path.join(root,'test-results','ooxml-sdk'),input=path.join(output,'input'),converted=path.join(output,'converted'),profile=path.join(output,'profile-'+Date.now());
const execute=promisify(execFile),soffice=process.env.SHEET_STUDIO_SOFFICE||'soffice',names=['authored-validation','visual-conditional','preserved-advanced-bar','table-linked-pivot','explicit-scatter-groups','explicit-combo-series'];
await fs.mkdir(input,{recursive:true});await fs.mkdir(converted,{recursive:true});
const before=new Map();for(const name of names){const bytes=await fs.readFile(path.join(source,name+'-0.xlsx'));before.set(name,await readSpreadsheet(bytes,name+'.xlsx'));await fs.writeFile(path.join(input,name+'.xlsx'),bytes);}
const result={engine:(await execute(soffice,['--version'],{timeout:30000})).stdout.trim(),at:new Date().toISOString(),scope:'Synthetic XLSX opened and re-saved by LibreOffice, then three Sheet Studio roundtrips. Not Microsoft Excel/VBA or a pixel comparison.',cases:[],differences:[]};
const values=book=>book.sheets.map(sheet=>{const calc=createEvaluator(book,sheet.id);return Object.fromEntries(Object.keys(sheet.cells).filter(ref=>sheet.cells[ref].raw!=='').sort().map(ref=>[ref,calc.result(ref).value]));});
const charts=book=>book.sheets.flatMap(sheet=>(sheet.charts||[]).map(chart=>({type:chart.type,series:chartData(chart,sheet,book).map(s=>({name:s.name,values:s.values,labels:s.labels,...(['scatter','bubble'].includes(chart.type)?{xs:s.xs}:{}),...(chart.type==='scatter'?{lines:s.lines}:{})}))})));
const appearance=book=>book.sheets.flatMap(sheet=>(sheet.charts||[]).map(chart=>chart.series.map(s=>({color:s.color?.toLowerCase(),type:s.type,secondary:!!s.secondary,lines:!!s.lines}))));
const validationBehavior=book=>{const sheet=book.sheets[0];return [['A1','예'],['A1','보류'],['B1','5'],['B1','99'],['B2','99'],['C1','6'],['C1','4'],['D1','46331'],['D1','45000'],['E1','0.5'],['E1','0.2'],['F1','가나다'],['F1','가'.repeat(11)],['G1','6'],['G1','4']].map(([key,raw])=>({key,raw,severity:validateCellUpdates(book,sheet,[{key,cell:{raw}}])?.severity||null}));};
try{
 const conversion=await execute(soffice,['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--nologo','--norestore','--convert-to','xlsx:Calc MS Excel 2007 XML','--outdir',converted,...names.map(name=>path.join(input,name+'.xlsx'))],{timeout:120000,maxBuffer:2*1024*1024});await fs.writeFile(path.join(output,'conversion.log'),conversion.stdout+conversion.stderr);
 for(const name of names){const original=before.get(name);let actual=await readSpreadsheet(await fs.readFile(path.join(converted,name+'.xlsx')),name+'.xlsx');
  if(name==='table-linked-pivot'){
   assert.deepEqual(values(actual)[0],values(original)[0],'pivot source values');
   const nonzero=sheet=>Object.values(sheet).filter(v=>typeof v==='number'&&v!==0).sort((a,b)=>a-b);assert.deepEqual(nonzero(values(actual)[1]),nonzero(values(original)[1]),'all nonzero pivot aggregates');
   assert.equal(actual.sheets[1].pivot,undefined,'external layout must not silently refresh with the old app spec');assert.equal(preservedPivots(actual,actual.sheets[1].id).length,1);
   result.differences.push({case:name,beforeReport:values(original)[1],afterReport:values(actual)[1],note:'LibreOffice adds a caption/header row, translates total labels and leaves empty intersections blank. Source and nonzero aggregates match. Sheet retains this external report/native pivot without restoring the old app layout.'});
  }else assert.deepEqual(values(actual),values(original),name+' cell values');
  if(name==='authored-validation'){assert.deepEqual(validationBehavior(actual),validationBehavior(original));const prompt=validationRules(actual.sheets[0]).find(rule=>rule.type==='list');assert.equal(prompt.promptTitle,'승인 여부');assert.equal(prompt.prompt,'목록에서 선택하세요.');assert.equal(prompt.severity,'warning');}
  if(name.startsWith('explicit-')){assert.deepEqual(charts(actual),charts(original),name+' chart types, names, points, labels and connection policy');if(JSON.stringify(appearance(actual))!==JSON.stringify(appearance(original)))result.differences.push({case:name,beforeAppearance:appearance(original),afterLibreOffice:appearance(actual),note:'External engine changes chart appearance. Source colors/axes are compared separately and preserved from this external result in subsequent Sheet saves; this is not an identical-rendering pass.'});}
  const expectedAfterLO={values:values(actual),charts:charts(actual),appearance:appearance(actual),...(name==='authored-validation'?{validation:validationBehavior(actual)}:{})};
  const summary=book=>book.sheets.map(s=>({name:s.name,tables:s.tables?.map(t=>({name:t.name,range:t.range}))||[],editableConditional:s.conditionalRules?.map(r=>r.type)||[],preservedConditional:s.unsupportedConditionalCount||0,pivot:!!s.pivot,nativePivots:preservedPivots(book,s.id).length}));
  if(JSON.stringify(summary(original))!==JSON.stringify(summary(actual)))result.differences.push({case:name,before:summary(original),afterLibreOffice:summary(actual),note:'External-engine metadata representation changed. This is recorded separately from values and functional checks; no claim of identical Excel editing or rendering.'});
  for(let round=0;round<3;round++){const bytes=(await writeSpreadsheet(actual)).data;await fs.writeFile(path.join(output,name+'-sheet-'+round+'.xlsx'),bytes);actual=await readSpreadsheet(bytes,name+'.xlsx');assert.deepEqual(values(actual),expectedAfterLO.values,name+' Sheet round '+round+' values');assert.deepEqual(charts(actual),expectedAfterLO.charts,name+' Sheet round '+round+' charts');assert.deepEqual(appearance(actual),expectedAfterLO.appearance,name+' Sheet round '+round+' appearance');if(name==='authored-validation')assert.deepEqual(validationBehavior(actual),expectedAfterLO.validation);}
  result.cases.push({case:name,status:result.differences.some(d=>d.case===name)?'passed_with_documented_representation_change':'passed',sheetRoundtrips:3});
 }
 await fs.writeFile(path.join(output,'report.json'),JSON.stringify(result,null,2)+'\n');console.log(JSON.stringify(result,null,2));
}finally{await fs.rm(profile,{recursive:true,force:true});}
