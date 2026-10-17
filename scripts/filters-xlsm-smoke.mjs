import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import JSZip from 'jszip';
import {fixture,cases} from '../tests/fixtures/filters.mjs';
import {readSpreadsheet} from '../src/io.mjs';
import {applyFilters} from '../src/filters.mjs';
import {buildNativePlan} from '../src/native-plan.mjs';
const vba=createRequire(import.meta.url)('../electron/vba.cjs'),root=path.resolve(import.meta.dirname,'..');
const source=path.join(root,'.verification/native-source.xlsm'),original=await fs.readFile(source),book=await readSpreadsheet(original,'native-source.xlsm');
book.nativeSource.originToken=await vba.rememberFile(source,original);
const sheet=fixture().sheets[0];sheet.id='filter-test';sheet.name='FilterTest';book.sheets.push(sheet);
const dateSheet=structuredClone(sheet);dateSheet.id='filter-dates';dateSheet.name='FilterDates';for(let r=4;r<=12;r++)if(dateSheet.cells['C'+r]){dateSheet.cells['C'+r].style={format:'date'};dateSheet.cells['C'+r].raw=String(Number(dateSheet.cells['C'+r].raw)+45000)};dateSheet.autoFilter.columns={2:{type:'values',kind:'date',values:['45010','45030']}};book.sheets.push(dateSheet);applyFilters(book,dateSheet);
sheet.autoFilter.columns=structuredClone(cases.find(c=>c.name==='combined columns').columns);applyFilters(book,sheet);
async function save(b){
  const {id}=await vba.start({mode:'save',source:b.nativeSource,plan:buildNativePlan(b)}),end=Date.now()+45000;
  while(Date.now()<end){const s=await vba.status(id);if(s.state==='error')throw Error(s.message);if(s.state==='completed')return vba.result(id);await new Promise(r=>setTimeout(r,250));}throw Error('Excel filter save timed out');
}
try{
  const result=await save(book),bytes=Buffer.from(result.data),round=await readSpreadsheet(bytes,'filters.xlsm'),s=round.sheets.find(s=>s.name==='FilterTest');
  applyFilters(round,round.sheets.find(s=>s.name==='FilterDates'));
  assert.deepEqual(round.sheets.find(s=>s.name==='FilterDates').autoFilter.hiddenRows,dateSheet.autoFilter.hiddenRows);assert.deepEqual(s.autoFilter.hiddenRows,sheet.autoFilter.hiddenRows);assert.deepEqual(s.autoFilter.range,sheet.autoFilter.range);assert.equal(s.cells.D15.cachedValue,90);
  const oldZip=await JSZip.loadAsync(original),newZip=await JSZip.loadAsync(bytes);
  if(oldZip.file('xl/vbaProject.bin'))assert.ok(newZip.file('xl/vbaProject.bin'),'Excel retains the VBA project part');
  round.nativeSource.originToken=result.originToken;delete s.autoFilter;const cleared=await save(round),final=await readSpreadsheet(new Uint8Array(cleared.data),'cleared.xlsm');
  assert.equal(final.sheets.find(s=>s.name==='FilterTest').autoFilter,undefined);assert.deepEqual(await fs.readFile(source),original);
  await fs.writeFile(path.join(root,'.verification/filters-saved.xlsm'),bytes);console.log('XLSM filter save/clear passed; original file unchanged and VBA project part retained; no macros executed.');
}finally{await vba.closeAll();}
