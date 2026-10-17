import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import JSZip from 'jszip';
import {createEvaluator,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {applyFilters,inferFilterRange,filterOptions,shiftAutoFilter,sortFilterRange} from '../src/filters.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {fixture,cases} from './fixtures/filters.mjs';

test('Excel AutoFilter oracle: criteria, blanks, wildcards, top/bottom, average, AND columns and SUBTOTAL',()=>{
  const oracle=JSON.parse(fs.readFileSync(new URL('./fixtures/excel-filters.json',import.meta.url),'utf8').replace(/^\uFEFF/,''));
  for(const c of cases){const book=fixture(),s=book.sheets[0];s.autoFilter.columns=c.columns;applyFilters(book,s);
    const expected=oracle.cases.find(v=>v.name===c.name),visible=Array.from({length:9},(_,i)=>i+3).filter(r=>!s.autoFilter.hiddenRows.includes(r));
    assert.deepEqual(visible,expected.visible,c.name);assert.equal(createEvaluator(book,s.id)('D15'),expected.subtotal,c.name+' subtotal');
  }
});
test('filter range respects offset headers, separators, selection and empty selection errors',()=>{
  const s=fixture().sheets[0];assert.deepEqual(inferFilterRange(s,{r1:4,r2:4,c1:2,c2:2}),s.autoFilter.range);
  assert.deepEqual(inferFilterRange(s,{r1:2,r2:2,c1:1,c2:3}),s.autoFilter.range);
  assert.deepEqual(inferFilterRange(s,{r1:2,r2:5,c1:1,c2:2}),{r1:2,r2:5,c1:1,c2:2});
  assert.throws(()=>inferFilterRange(s,{r1:30,r2:30,c1:20,c2:20}),/데이터/);
});
test('value list ignores its own criterion and narrows against other columns; reapply is explicit',()=>{
  const b=fixture(),s=b.sheets[0];s.autoFilter.columns=structuredClone(cases.find(c=>c.name==='combined columns').columns);applyFilters(b,s);
  assert.deepEqual(filterOptions(b,s,1).values.map(v=>v.value),['대전','부산','서울','']);
  const before=[...s.autoFilter.hiddenRows];s.cells.C4.raw='100';assert.deepEqual(s.autoFilter.hiddenRows,before);applyFilters(b,s);assert.ok(!s.autoFilter.hiddenRows.includes(3));
});
test('sort preserves header, whole records, stable ties and blanks last in both directions',()=>{
  const b=fixture(),s=b.sheets[0];s.cells.D4={raw:'=C4*2',style:{bold:true}};sortFilterRange(b,s,2,true);
  assert.equal(s.cells.B3.raw,'지역');assert.equal(s.cells.C4.raw,'90');assert.equal(s.cells.C12,undefined);
  assert.equal(s.cells.D11.raw,'=C11*2');assert.equal(s.cells.D11.style.bold,true);assert.equal(s.cells.A1.raw,'범위 밖 제목');
  sortFilterRange(b,s,2,false);assert.equal(s.cells.C4.raw,'10');assert.equal(s.cells.C12,undefined);
});
test('filters survive JSON, undo snapshots and structural shifts; deleted header removes arrows',()=>{
  const b=fixture(),s=b.sheets[0];s.autoFilter.columns=structuredClone(cases[1].columns);applyFilters(b,s);
  assert.deepEqual(parseWorkbook(serializeWorkbook(b)).sheets[0].autoFilter,s.autoFilter);
  shiftAutoFilter(s,'c',1,1);assert.equal(s.autoFilter.range.c1,2);assert.ok(s.autoFilter.columns[2]);
  shiftAutoFilter(s,'r',1,1);assert.equal(s.autoFilter.range.r1,3);assert.equal(s.autoFilter.hiddenRows[0],5);
  shiftAutoFilter(s,'r',3,-1);assert.equal(s.autoFilter,undefined);
});
test('SUBTOTAL excludes filtered rows for 9 and 109, and manually hidden rows only for 109',()=>{
  const b=fixture(),s=b.sheets[0];s.autoFilter.columns=structuredClone(cases[1].columns);applyFilters(b,s);s.rowHeights[3]=0;
  const e=createEvaluator(b,s.id);assert.equal(e('D15'),100);assert.equal(e('D16'),90);
});
test('malformed filter metadata cannot hide outside range or create unbounded conditions',()=>{
  for(const patch of [{hiddenRows:[0]},{columns:{100:{type:'values',values:['a']}}},{columns:{1:{type:'custom',kind:'text',join:'and',conditions:[{operator:'eval',value:'x'}]}}}]){
    const b=fixture();Object.assign(b.sheets[0].autoFilter,patch);assert.throws(()=>parseWorkbook(JSON.stringify(b)),/필터/);
  }
});
test('XLSX roundtrip preserves native filter criteria, hidden rows, arrows and cached SUBTOTAL',async()=>{
  for(const c of cases){const b=fixture(),s=b.sheets[0];s.autoFilter.columns=structuredClone(c.columns);applyFilters(b,s);
    const out=await writeSpreadsheet(b),zip=await JSZip.loadAsync(out.data),xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
    assert.match(xml,/<autoFilter ref="B3:D12"/);assert.equal(xml.includes('filterColumn'),!!Object.keys(c.columns).length);
    const re=await readSpreadsheet(out.data,'filters.xlsx'),restored=re.sheets[0];assert.deepEqual(restored.autoFilter.hiddenRows,s.autoFilter.hiddenRows,c.name);
    applyFilters(re,restored);assert.deepEqual(restored.autoFilter.hiddenRows,s.autoFilter.hiddenRows,c.name+' reapply');
    assert.equal(createEvaluator(re,restored.id)('D15'),createEvaluator(b,s.id)('D15'));
  }
});

test('date conditions use Excel serials and numeric/date/text menus follow column values',async()=>{
  const b=fixture(),s=b.sheets[0];for(let r=4;r<=12;r++)if(s.cells['C'+r])s.cells['C'+r].style={format:'date'};
  assert.equal(filterOptions(b,s,2).kind,'date');
  s.autoFilter.columns={2:{type:'custom',kind:'date',join:'and',conditions:[{operator:'greaterThanOrEqual',value:'1900-01-20'},{operator:'lessThan',value:'1900-03-01'}]}};
  applyFilters(b,s);assert.deepEqual(s.autoFilter.hiddenRows,[3,9,10,11]);
  const out=await writeSpreadsheet(b),re=await readSpreadsheet(out.data,'dates.xlsx');applyFilters(re,re.sheets[0]);assert.deepEqual(re.sheets[0].autoFilter.hiddenRows,s.autoFilter.hiddenRows);
});

test('native Excel filtered sort moves only visible records and keeps excluded rows fixed',()=>{
  const b=fixture(),s=b.sheets[0];s.autoFilter.columns=structuredClone(cases[1].columns);applyFilters(b,s);const hidden=[...s.autoFilter.hiddenRows];sortFilterRange(b,s,2,true);
  const oracle=JSON.parse(fs.readFileSync(new URL('./fixtures/excel-filter-sort.json',import.meta.url),'utf8').replace(/^\uFEFF/,''));
  const e=createEvaluator(b,s.id),records=Array.from({length:9},(_,i)=>['B','C','D'].map(c=>s.cells[c+(i+4)]?.raw?e(c+(i+4)):null));
  assert.deepEqual(records,oracle.cases[0].records);assert.deepEqual(s.autoFilter.hiddenRows,hidden);
});

test('date checkbox values export date groups, collapse times to days and survive reapply',async()=>{
  const b=fixture(),s=b.sheets[0];for(let r=4;r<=12;r++)if(s.cells['C'+r])s.cells['C'+r].style={format:'date'};
  s.cells.C6.raw='10.5';assert.equal(filterOptions(b,s,2).values.filter(v=>v.key==='10').length,1);
  s.autoFilter.columns={2:{type:'values',kind:'date',values:['10','60']}};applyFilters(b,s);
  const out=await writeSpreadsheet(b),zip=await JSZip.loadAsync(out.data);assert.match(await zip.file('xl/worksheets/sheet1.xml').async('string'),/dateGroupItem year="1900" month="2" day="29"/);
  const re=await readSpreadsheet(out.data,'date-check.xlsx');applyFilters(re,re.sheets[0]);assert.deepEqual(re.sheets[0].autoFilter.hiddenRows,s.autoFilter.hiddenRows);
});
