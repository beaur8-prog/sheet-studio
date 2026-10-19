import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {colorFixture} from './fixtures/filter-colors.mjs';
import {applyFilters,filterOptions,shiftAutoFilter} from '../src/filters.mjs';
import {parseWorkbook,serializeWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {nativeFilterPlan} from '../src/filter-native.mjs';
import {createColorResolver} from '../src/spreadsheet-colors.mjs';
import {xml,child,children,stringify} from '../src/chart-xml.mjs';
const rule=(target,color)=>({type:'color',target,color});
const visible=s=>Array.from({length:7},(_,i)=>i+1).filter(r=>!s.autoFilter.hiddenRows.includes(r));

test('fill and font filters use displayed conditional colors, normalize hex and include formatted blanks',()=>{
 const b=colorFixture(),s=b.sheets[0];
 for(const [target,color,expected] of [['background','#ff0000',[1,3,6,7]],['background','#0000ff',[2]],['background',null,[4]],['background','#ffffff',[5]],['color','#008000',[1,6,7]],['color','#ff0000',[2,4]],['color',null,[3]],['color','#000000',[5]]]){
  s.autoFilter.columns={0:rule(target,color)};applyFilters(b,s);assert.deepEqual(visible(s),expected,target+':'+color);
 }
 s.autoFilter.columns={0:rule('background','#ff0000')};applyFilters(b,s);const before=[...s.autoFilter.hiddenRows];s.cells.B8.raw='1';assert.deepEqual(s.autoFilter.hiddenRows,before);applyFilters(b,s);assert.deepEqual(visible(s),[1,3,6]);
});
test('color criteria combine with other columns and their menus omit only their own condition',()=>{
 const b=colorFixture(),s=b.sheets[0];s.autoFilter.columns={0:rule('background','#ff0000'),1:{type:'custom',kind:'number',join:'and',conditions:[{operator:'greaterThan',value:'20'}]}};applyFilters(b,s);assert.deepEqual(visible(s),[3,6,7]);
 assert.deepEqual(filterOptions(b,s,0).colors.background,[{color:null,count:1},{color:'#ff0000',count:3},{color:'#ffffff',count:1}]);assert.deepEqual(filterOptions(b,s,1).values.map(v=>v.value),[10,30,60,70]);
 const copy=parseWorkbook(serializeWorkbook(b));assert.deepEqual(copy.sheets[0].autoFilter,s.autoFilter);shiftAutoFilter(copy.sheets[0],'c',0,1);assert.equal(copy.sheets[0].autoFilter.columns[1].type,'color');
 for(const bad of [{target:'background',color:'red'},{target:'color',color:undefined},{target:'border',color:'#ff0000'}]){const badBook=colorFixture();badBook.sheets[0].autoFilter.columns={0:{type:'color',...bad}};assert.throws(()=>parseWorkbook(JSON.stringify(badBook)),/필터/);}
});
test('native XLSX color filters share appended differential styles without disturbing conditional rules',async()=>{
 for(const [target,color] of [['background','#ff0000'],['background',null],['color','#008000'],['color',null]]){
  let b=colorFixture(),s=b.sheets[0];s.autoFilter.columns={0:rule(target,color)};applyFilters(b,s);const expected=visible(s);
  for(let i=0;i<2;i++){
   const out=await writeSpreadsheet(b),zip=await JSZip.loadAsync(out.data),sheet=await zip.file('xl/worksheets/sheet1.xml').async('string'),styles=await zip.file('xl/styles.xml').async('string');
   assert.match(sheet,new RegExp(`<colorFilter dxfId="1" cellColor="${target==='background'?1:0}"`));assert.match(sheet,/cfRule[^>]*dxfId="0"/);assert.match(styles,/<dxfs count="2"/);
   b=await readSpreadsheet(out.data,'colors.xlsx');s=b.sheets[0];assert.deepEqual(s.autoFilter.columns[0],rule(target,color));assert.equal(s.autoFilter.unsupported,undefined);applyFilters(b,s);assert.deepEqual(visible(s),expected);
  }
 }
});
test('Excel color import resolves actual theme, tint and indexed values consistently for cells and criteria',async()=>{
 const b=colorFixture(),s=b.sheets[0];s.autoFilter.columns={0:rule('background','#ff0000')};applyFilters(b,s);const zip=await JSZip.loadAsync((await writeSpreadsheet(b)).data);
 let styles=await zip.file('xl/styles.xml').async('string');styles=styles.replaceAll('rgb="FFFF0000"','theme="4" tint="0.4"');zip.file('xl/styles.xml',styles);
 const loaded=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'theme.xlsx'),sheet=loaded.sheets[0];assert.equal(sheet.cells.A2.style.background,'#95b3d7');assert.equal(sheet.conditionalRules[0].style.background,'#95b3d7');assert.equal(sheet.autoFilter.columns[0].color,'#95b3d7');assert.deepEqual(visible(sheet),[1,3,6,7]);
 const resolve=createColorResolver('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:clrScheme><a:dk1><a:sysClr val="windowText" lastClr="112233"/></a:dk1><a:accent1><a:srgbClr val="804020"/></a:accent1></a:clrScheme></a:theme>');assert.equal(resolve({theme:1}),'#112233');assert.equal(resolve({theme:4}),'#804020');assert.equal(resolve({indexed:10}),'#ff0000');assert.equal(resolve({auto:true}),undefined);
});
test('Windows native plans use OLE RGB and the corresponding fill/font operators',()=>{
 const b=colorFixture(),s=b.sheets[0];for(const [target,color,operator,criteria1] of [['background','#123456',8,0x563412],['color','#123456',9,0x563412],['background',null,12,undefined],['color',null,13,undefined]]){s.autoFilter.columns={0:rule(target,color)};const p=nativeFilterPlan(s.autoFilter).filters[0];assert.equal(p.field,1);assert.equal(p.operator,operator);assert.equal(p.criteria1,criteria1);}
});
test('unknown color style references and gradient criteria preserve saved hidden rows and block destructive reapply',async()=>{
 const b=colorFixture(),s=b.sheets[0];s.autoFilter.columns={0:rule('background','#ff0000')};applyFilters(b,s);const bytes=(await writeSpreadsheet(b)).data;
 for(const invalid of ['missing','gradient']){
  const zip=await JSZip.loadAsync(bytes);
  if(invalid==='missing')zip.file('xl/worksheets/sheet1.xml',(await zip.file('xl/worksheets/sheet1.xml').async('string')).replace(/<colorFilter dxfId="\d+"/,'<colorFilter dxfId="99999"'));
  else {const doc=xml(await zip.file('xl/styles.xml').async('string')),dxf=children(child(doc.documentElement,'dxfs'),'dxf')[1];dxf.replaceChild(doc.importNode(xml('<fill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><gradientFill degree="0"><stop position="0"><color rgb="FFFF0000"/></stop></gradientFill></fill>').documentElement,true),child(dxf,'fill'));zip.file('xl/styles.xml',stringify(doc));}
  const loaded=await readSpreadsheet(await zip.generateAsync({type:'uint8array'}),'unknown-color.xlsx'),sheet=loaded.sheets[0];assert.deepEqual(sheet.autoFilter.unsupported,[0],invalid);assert.deepEqual(visible(sheet),[1,3,6,7]);assert.throws(()=>applyFilters(loaded,sheet),/미지원/);await assert.rejects(()=>writeSpreadsheet(loaded),/미지원/);
 }
});
