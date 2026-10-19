import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {createRequire} from 'node:module';
import {borderPatch,cellBorders,bordersCss,validateBorders,borderSummary} from '../src/borders.mjs';
import {createWorkbook,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {attachNativeSource,buildNativePlan} from '../src/native-plan.mjs';
const vba=createRequire(import.meta.url)('../electron/vba.cjs');
const range={r1:0,r2:1,c1:0,c2:1},edge={style:'medium',color:'#ff0000'};
const patch=(r,c,decisions,style={})=>borderPatch(style,{r1:r,r2:r,c1:c,c2:c},range,decisions,edge.style,edge.color);
test('range outlines, inner edges, individual sides and merged cells retain separate borders',()=>{
  const outline={top:true,bottom:true,left:true,right:true,insideH:false,insideV:false};
  assert.deepEqual(cellBorders(patch(0,0,outline)),{top:edge,right:null,bottom:null,left:edge});
  assert.deepEqual(cellBorders(patch(1,1,outline)),{top:null,right:edge,bottom:edge,left:null});
  const inside={insideH:true,insideV:true};
  assert.deepEqual(cellBorders(patch(0,0,inside)),{top:null,right:edge,bottom:edge,left:null});
  assert.deepEqual(cellBorders(patch(1,1,inside)),{top:edge,right:null,bottom:null,left:edge});
  assert.deepEqual(cellBorders(borderPatch({},range,range,outline,'double','#0070c0')),{top:{style:'double',color:'#0070c0'},right:{style:'double',color:'#0070c0'},bottom:{style:'double',color:'#0070c0'},left:{style:'double',color:'#0070c0'}});
  const old={border:true,borders:{bottom:{style:'dashed',color:'#0000ff'}}};
  const changed=patch(1,0,{bottom:false},old);assert.equal(changed.borders.bottom,null);assert.deepEqual(changed.borders.left,{style:'thin',color:'#8a8a8a'});assert.ok(bordersCss(changed).includes('border-left:1px solid #8a8a8a;'));
});
test('malformed persisted border settings are rejected before rendering/export',()=>{
  for(const invalid of [null,[],{diagonal:null},{top:{style:'unknown',color:'#000000'}},{top:{style:'thin',color:'red;'}},{top:{style:'thin',color:'#000000',script:'x'}}]){
    assert.throws(()=>validateBorders(invalid),/테두리/);const b=createWorkbook();b.sheets[0].cells.A1={raw:'',style:{borders:invalid}};assert.throws(()=>serializeWorkbook(b),/테두리/);
  }
});
test('JSON and XLSX preserve border sides, styles, colors and old all-borders files',async()=>{
  const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'report',style:{borders:{top:edge,right:{style:'double',color:'#0070c0'},bottom:{style:'dotted',color:'#00b050'},left:null},verticalAlign:'bottom'}};s.cells.B2={raw:'legacy',style:{border:true}};s.cells.A5={raw:'',style:{borders:{bottom:edge}}};s.cells.C20={raw:'',style:{borders:{right:edge}}};
  assert.deepEqual(parseWorkbook(serializeWorkbook(b)).sheets[0].cells.A1,s.cells.A1);
  const output=await writeSpreadsheet(b),excel=new ExcelJS.Workbook();await excel.xlsx.load(output.data);const actual=excel.worksheets[0].getCell('A1');
  assert.deepEqual(actual.border.top,{style:'medium',color:{argb:'FFFF0000'}});assert.equal(actual.border.right.style,'double');assert.equal(actual.border.bottom.style,'dotted');assert.equal(actual.border.left?.style,undefined);assert.equal(actual.alignment.vertical,'bottom');assert.equal(excel.worksheets[0].getCell('B2').border.left.style,'thin');
  const round=await readSpreadsheet(output.data,'borders.xlsx');assert.equal(round.sheets[0].cells.A5?.style?.borders?.bottom?.style,'medium');assert.equal(round.sheets[0].cells.C20?.style?.borders?.right?.style,'medium');for(const side of ['top','right','bottom','left'])assert.deepEqual(round.sheets[0].cells.A1.style.borders[side],s.cells.A1.style.borders[side]&&{...s.cells.A1.style.borders[side],color:s.cells.A1.style.borders[side].color.toUpperCase()});
  // Re-exporting a one-sided imported border must not expand it to all four sides.
  const check=new ExcelJS.Workbook();await check.xlsx.load((await writeSpreadsheet(round)).data);assert.equal(check.worksheets[0].getCell('A1').border.left?.style,undefined);
});
test('native Excel save plans validate per-edge settings and explicit removal',async()=>{
  const b=createWorkbook();b.sheets[0].cells.A1={raw:'1'};attachNativeSource(b,(await writeSpreadsheet(b)).data,'borders.xlsm');
  b.sheets[0].cells.A1.style={border:false,borders:{top:edge,bottom:null,right:null,left:null}};let plan=buildNativePlan(b);vba.validatePayload({mode:'save',source:b.nativeSource,plan});assert.deepEqual(plan.sheets[0].cells[0].style.borders.top,edge);
  plan.sheets[0].cells[0].style.borders.top.color='invalid';assert.throws(()=>vba.validatePayload({mode:'save',source:b.nativeSource,plan}),/테두리/);
  plan.sheets[0].cells[0].style.borders=null;vba.validatePayload({mode:'save',source:b.nativeSource,plan});
});

test('border summary distinguishes all, none and mixed edges without changing them',()=>{
  const targets=[{cell:{style:patch(0,0,{top:true})},rect:{r1:0,r2:0,c1:0,c2:0}},{cell:{},rect:{r1:0,r2:0,c1:1,c2:1}}];
  assert.equal(borderSummary(targets,range).top,null);assert.equal(borderSummary(targets,range).insideV,false);
});
