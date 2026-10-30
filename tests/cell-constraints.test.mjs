import {createHistory} from '../src/history.mjs';
import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {createWorkbook,evaluateFormulaArray,rewriteWorkbookSheetReferences,transformFormulaReferences} from '../src/engine.mjs';
import {cellIsLocked,cellFormulaHidden,cellEditError,assertCellUpdates,assertStructureAllowed,assertSortAllowed,validateCellUpdates,validationChoices} from '../src/cell-constraints.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {shiftWorksheetFeatures} from '../src/conditional.mjs';
import {xml,all,attr} from '../src/chart-xml.mjs';

async function validationFixture(){
  const source=new ExcelJS.Workbook(),s=source.addWorksheet('Input'),list=source.addWorksheet('Lists');
  list.getCell('A1').value='서울';list.getCell('A2').value='부산';list.getCell('A3').value='대구';
  const base={showErrorMessage:true,errorStyle:'stop',errorTitle:'조건 확인',error:'허용된 값을 입력하세요.'};
  s.getCell('A1').dataValidation={...base,type:'list',formulae:['"예,아니오"']};
  s.getCell('B2').dataValidation={...base,type:'whole',operator:'between',formulae:[1,10]};
  s.getCell('C2').dataValidation={...base,type:'custom',formulae:['AND(ISNUMBER(C2),C2>B2)']};
  s.getCell('D2').dataValidation={...base,type:'list',formulae:["'Lists'!$A$1:$A$3"]};
  s.getCell('E2').dataValidation={...base,type:'textLength',operator:'lessThanOrEqual',formulae:[3],errorStyle:'warning'};
  s.getCell('B2').value=5;s.getCell('C2').value=6;s.getCell('D2').value='서울';return readSpreadsheet(await source.xlsx.writeBuffer(),'validation.xlsx');
}
test('validation covers inline/range lists, numeric bounds, custom expressions and warning rules',async()=>{
  const book=await validationFixture(),s=book.sheets[0],change=(key,raw)=>({key,cell:{raw}});
  assert.deepEqual(validationChoices(book,s,'A1').values,['예','아니오']);assert.deepEqual(validationChoices(book,s,'D2').values,['서울','부산','대구']);
  for(const [key,raw] of [['A1','예'],['B2','10'],['C2','8'],['D2','부산'],['E2','가나다']])assert.equal(validateCellUpdates(book,s,[change(key,raw)]),null,key);
  for(const [key,raw] of [['A1','그 외'],['B2','11'],['B2','1.5'],['C2','3'],['D2','광주']])assert.equal(validateCellUpdates(book,s,[change(key,raw)])?.severity,'stop',key);
  assert.equal(validateCellUpdates(book,s,[change('E2','가나다라')]).severity,'warning');
  // A custom rule sees the entire pending edit, including its dependent cell.
  assert.equal(validateCellUpdates(book,s,[change('B2','8'),change('C2','9')]),null);assert.equal(validateCellUpdates(book,s,[change('B2','8'),change('C2','7')]).ref,'C2');
  assert.deepEqual(evaluateFormulaArray(book,s.id,"='Lists'!A1:A3").values,['서울','부산','대구']);
});
test('protected cells, unlocked row/column defaults and hidden formulas retain their permissions',async()=>{
  const source=new ExcelJS.Workbook(),s=source.addWorksheet('Protected');s.getCell('A1').value='locked';s.getCell('B1').value=1;s.getCell('B1').protection={locked:false};s.getCell('C1').value={formula:'B1*2',result:2};s.getCell('C1').protection={hidden:true};s.getColumn(4).protection={locked:false};s.getRow(5).protection={locked:false};s.getCell('A5').value='row template';await s.protect('secret',{spinCount:1000,formatCells:true});
  const book=await readSpreadsheet(await source.xlsx.writeBuffer(),'protected.xlsx'),sheet=book.sheets[0];
  assert.equal(cellIsLocked(book,sheet,'A1'),true);assert.equal(cellIsLocked(book,sheet,'B1'),false);assert.equal(cellIsLocked(book,sheet,'D20'),false);assert.equal(cellIsLocked(book,sheet,'B5'),false);assert.equal(cellFormulaHidden(book,sheet,'C1'),true);
  assert.throws(()=>assertCellUpdates(book,sheet,[{key:'B1'},{key:'A1'}]),/보호/);assert.doesNotThrow(()=>assertCellUpdates(book,sheet,[{key:'B1'},{key:'D20'}]));assert.throws(()=>assertStructureAllowed(sheet,'r',1,1,1),/보호/);assert.throws(()=>assertSortAllowed(book,sheet,{r1:0,r2:1,c1:0,c2:1}),/보호/);
  sheet.cells.B1.raw='4';sheet.cells.D20={raw:'new'};const out=new ExcelJS.Workbook();await out.xlsx.load((await writeSpreadsheet(book)).data);assert.equal(out.worksheets[0].getCell('B1').value,4);assert.equal(out.worksheets[0].getColumn(4).protection.locked,false);assert.equal(out.worksheets[0].getRow(5).protection.locked,false);assert.ok(out.worksheets[0].sheetProtection.hashValue);assert.equal(out.worksheets[0].getCell('D20').protection.locked,false);
});
test('array results cannot be partly edited, filled, sorted or structurally cut',()=>{
  const book=createWorkbook(),s=book.sheets[0];s.cells={A1:{raw:'=SUM(B1:B2)',xlsxFormula:{raw:'=SUM(B1:B2)',formula:'SUM(B1:B2)',shareType:'array',ref:'A1:A3'}},A2:{raw:'2'},A3:{raw:'3'}};
  assert.match(cellEditError(book,s,'A2'),/배열/);assert.equal(cellEditError(book,s,'A1'),null);
  assert.throws(()=>assertCellUpdates(book,s,[{key:'A2'}]),/배열/);assert.doesNotThrow(()=>assertCellUpdates(book,s,[{key:'A1'}],{editAnchor:true}));assert.doesNotThrow(()=>assertCellUpdates(book,s,[{key:'A1'},{key:'A2'},{key:'A3'}]));
  assert.throws(()=>assertStructureAllowed(s,'r',1,1,1),/배열/);assert.throws(()=>assertStructureAllowed(s,'r',1,1,-1),/배열/);assert.doesNotThrow(()=>assertStructureAllowed(s,'r',0,3,-1));assert.throws(()=>assertSortAllowed(book,s,{r1:0,r2:2,c1:0,c2:1}),/배열/);
});
test('preserved validation/array references follow insertion, deletion and sheet renaming',async()=>{
  const book=await validationFixture(),s=book.sheets[0];s.cells.F2={raw:'=SUM(B2:B3)',xlsxFormula:{raw:'=SUM(B2:B3)',formula:'SUM(B2:B3)',shareType:'array',ref:'F2:F3'}};
  shiftWorksheetFeatures(s,'r',0,1,s.name);
  let rules=all(xml(s.xlsxFeatures.dataValidations),'dataValidation');const custom=rules.find(n=>attr(n,'type')==='custom');assert.equal(attr(custom,'sqref'),'C3');assert.equal(all(custom,'formula1')[0].textContent,'AND(ISNUMBER(C3),C3>B3)');assert.equal(s.cells.F2.xlsxFormula.ref,'F3:F4');
  shiftWorksheetFeatures(s,'r',0,-1,s.name);assert.equal(s.cells.F2.xlsxFormula.ref,'F2:F3');
  rewriteWorkbookSheetReferences(book,'Lists','도시 목록');assert.ok(s.xlsxFeatures.dataValidations.includes("'도시 목록'!$A$1:$A$3"));
  assert.equal(transformFormulaReferences('=SUM(A1:A1048576)',{axis:'r',index:0,delta:1,targetSheet:s.name,ownerSheet:s.name}),'=SUM(A2:A1048576)');
  const saved=await JSZip.loadAsync((await writeSpreadsheet(book)).data);rules=all(xml(await saved.file('xl/worksheets/sheet1.xml').async('string')),'dataValidation');assert.equal(attr(rules.find(n=>attr(n,'type')==='whole'),'sqref'),'B2');
});

test('validation never changes the live history-backed workbook before acceptance',async()=>{
  const book=await validationFixture(),history=createHistory(book),sheet=history.book.sheets[0];history.begin();
  const failure=validateCellUpdates(history.book,sheet,[{key:'B2',cell:{raw:'99'}}]);assert.equal(failure.severity,'stop');assert.equal(sheet.cells.B2.raw,'5');assert.equal(history.takeDirty().size,0);history.rollback();
  assert.equal(validateCellUpdates(history.book,sheet,[{key:'B2',cell:{raw:'7'}}]),null);assert.equal(sheet.cells.B2.raw,'5');
});
