import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {createWorkbook,evaluateFormula,rewriteWorkbookSheetReferences} from '../src/engine.mjs';
import {createHistory} from '../src/history.mjs';
import {createValidationRule,planValidationChange,applyValidationChange,validationAtSelection,validationBounds} from '../src/data-validation.mjs';
import {validationRules,validationChoices,validateCellUpdates,reviewValidation} from '../src/cell-constraints.mjs';
import {shiftWorksheetFeatures} from '../src/conditional.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {xml,all,attr} from '../src/chart-xml.mjs';
const fixture=()=>{const book=createWorkbook();return {book,sheet:book.sheets[0]};};
const set=(book,sheet,range,input)=>applyValidationChange(sheet,planValidationChange(book,sheet,range,input));
const issue=(book,sheet,key,raw)=>validateCellUpdates(book,sheet,[{key,cell:{raw}}]);

test('authored list, whole, decimal, date, time, text length and custom rules enforce their documented conditions',()=>{
  const {book,sheet}=fixture();sheet.cells={H1:{raw:'5'}};
  set(book,sheet,'A1:A5',{type:'list',source:'예,아니오',showInput:true,promptTitle:'승인',prompt:'목록을 선택하세요.'});
  assert.deepEqual(validationChoices(book,sheet,'A4').values,['예','아니오']);assert.equal(issue(book,sheet,'A1','예'),null);assert.equal(issue(book,sheet,'A1','보류').severity,'stop');
  assert.equal(validationRules(sheet)[0].prompt,'목록을 선택하세요.');
  set(book,sheet,'B1:B5',{type:'whole',operator:'between',formulae:['1','10']});
  for(const raw of ['1','10','5'])assert.equal(issue(book,sheet,'B2',raw),null);for(const raw of ['0','11','1.5',"'5"])assert.equal(issue(book,sheet,'B2',raw).severity,'stop');
  set(book,sheet,'C1:C5',{type:'decimal',operator:'greaterThan',formulae:['=$H$1'],severity:'warning'});
  assert.equal(issue(book,sheet,'C1','5.1'),null);assert.equal(issue(book,sheet,'C1','5').severity,'warning');
  for(const date1904 of [false,true]){
    book.date1904=date1904;set(book,sheet,'D1:D5',{type:'date',operator:'between',formulae:['2026-01-01','2026-12-31']});
    assert.equal(issue(book,sheet,'D4',String(evaluateFormula(book,sheet.id,'=DATE(2026,6,1)'))),null);
    assert.equal(issue(book,sheet,'D4',String(evaluateFormula(book,sheet.id,'=DATE(2025,12,31)'))).severity,'stop');
  }
  set(book,sheet,'E1:E5',{type:'time',operator:'between',formulae:['09:00','18:00']});assert.equal(issue(book,sheet,'E1','0.5'),null);assert.equal(issue(book,sheet,'E1','0.25').severity,'stop');
  set(book,sheet,'F1:F5',{type:'textLength',operator:'lessThanOrEqual',formulae:['3'],severity:'information'});assert.equal(issue(book,sheet,'F1','가나다'),null);assert.equal(issue(book,sheet,'F1','가나다라').severity,'information');
  set(book,sheet,'G1:G5',{type:'custom',formulae:['=AND(ISNUMBER(G1),G1>$H$1)']});assert.equal(issue(book,sheet,'G3','6'),null);assert.equal(issue(book,sheet,'G3','4').severity,'stop');
});
test('prompt-only validation maps standard none to unrestricted input and exports none, repairing the old any alias without changing the live document',async()=>{
  let {book,sheet}=fixture();set(book,sheet,'A1:A5',{type:'any',showInput:true,promptTitle:'자유 메모',prompt:'형식 제한 없이 입력하세요.'});assert.match(sheet.xlsxFeatures.dataValidations,/type="none"/);
  for(const type of ['none','any']){sheet.xlsxFeatures.dataValidations=sheet.xlsxFeatures.dataValidations.replace(/type="(?:none|any)"/,`type="${type}"`);assert.equal(validationRules(sheet)[0].type,'any');for(const raw of ['메모','123','TRUE',''])assert.equal(issue(book,sheet,'A1',raw),null);const before=JSON.stringify(book),data=(await writeSpreadsheet(book)).data;assert.equal(JSON.stringify(book),before);const zip=await JSZip.loadAsync(data);assert.match(await zip.file('xl/worksheets/sheet1.xml').async('string'),/type="none"/);assert.doesNotMatch(await zip.file('xl/worksheets/sheet1.xml').async('string'),/type="any"/);book=await readSpreadsheet(data,'prompt.xlsx');sheet=book.sheets[0];assert.equal(validationRules(sheet)[0].promptTitle,'자유 메모');assert.equal(validationRules(sheet)[0].type,'any');}
  delete sheet.xlsxFeatures;set(book,sheet,'A1',{type:'any'});assert.equal(sheet.xlsxFeatures,undefined,'unrestricted rule without guidance remains absent');
});
test('clearing or replacing part of a rule retains other rectangles and rebases relative formulas',()=>{
  const {book,sheet}=fixture();set(book,sheet,'B2:D5',{type:'custom',formulae:['B2>A2'],promptTitle:'유지',showInput:true,prompt:'안내'});
  const source=sheet.xlsxFeatures.dataValidations,plan=planValidationChange(book,sheet,'B2:C3',{type:'list',source:'x,y'});
  assert.equal(sheet.xlsxFeatures.dataValidations,source,'planning cannot mutate the live document');applyValidationChange(sheet,plan);
  let old=validationRules(sheet).find(r=>r.type==='custom');assert.deepEqual(old.anchor,{r:1,c:3});assert.equal(old.formulae[0],'D2>C2');
  assert.equal(issue(book,sheet,'D3','4'),null);sheet.cells.C3={raw:'5'};assert.equal(issue(book,sheet,'D3','4').severity,'stop');
  assert.deepEqual(validationChoices(book,sheet,'B2').values,['x','y']);assert.equal(validationChoices(book,sheet,'B4'),null);
  set(book,sheet,'C4:C4',null);old=validationRules(sheet).find(r=>r.type==='custom');assert.ok(old.ranges.length>=3);assert.equal(issue(book,sheet,'C4','-10'),null);assert.equal(issue(book,sheet,'D5','-10').severity,'stop');
  set(book,sheet,'A1:XFD1048576',null);assert.equal(sheet.xlsxFeatures.dataValidations,undefined);
});
test('a selection inside an existing rule shows relative operands anchored to that selection',()=>{
  const {book,sheet}=fixture();set(book,sheet,'B2:B10',{type:'whole',operator:'greaterThan',formulae:['A2']});
  const current=validationAtSelection(sheet,'B5:B7');assert.equal(current.mixed,false);assert.equal(current.rule.formulae[0],'A5');
  assert.equal(validationAtSelection(sheet,'B5:C7').mixed,true);assert.equal(validationAtSelection(sheet,'A1:C7').mixed,true);
});
test('whole-column rules remain compact and preserving unrelated opaque rules does not enumerate cells',()=>{
  const {book,sheet}=fixture();sheet.xlsxFeatures={dataValidations:'<dataValidations xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" disablePrompts="1"><dataValidation type="custom" sqref="Z1" allowBlank="1" customAttr="keep"><formula1>TRUE</formula1></dataValidation></dataValidations>'};
  set(book,sheet,'A1:A1048576',{type:'list',source:'예,아니오'});assert.equal(Object.keys(sheet.cells).length,0);assert.ok(sheet.xlsxFeatures.dataValidations.length<1000);assert.match(sheet.xlsxFeatures.dataValidations,/customAttr="keep"/);assert.match(sheet.xlsxFeatures.dataValidations,/disablePrompts="1"/);
  set(book,sheet,'A500000:A500001',null);assert.equal(validationChoices(book,sheet,'A500000'),null);assert.deepEqual(validationChoices(book,sheet,'A1000000').values,['예','아니오']);
});
test('rule editing is one undoable metadata change and failed authoring leaves no partial mutation',()=>{
  const {book}=fixture(),history=createHistory(book),sheet=history.book.sheets[0];
  const plan=planValidationChange(history.book,sheet,'A1:B9',{type:'whole',formulae:['1','10']});history.begin();applyValidationChange(sheet,plan);history.commit();assert.equal(validationRules(sheet).length,1);
  history.undo();assert.equal(history.book.sheets[0].xlsxFeatures,undefined);history.undo(true);assert.equal(validationRules(history.book.sheets[0]).length,1);
  const before=JSON.stringify(history.raw);assert.throws(()=>planValidationChange(history.book,history.book.sheets[0],'A1:B9',{type:'whole',formulae:['20','10']}),/最|최솟값/);assert.equal(JSON.stringify(history.raw),before);
});
test('protected sheets, macro workbooks, pivot results and invalid boundaries reject rule changes',()=>{
  const {book,sheet}=fixture();sheet.xlsxFeatures={sheetProtection:'<sheetProtection xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" sheet="1"/>'};assert.throws(()=>planValidationChange(book,sheet,'A1',{type:'list',source:'x'}),/보호/);delete sheet.xlsxFeatures;
  book.nativeSource={format:'xlsm'};assert.throws(()=>planValidationChange(book,sheet,'A1',{type:'list',source:'x'}),/XLSM/);delete book.nativeSource;
  sheet.pivot={};assert.throws(()=>planValidationChange(book,sheet,'A1',{type:'list',source:'x'}),/피벗/);delete sheet.pivot;
  for(const ref of ['A0','XFE1','A1048577','C3:A1','A1:B2:C3','Sheet1!A1'])assert.throws(()=>validationBounds(ref),/범위/);
  for(const input of [{type:'constructor'},{type:'whole',operator:'constructor',formulae:['1']},{type:'date',formulae:['2026-02-30','2026-03-01']},{type:'time',formulae:['25:00','25:01']},{type:'textLength',formulae:['-1','10']},{type:'list',source:'x'.repeat(256)},{type:'list',source:'x',promptTitle:'x'.repeat(33)}])assert.throws(()=>createValidationRule(input));
});
test('authored rules retain drop-down flags, prompts, warning types, relative references and namespaces through three XLSX roundtrips',async()=>{
  let {book,sheet}=fixture();const other={...createWorkbook().sheets[0],id:'lists',name:'도시 목록',cells:{A1:{raw:'서울'},A2:{raw:'부산'}}};book.sheets.push(other);
  set(book,sheet,'A2:A5',{type:'list',source:"='도시 목록'!$A$1:$A$2",showList:false,showInput:true,promptTitle:'도시 선택',prompt:'서울 또는 부산\n직접 입력',severity:'warning',title:'도시 확인',message:'목록의 도시를 사용하세요.'});
  set(book,sheet,'B2:B5',{type:'whole',operator:'greaterThan',formulae:['A2'],allowBlank:false});
  shiftWorksheetFeatures(sheet,'r',0,1,sheet.name);assert.equal(validationRules(sheet)[1].formulae[0],'A3');rewriteWorkbookSheetReferences(book,'도시 목록','국내 도시');other.name='국내 도시';
  for(let round=0;round<3;round++){
    const result=await writeSpreadsheet(book),zip=await JSZip.loadAsync(result.data),nodes=all(xml(await zip.file('xl/worksheets/sheet1.xml').async('string')),'dataValidation');
    assert.equal(attr(nodes[0],'sqref'),'A3:A6');assert.equal(attr(nodes[0],'showDropDown'),'1');assert.equal(attr(nodes[0],'errorStyle'),'warning');assert.equal(attr(nodes[0],'prompt'),'서울 또는 부산\n직접 입력');
    const independent=new ExcelJS.Workbook();await independent.xlsx.load(result.data);assert.equal(independent.worksheets[0].getCell('A3').dataValidation.type,'list');assert.equal(independent.worksheets[0].getCell('A3').dataValidation.promptTitle,'도시 선택');
    book=await readSpreadsheet(result.data,'validation.xlsx');sheet=book.sheets[0];assert.match(validationRules(sheet)[0].formulae[0],/국내 도시/);assert.equal(validationChoices(book,sheet,'A3'),null);assert.equal(issue(book,sheet,'A3','대구').severity,'warning');assert.equal(validationRules(sheet)[1].formulae[0],'A3');
  }
});

test('review finds pasted/formula invalid values and silent alerts without changing values, including positional custom formulas',()=>{
  const {book,sheet}=fixture();sheet.cells={A1:{raw:'5'},A2:{raw:'99'},A3:{raw:'=50+50'},A4:{raw:''},B1:{raw:'1'},B2:{raw:'1'},B3:{raw:'1'}};
  set(book,sheet,'A1:A1048576',{type:'whole',formulae:['1','10'],showError:false});set(book,sheet,'B1:B3',{type:'custom',formulae:['=ROW()=1']});
  const before=JSON.stringify(book),review=reviewValidation(book,sheet,validationBounds('A1:B20'));
  assert.equal(review.checked,6);assert.equal(review.invalid,4);assert.deepEqual(review.results.map(r=>r.ref),['A2','A3','B2','B3']);assert.equal(JSON.stringify(book),before);
  const limited=reviewValidation(book,sheet,validationBounds('A1:B20'),{maxCells:2,limit:1});assert.equal(limited.checked,2);assert.equal(limited.results.length,1);assert.equal(limited.truncated,true);
  assert.equal(issue(book,sheet,'A2','99'),null,'silent alerts remain silent during direct entry');
});
