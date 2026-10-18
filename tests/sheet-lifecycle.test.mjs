import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,rewriteSheetReferences,rewriteWorkbookSheetReferences,evaluateCell,serializeWorkbook,parseWorkbook} from '../src/engine.mjs';

test('rename preserves strings and updates quoted, Unicode, locked and repeated references',()=>{
 assert.equal(rewriteSheetReferences('=SUM(資料!$A$1:B2)+資料!C3+"資料!A1"','資料',"O'Brien"),'=SUM(\'O\'\'Brien\'!$A$1:B2)+\'O\'\'Brien\'!C3+"資料!A1"');
 assert.equal(rewriteSheetReferences("='O''Brien'!A1", "O'Brien", 'New'),"='New'!A1");
 assert.equal(rewriteSheetReferences('=INDIRECT("Source!A1")+Other!A1','Source',null),'=INDIRECT("Source!A1")+Other!A1');
 assert.equal(rewriteSheetReferences('=SUM(source!$A$1:B2)+Source!C3','Source',null),'=SUM(#REF!)+#REF!');
 assert.equal(rewriteSheetReferences('=SUM(Source!A1:Source!B2)','Source',null),'=SUM(#REF!)');
 const many='='+Array(150).fill('Source!A1').join('+');
 assert.equal(rewriteSheetReferences(many,'Source',null),'='+Array(150).fill('#REF!').join('+'));
});

test('deleted sheet references stay invalid after recreation and JSON roundtrip',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'=Source!A1'};
 const source={...structuredClone(s),id:'source',name:'Source',cells:{A1:{raw:'42'}}};b.sheets.push(source);
 assert.equal(evaluateCell(b,s.id,'A1'),42);
 s.conditionalRules=[{id:'cf',type:'expression',range:'A1:A1',formulae:['Source!A1=42'],style:{background:'#ff0000'}}];
 const undo=serializeWorkbook(b);
 rewriteWorkbookSheetReferences(b,'Source',null);b.sheets.pop();
 b.sheets.push({...source,id:'new',cells:{A1:{raw:'999'}}});
 const round=parseWorkbook(serializeWorkbook(b));
 assert.equal(evaluateCell(round,s.id,'A1'),'#REF!');
 assert.equal(round.sheets[0].conditionalRules[0].formulae[0],'#REF!=42');
 assert.equal(evaluateCell(parseWorkbook(undo),s.id,'A1'),42);
});

test('renaming updates cell and conditional formulas throughout the workbook',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'1'};s.cells.B1={raw:'=Sheet1!A1'};
 s.conditionalRules=[{formulae:['Sheet1!A1=1','"Sheet1!A1"']}];
 rewriteWorkbookSheetReferences(b,'Sheet1','Renamed');s.name='Renamed';
 assert.equal(evaluateCell(b,s.id,'B1'),1);
 assert.deepEqual(s.conditionalRules[0].formulae,["'Renamed'!A1=1",'"Sheet1!A1"']);
});
