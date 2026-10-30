import test from 'node:test';
import assert from 'node:assert/strict';
import {worksheetMetadataXML, splitWorksheetXML} from '../src/worksheet-metadata.mjs';
import {xml, all, attr, stringify} from '../src/chart-xml.mjs';

test('metadata scanning drops cell DOMs, keeps hidden rows and preserves original cell bytes on export', () => {
  const data='<s:sheetData><s:row r="1"><s:c r="A1" t="inlineStr"><s:is><s:t>&lt;sheetData&gt;한글</s:t></s:is></s:c></s:row><s:row r="2" hidden="1" ht="25"><s:c r="A2"><s:v>2</s:v></s:c></s:row></s:sheetData>';
  const input='<?xml version="1.0"?>\n<s:worksheet xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+data+'<s:autoFilter ref="A1:A2"/></s:worksheet>\n';
  const parts=splitWorksheetXML(input),doc=xml(parts.metadata),hidden=[];
  worksheetMetadataXML(input,{onRow:row=>hidden.push(row)});
  assert.equal(all(doc,'c').length,0);assert.equal(all(doc,'row').length,0);assert.equal(hidden[0].r,'2');
  all(doc,'autoFilter')[0].setAttribute('ref','A1:A3');
  const output=parts.restore(stringify(doc));assert.ok(output.includes(data));assert.equal(attr(all(xml(output),'autoFilter')[0],'ref'),'A1:A3');
  assert.throws(()=>worksheetMetadataXML('<!DOCTYPE x><x/>'),/DOCTYPE/);
});

test('selected date cells stream their values without materializing unrelated cells', () => {
  const cells=[];
  const result=worksheetMetadataXML('<worksheet><sheetData><row r="1"><c r="A1" t="d"><v>2026-10-20</v></c><c r="B1"><f>A1</f><v>42</v></c><c r="C1"><v>99</v></c></row></sheetData></worksheet>',{wantCell:cell=>cell.r!=='C1',onCell:cell=>cells.push(cell)});
  assert.deepEqual(cells.map(cell=>[cell.r,cell.value,cell.formula]),[['A1','2026-10-20',false],['B1','42',true]]);
  assert.equal(all(xml(result),'c').length,0);
});

test('streamed text metadata distinguishes empty inline strings, whitespace, CDATA and absent formula values',()=>{
 const cells=[];
 worksheetMetadataXML('<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t/></is></c><c r="B1" t="inlineStr"><is><r><t> </t></r></is></c><c r="C1" t="inlineStr"><is><t><![CDATA[x]]></t></is></c><c r="D1" t="str"><f>FUTURE()</f><v/></c><c r="E1" t="str"><f>FUTURE()</f></c></row></sheetData></worksheet>',{wantCell:()=>true,onCell:cell=>cells.push(cell)});
 assert.deepEqual(cells.map(c=>[c.r,c.inlineEmpty,c.stringValuePresent,c.formula]),[['A1',true,undefined,false],['B1',false,undefined,false],['C1',false,undefined,false],['D1',undefined,true,true],['E1',undefined,undefined,true]]);
});
