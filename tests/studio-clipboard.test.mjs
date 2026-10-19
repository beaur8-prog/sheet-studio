import test from 'node:test';
import assert from 'node:assert/strict';
import {tableToClipboard,tableText,colorHex} from '../src/studio-table.mjs';
import {copyTable} from '../src/table-exchange.mjs';
test('rich spreadsheet copy carries display values, styles, merged cells and quoted TSV without formulas',()=>{
 const rich=copyTable({r:0,c:0,merges:['A1:B1'],data:[[{raw:'보고서',style:{bold:true,background:'#fff2cc',borders:{bottom:{style:'double',color:'#ff0000'}}}},{raw:''}],[{raw:'=SUM(C2:D2)',style:{format:'currency',decimals:0}},{raw:"'001",style:{color:'#0070c0'}}]],values:[['보고서',''],[1234,'001']]},{columnWidths:[120,80],rowHeights:[35,25]});
 assert.ok(rich.html.includes('colspan="2"'));assert.ok(rich.html.includes('3px double #ff0000'));assert.ok(rich.html.includes('background-color:#fff2cc'));assert.ok(rich.html.includes('width:120px'));assert.ok(rich.html.includes('₩1,234'));assert.ok(!rich.html.includes('SUM('));assert.ok(rich.text.includes('001'));
});
test('HTML and TSV escape literal content and dangerous CSS, retaining cell line breaks',()=>{
 const rich=tableToClipboard({cells:[[{text:'<img src=x onerror=alert(1)>\n"Q"\tX',style:{fontFamily:'Arial;background:url(x)',color:'red;url(x)'}}]]});assert.ok(!rich.html.includes('<img'));assert.ok(!rich.html.includes('url('));assert.ok(rich.html.includes('&lt;img'));assert.ok(rich.html.includes('<br>'));assert.equal(rich.text,'"<img src=x onerror=alert(1)>\n""Q""\tX"');assert.equal(tableText([[{text:'a'},{text:'b\tc'}]]),'a\t"b\tc"');assert.equal(colorHex('rgb(255, 0, 128)'),'#ff0080');
});
