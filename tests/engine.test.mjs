import test from 'node:test';
import assert from 'node:assert/strict';
import { address, colName, colIndex, parseAddress, createWorkbook, evaluateCell, evaluateFormula, createEvaluator, formatValue, shiftFormula, transformFormulaReferences, parseCSV, toCSV, serializeWorkbook, parseWorkbook } from '../src/engine.mjs';

function book(cells = {}) {
  const wb = createWorkbook();
  wb.sheets[0].cells = Object.fromEntries(Object.entries(cells).map(([ref, raw]) => [ref, { raw: String(raw), style: {} }]));
  return wb;
}
const calc = (formula, cells = {}) => evaluateFormula(book(cells), 'sheet-1', formula);

test('address conversion covers full Excel grid and rejects invalid positions', () => {
  assert.equal(colName(0), 'A'); assert.equal(colName(25), 'Z'); assert.equal(colName(26), 'AA'); assert.equal(colName(16383), 'XFD');
  assert.equal(colIndex('XFD'), 16383); assert.equal(colIndex('XFE'), -1);
  assert.deepEqual(parseAddress('$BC$100'), { r: 99, c: 54 });
  assert.equal(address(1048575, 16383), 'XFD1048576'); assert.equal(address(-1, 0), '#REF!');
  for (const bad of ['A0', 'A01', 'XFE1', 'A1048577', '', 'AAA', '1']) assert.equal(parseAddress(bad), null);
});

test('operators honor Excel precedence, comparison types, percentages, and quoted strings', () => {
  assert.equal(calc('=2+3*4^2'), 50);
  assert.equal(calc('=(2+3)*4'), 20);
  assert.equal(calc('=-2^2'), 4);
  assert.equal(calc('=2^3^2'), 64);
  assert.equal(calc('=200*10%'), 20);
  assert.equal(calc('="He said ""hi"""&"!"'), 'He said "hi"!');
  assert.equal(calc('="ABC"="abc"'), true);
  assert.equal(calc('=1="1"'), false);
  assert.equal(calc('=1<"a"'), true);
  assert.equal(calc('=5>=5'), true);
});

test('cell references recalculate and empty references behave consistently', () => {
  const wb = book({ A1: '10', A2: '20', B1: '=SUM(A1:A2)', B2: '=B1*2', C1: "'003", C2: '12%', C3: '=D1', C4: '=""' });
  assert.equal(evaluateCell(wb, 'sheet-1', 'B2'), 60);
  wb.sheets[0].cells.A1.raw = '15';
  assert.equal(evaluateCell(wb, 'sheet-1', 'B2'), 70);
  assert.equal(evaluateCell(wb, 'sheet-1', 'C1'), '003');
  assert.equal(evaluateCell(wb, 'sheet-1', 'C2'), 0.12);
  assert.equal(evaluateCell(wb, 'sheet-1', 'C3'), 0);
  assert.equal(evaluateCell(wb, 'sheet-1', 'C4'), '');
  assert.equal(evaluateCell(wb, 'sheet-1', 'D1'), '');
  assert.equal(evaluateFormula(wb, 'sheet-1', '=D1=""'), true);
  assert.equal(evaluateFormula(wb, 'sheet-1', '=D1=0'), true);
  assert.equal(createEvaluator(wb, 'sheet-1')('B2'), 70);
});

test('aggregates distinguish numeric text, booleans, empty cells, and error values', () => {
  const cells = { A1: 10, A2: 20, A3: "'30", A4: 'TRUE', A5: '=""', A6: '' };
  assert.equal(calc('=SUM(A1:A6)', cells), 30);
  assert.equal(calc('=SUM(A1:A6,"30",TRUE)', cells), 61);
  assert.equal(calc('=AVERAGE(A1:A6)', cells), 15);
  assert.equal(calc('=MIN(A1:A6)', cells), 10);
  assert.equal(calc('=MAX(A1:A6)', cells), 20);
  assert.equal(calc('=COUNT(A1:A6)', cells), 2);
  assert.equal(calc('=COUNT(A1:A6,"30",TRUE)', cells), 4);
  assert.equal(calc('=COUNT("")'), 0);
  assert.equal(calc('=COUNTA(A1:A6)', cells), 5);
  assert.equal(calc('=AVERAGE(Z1:Z5)'), '#DIV/0!');
  assert.equal(calc('=SUM("bad")'), '#VALUE!');
  assert.equal(calc('=SUM(A1:A2)', { A1: 1, A2: '=1/0' }), '#DIV/0!');
  assert.equal(calc('=COUNT(A1:A2)', { A1: 1, A2: '=1/0' }), 1);
  assert.equal(calc('=COUNTA(A1:A2)', { A1: 1, A2: '=1/0' }), 2);
});

test('IF and IFERROR evaluate branches lazily and logical functions inspect ranges', () => {
  assert.equal(calc('=IF(TRUE,42,1/0)'), 42);
  assert.equal(calc('=IF(FALSE,UNKNOWN(),7)'), 7);
  assert.equal(calc('=IFERROR(1/0,"unavailable")'), 'unavailable');
  assert.equal(calc('=IFERROR(7,1/0)'), 7);
  assert.equal(calc('=IF(2<3,"yes","no")'), 'yes');
  assert.equal(calc('=IF(FALSE,1)'), false);
  assert.equal(calc('=AND(TRUE,3>2,NOT(FALSE))'), true);
  assert.equal(calc('=OR(FALSE,0,2)'), true);
  assert.equal(calc('=AND(A1:A3)', { A1: 'text', A2: 2, A3: 1 }), true);
  assert.equal(calc('=AND(A1:A3)', { A1: 'text' }), '#VALUE!');
});

test('math functions use Excel rounding and signed modulo', () => {
  assert.equal(calc('=ROUND(-1.25,1)'), -1.3);
  assert.equal(calc('=ROUND(1.005,2)'), 1.01);
  assert.equal(calc('=ROUNDUP(-1.21,1)'), -1.3);
  assert.equal(calc('=ROUNDUP(0.000000000001,0)'), 1);
  assert.equal(calc('=ROUNDDOWN(-1.29,1)'), -1.2);
  assert.equal(calc('=ROUND(1234,-2)'), 1200);
  assert.equal(calc('=ABS(-3)+SQRT(9)+POWER(2,3)'), 14);
  assert.equal(calc('=MOD(-3,2)'), 1);
  assert.equal(calc('=MOD(3,-2)'), -1);
  assert.equal(calc('=INT(-2.1)'), -3);
  assert.equal(calc('=SQRT(-1)'), '#NUM!');
  assert.equal(calc('=MOD(1,0)'), '#DIV/0!');
});

test('text functions preserve literal strings and validate lengths', () => {
  assert.equal(calc('=LEN("안녕하세요")'), 5);
  assert.equal(calc('=LEFT("abcd",2)&RIGHT("abcd",2)'), 'abcd');
  assert.equal(calc('=RIGHT("abcd",0)'), '');
  assert.equal(calc('=MID("abcdef",2,3)'), 'bcd');
  assert.equal(calc('=MID("abcdef",0,3)'), '#VALUE!');
  assert.equal(calc('=LEFT("abc",-1)'), '#VALUE!');
  assert.equal(calc('=TRIM("  one   two  ")'), 'one two');
  assert.equal(calc('=UPPER("abc")&LOWER("XYZ")'), 'ABCxyz');
  assert.equal(calc('=CONCAT(A1:A3,"!")', { A1: 'a', A2: 'b', A3: 'c' }), 'abc!');
});

test('conditional aggregates handle comparisons, wildcard escaping, and case insensitivity', () => {
  const cells = { A1: 'Apple', A2: 'apple pie', A3: 'pear', A4: 'a*', B1: 10, B2: 20, B3: 30, B4: 40 };
  assert.equal(calc('=COUNTIF(A1:A4,"app*")', cells), 2);
  assert.equal(calc('=COUNTIF(A1:A4,"a~*")', cells), 1);
  assert.equal(calc('=SUMIF(A1:A4,"APP*",B1:B4)', cells), 30);
  assert.equal(calc('=SUMIF(B1:B4,">=20")', cells), 90);
  assert.equal(calc('=COUNTIF(B1:B4,">"&B1)', cells), 3);
  assert.equal(calc('=COUNTIF(A1:A4,">0")', cells), 0);
  assert.equal(calc('=SUMIF(A1:A4,"*",B1:B2)', cells), '#VALUE!');
});

test('lookup functions find exact, approximate, reverse, and fallback matches', () => {
  const cells = { A1: 'alpha', A2: 'beta', A3: 'gamma', B1: 10, B2: 20, B3: 30, C1: 1, C2: 3, C3: 5 };
  assert.equal(calc('=VLOOKUP("BETA",A1:B3,2,FALSE)', cells), 20);
  assert.equal(calc('=VLOOKUP("bet*",A1:B3,2,0)', cells), 20);
  assert.equal(calc('=VLOOKUP(">beta",A1:B3,2,0)', { ...cells, A2: '>beta' }), 20);
  assert.equal(calc('=VLOOKUP("unknown",A1:B3,2,FALSE)', cells), '#N/A');
  assert.equal(calc('=VLOOKUP(4,C1:C3,1)', cells), 3);
  assert.equal(calc('=VLOOKUP("alpha",A1:B3,3,FALSE)', cells), '#REF!');
  assert.equal(calc('=XLOOKUP("beta",A1:A3,B1:B3)', cells), 20);
  assert.equal(calc('=XLOOKUP("beta",A1:A3,B1:B3,1/0)', cells), 20);
  assert.equal(calc('=XLOOKUP("unknown",A1:A3,B1:B3,"missing")', cells), 'missing');
  assert.equal(calc('=XLOOKUP(4,C1:C3,B1:B3,, -1)', cells), 20);
  assert.equal(calc('=XLOOKUP(4,C1:C3,B1:B3,, 1)', cells), 30);
  assert.equal(calc('=XLOOKUP("b*",A1:A3,B1:B3,,2)', cells), 20);
  assert.equal(calc('=XLOOKUP("a",A1:A3,B1:B3,,0,-1)', { A1: 'a', A2: 'b', A3: 'a', B1: 1, B2: 2, B3: 3 }), 3);
});

test('dates follow Excel serials, leap-day compatibility, and normalized month/day', () => {
  assert.equal(calc('=DATE(1900,1,1)'), 1);
  assert.equal(calc('=DATE(1900,2,28)'), 59);
  assert.equal(calc('=DATE(1900,2,29)'), 60);
  assert.equal(calc('=DATE(1900,3,1)'), 61);
  assert.equal(calc('=YEAR(DATE(2026,10,17))'), 2026);
  assert.equal(calc('=MONTH(DATE(2026,13,1))'), 1);
  assert.equal(calc('=DAY(DATE(2026,3,0))'), 28);
  assert.equal(calc('=DAY(60)'), 29);
  assert.equal(calc('=YEAR(-1)'), '#NUM!');
  assert.ok(calc('=TODAY()') > 40000);
});

test('sheet references support spaces, apostrophes, case-insensitive lookup, and cross-sheet cycles', () => {
  const wb = book({ A1: "='Sales Q1'!B2+SUM('Sales Q1'!C1:C2)", A2: "='O''Brien'!A1" });
  wb.sheets.push({ ...createWorkbook().sheets[0], id: 'sales', name: 'Sales Q1', cells: { B2: { raw: '40' }, C1: { raw: '1' }, C2: { raw: '2' } } });
  wb.sheets.push({ ...createWorkbook().sheets[0], id: 'other', name: "O'Brien", cells: { A1: { raw: '=Sheet1!A1' } } });
  assert.equal(evaluateCell(wb, 'sheet-1', 'A1'), 43);
  assert.equal(evaluateCell(wb, 'sheet-1', 'A2'), 43);
  assert.equal(evaluateFormula(wb, 'sheet-1', "='sales q1'!B2"), 40);
  wb.sheets[1].cells.B2.raw = '=Sheet1!A1';
  assert.equal(evaluateCell(wb, 'sheet-1', 'A1'), '#CYCLE!');
  assert.equal(evaluateFormula(wb, 'sheet-1', '=Missing!A1'), '#REF!');
});

test('formula errors and hostile inputs are bounded without running JavaScript', () => {
  assert.equal(calc('=1/0'), '#DIV/0!');
  assert.equal(calc('=A1', { A1: '=A1' }), '#CYCLE!');
  assert.equal(calc('=XFE1'), '#REF!');
  assert.equal(calc('=Missing(1)'), '#NAME?');
  assert.equal(calc('=foo'), '#NAME?');
  assert.equal(calc('=SUM(A1:XFD1048576)'), '#NUM!');
  assert.equal(calc('=1+'), '#VALUE!');
  assert.equal(calc('=globalThis.process.exit()'), '#NAME?');
  assert.equal(calc('=' + '('.repeat(200) + '1' + ')'.repeat(200)), '#VALUE!');
  assert.equal(calc('=10^1000'), '#NUM!');
  assert.equal(calc('=IFERROR(#REF!,42)'), 42);
});

test('formula fill shifts relative references, preserves strings, sheet names, and absolute locks', () => {
  assert.equal(shiftFormula('=A1+$B1+C$2+$D$3', 2, 3), '=D3+$B3+F$2+$D$3');
  assert.equal(shiftFormula('=SUM(A1:B2)+"A1"', 1, 1), '=SUM(B2:C3)+"A1"');
  assert.equal(shiftFormula("='A1'!A1+Sheet1!$B2", 1, 2), "='A1'!C2+Sheet1!$B3");
  assert.equal(shiftFormula('=A1', -1, 0), '=#REF!');
  assert.equal(shiftFormula('=XFD1', 0, 1), '=#REF!');
  assert.equal(shiftFormula('text A1', 1, 1), 'text A1');
});

test('CSV import/export round-trips commas, quotes, embedded CRLF, unicode, and empty trailing fields', () => {
  const rows = [['이름', '값', '메모'], ['alpha, beta', '12', 'line 1\r\nline 2'], ['"quoted"', '', '']];
  assert.deepEqual(parseCSV(toCSV(rows)), rows);
  assert.deepEqual(parseCSV('\uFEFFa,b\r\n1,2\r\n'), [['a', 'b'], ['1', '2']]);
  assert.deepEqual(parseCSV('a,'), [['a', '']]);
  assert.deepEqual(parseCSV(''), []);
  assert.deepEqual(parseCSV('""'), [['']]);
  assert.throws(() => parseCSV('"unclosed'));
  assert.throws(() => parseCSV('"closed"invalid'));
});

test('structural edits update sheet-inherited ranges, absolute locks, and skip functions and strings', () => {
  const change = (formula, opts = {}) => transformFormulaReferences(formula, { axis: 'r', index: 0, delta: 1, targetSheet: 'Other', ownerSheet: 'Sheet1', ...opts });
  assert.equal(change("=SUM('Other'!$A$1:A5)+Sheet1!A1"), "=SUM('Other'!$A$2:A6)+Sheet1!A1");
  assert.equal(change("=SUM('Other'!A1:A5)", { targetSheet: 'Sheet1' }), "=SUM('Other'!A1:A5)");
  assert.equal(change('=SUM(other!A1:A5)', { targetSheet: 'OTHER' }), '=SUM(other!A2:A6)');
  assert.equal(change('=LOG10(A1)+"A1"+SUM($A$1:B2)', { targetSheet: 'Sheet1' }), '=LOG10(A2)+"A1"+SUM($A$2:B3)');
  assert.equal(change("='O''Brien'!A1:B3", { targetSheet: "O'Brien" }), "='O''Brien'!A2:B4");
  assert.equal(change('A1:C3', { targetSheet: 'Sheet1' }), 'A2:C4');
  assert.equal(change('=A1', { targetSheet: 'Sheet1', axis: 'c', index: 0 }), '=B1');
  assert.equal(change('=XFD1', { targetSheet: 'Sheet1', axis: 'c', index: 0 }), '=#REF!');
});

test('structural deletions contract ranges at boundaries and invalidate fully deleted references', () => {
  const remove = (formula, opts = {}) => transformFormulaReferences(formula, { axis: 'r', index: 0, delta: -1, targetSheet: 'Sheet1', ownerSheet: 'Sheet1', ...opts });
  assert.equal(remove('=SUM(A1:A5)'), '=SUM(A1:A4)');
  assert.equal(remove('=SUM(A1:A5)', { index: 4 }), '=SUM(A1:A4)');
  assert.equal(remove('=SUM(A1:A5)', { index: 2 }), '=SUM(A1:A4)');
  assert.equal(remove('=SUM(A2:A5)'), '=SUM(A1:A4)');
  assert.equal(remove('=SUM(A5:A1)'), '=SUM(A4:A1)');
  assert.equal(remove('=SUM(A1:B1)'), '=SUM(#REF!)');
  assert.equal(remove('=$A$1+A2'), '=#REF!+A1');
  assert.equal(remove('=SUM(B1:D4)', { axis: 'c', index: 1 }), '=SUM(B1:C4)');
  assert.equal(remove('=SUM(B1:B4)', { axis: 'c', index: 1 }), '=SUM(#REF!)');
  assert.equal(remove("=SUM('Other'!A1:A5)", { targetSheet: 'Other' }), "=SUM('Other'!A1:A4)");
  assert.equal(remove('=SUM(A1:A5)', { index: 5 }), '=SUM(A1:A5)');
});

test('workbook persistence validates untrusted models and preserves supported content', () => {
  const wb = book({ A1: '=SUM(1,2)', B1: '안녕하세요' });
  wb.sheets[0].cells.A1.style = { bold: true, color: '#107c41', format: 'currency' };
  wb.sheets[0].charts.push({ type: 'bar', range: 'A1:B5' });
  assert.deepEqual(parseWorkbook(serializeWorkbook(wb)), wb);
  assert.throws(() => parseWorkbook('{"__proto__":{"polluted":true}}'));
  assert.throws(() => parseWorkbook(JSON.stringify({ ...wb, activeSheetId: 'missing' })));
  assert.throws(() => parseWorkbook(JSON.stringify({ ...wb, sheets: [] })));
  assert.throws(() => serializeWorkbook(book({ A0: 'bad' })));
  const duplicate = structuredClone(wb); duplicate.sheets.push(structuredClone(duplicate.sheets[0]));
  assert.throws(() => serializeWorkbook(duplicate));
  const invalidStyle = structuredClone(wb); invalidStyle.sheets[0].cells.A1.style.format = { invalid: true };
  assert.throws(() => serializeWorkbook(invalidStyle));
  const injectedColor = structuredClone(wb); injectedColor.sheets[0].cells.A1.style.color = 'red" onmouseover="alert(1)';
  assert.throws(() => serializeWorkbook(injectedColor));
  const invalidMerge = structuredClone(wb); invalidMerge.sheets[0].merges = ['A1'];
  assert.throws(() => serializeWorkbook(invalidMerge));
  const invalidChart = structuredClone(wb); invalidChart.sheets[0].charts[0].range = 'oops';
  assert.throws(() => serializeWorkbook(invalidChart));
});

test('display formatting keeps errors and supports general, number, percent, currency, and date', () => {
  assert.equal(formatValue(0.1 + 0.2), '0.3');
  assert.equal(formatValue(true), 'TRUE');
  assert.equal(formatValue('#DIV/0!', { format: 'currency' }), '#DIV/0!');
  assert.equal(formatValue(1234.5, { format: 'number' }), '1,234.50');
  assert.equal(formatValue(0.125, { format: 'percent' }), '12.50%');
  assert.equal(formatValue(1234.5, { format: 'currency', decimals: 0 }), '₩1,235');
  assert.equal(formatValue(0.125, { format: '0.0%' }), '12.5%');
  assert.equal(formatValue(61, { format: 'date' }), '1900-03-01');
  assert.equal(formatValue(60, { format: 'yyyy/mm/dd' }), '1900/02/29');
});

test('imported numeric and date formats display correctly and category changes supersede old custom formats', () => {
  assert.equal(formatValue(1.23456, { format: 'number', numberFormat: '0.000' }), '1.235');
  assert.equal(formatValue(1.2, { format: 'number', numberFormat: '0.##' }), '1.2');
  assert.equal(formatValue(0.125, { format: 'percent', numberFormat: '0%' }), '13%');
  assert.equal(formatValue(0.125, { format: 'number', numberFormat: '0%' }), '0.13');
  assert.equal(formatValue(1234.5, { format: 'currency', numberFormat: '$#,##0.00' }), '$1,234.50');
  assert.equal(formatValue(61, { format: 'date', numberFormat: 'm/d/yy' }), '3/1/00');
  assert.equal(formatValue(61.5625, { format: 'date', numberFormat: 'yyyy-mm-dd hh:mm:ss' }), '1900-03-01 13:30:00');
  assert.equal(formatValue(61.5625, { format: 'date', numberFormat: 'h:mm AM/PM' }), '1:30 PM');
  assert.equal(formatValue(1.5, { format: 'date', numberFormat: '[h]:mm' }), '36:00');
});
