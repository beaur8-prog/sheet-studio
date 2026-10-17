import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { readSpreadsheet, writeSpreadsheet } from '../src/io.mjs';

function sample() {
  return {
    name: '분기 보고서', activeSheetId: 'summary',
    sheets: [{
      id: 'summary', name: '요약',
      cells: {
        A1: { raw: '매출', style: { bold: true, italic: true, underline: true, fontFamily: 'Arial', fontSize: 14, color: '#163a24', background: '#e4f3e9', align: 'center' } },
        A2: { raw: '12.5', style: { format: 'currency' } },
        B2: { raw: '7.5', style: { format: 'number' } },
        C2: { raw: '=SUM(A2:B2)', cachedValue: 20, style: { format: 'number' } },
        A3: { raw: "'00123", style: {} },
        B3: { raw: "'=literal", style: {} },
        C3: { raw: 'TRUE', style: {} },
        A4: { raw: '46022', style: { format: 'date' } },
        B4: { raw: '0.15', style: { format: 'percent' } },
        C4: { raw: '=1/0', cachedValue: '#DIV/0!', style: {} },
      },
      columnWidths: { 0: 180, 1: 95 }, rowHeights: { 0: 36, 7: 42 }, merges: ['A6:C6'], charts: [],
    }, {
      id: 'detail', name: '원본', cells: { A1: { raw: '한글, "따옴표"\n둘째 줄', style: {} } },
      columnWidths: {}, rowHeights: {}, merges: [], charts: [],
    }],
  };
}

test('XLSX roundtrip retains sheets, formulas, cached results, values, styles and dimensions', async () => {
  const output = await writeSpreadsheet(sample());
  assert.equal(output.extension, 'xlsx');
  assert.ok(output.data instanceof Uint8Array);
  const check = new ExcelJS.Workbook();
  await check.xlsx.load(output.data);
  assert.equal(check.getWorksheet('요약').getCell('C2').formula, 'SUM(A2:B2)');
  assert.equal(check.getWorksheet('요약').getCell('C2').result, 20);
  assert.equal(check.getWorksheet('요약').getCell('A3').value, '00123');
  assert.equal(check.getWorksheet('요약').getCell('B3').value, '=literal');
  assert.equal(check.getWorksheet('요약').getCell('C3').value, true);
  assert.deepEqual(check.getWorksheet('요약').getCell('C4').result, { error: '#DIV/0!' });
  const result = await readSpreadsheet(output.data, '분기 보고서.xlsx');
  assert.equal(result.name, '분기 보고서');
  assert.deepEqual(result.sheets.map(sheet => sheet.name), ['요약', '원본']);
  const sheet = result.sheets[0];
  assert.equal(sheet.cells.C2.raw, '=SUM(A2:B2)');
  assert.equal(sheet.cells.A3.raw, "'00123");
  assert.equal(sheet.cells.B3.raw, "'=literal");
  assert.equal(sheet.cells.A4.raw, '46022');
  assert.equal(sheet.cells.A4.style.format, 'date');
  assert.equal(sheet.cells.B4.style.format, 'percent');
  assert.equal(sheet.cells.A1.style.color, '#163A24');
  assert.equal(sheet.cells.A1.style.background, '#E4F3E9');
  assert.equal(sheet.cells.A1.style.bold, true);
  assert.equal(sheet.cells.A1.style.fontSize, 14);
  assert.equal(sheet.columnWidths[0], 180);
  assert.equal(sheet.rowHeights[0], 36);
  assert.equal(sheet.rowHeights[7], 42);
  assert.deepEqual(sheet.merges, ['A6:C6']);
});

test('CSV handles BOM, quotes, embedded line breaks, trailing empty cells, IDs and formula-like text', async () => {
  const data = new TextEncoder().encode('\uFEFF이름,메모,금액,코드\r\n"가,나","그가 ""안녕""\n이라고",12.5,00123\r\n=HYPERLINK("bad"),,,\r\n');
  const result = await readSpreadsheet(data, '한국어.csv');
  const sheet = result.sheets[0];
  assert.equal(sheet.cells.A2.raw, '가,나');
  assert.equal(sheet.cells.B2.raw, '그가 "안녕"\n이라고');
  assert.equal(sheet.cells.C2.raw, '12.5');
  assert.equal(sheet.cells.D2.raw, "'00123");
  assert.equal(sheet.cells.A3.raw, "'=HYPERLINK(\"bad\")");
  const output = await writeSpreadsheet(result, 'csv');
  const text = new TextDecoder().decode(output.data);
  assert.ok(text.includes('"가,나","그가 ""안녕""\n이라고",12.5,00123'));
});

test('CSV export uses active sheet and evaluated formula values', async () => {
  const book = sample();
  let output = await writeSpreadsheet(book, 'csv');
  assert.ok(new TextDecoder().decode(output.data).includes('12.5,7.5,20'));
  book.activeSheetId = 'detail';
  output = await writeSpreadsheet(book, 'csv');
  assert.equal(new TextDecoder().decode(output.data), '"한글, ""따옴표""\n둘째 줄"');
});

test('XLSX imports real date cells and translated shared formulas', async () => {
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet('Dates');
  sheet.getCell('A1').value = new Date('2026-01-01T00:00:00Z');
  sheet.getCell('A1').numFmt = 'yyyy-mm-dd';
  sheet.getCell('A2').value = 2;
  sheet.getCell('B1').value = { formula: 'A1+1', result: 46024 };
  sheet.getCell('B2').value = { sharedFormula: 'B1', result: 3 };
  const result = await readSpreadsheet(await source.xlsx.writeBuffer(), 'dates.xlsx');
  assert.equal(result.sheets[0].cells.A1.style.format, 'date');
  assert.equal(result.sheets[0].cells.A1.raw, String(Date.parse('2026-01-01T00:00:00Z') / 86400000 + 25569));
  assert.equal(result.sheets[0].cells.B2.raw, '=A2+1');
});

test('XLSX preserves percentage-looking and error-looking text while exporting typed percentages numerically', async () => {
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet('Text');
  sheet.getCell('A1').value = '12%';
  sheet.getCell('A2').value = '#N/A';
  sheet.getCell('A3').value = ' 12 ';
  const book = await readSpreadsheet(await source.xlsx.writeBuffer(), 'types.xlsx');
  assert.equal(book.sheets[0].cells.A1.raw, "'12%");
  assert.equal(book.sheets[0].cells.A2.raw, "'#N/A");
  assert.equal(book.sheets[0].cells.A3.raw, "' 12 ");
  book.sheets[0].cells.B1 = { raw: '25%', style: { format: 'percent', decimals: 3, border: true }, note: '설명' };
  const output = await writeSpreadsheet(book);
  const verified = new ExcelJS.Workbook();
  await verified.xlsx.load(output.data);
  assert.equal(verified.worksheets[0].getCell('A1').value, '12%');
  assert.equal(verified.worksheets[0].getCell('A2').value, '#N/A');
  assert.equal(verified.worksheets[0].getCell('B1').value, 0.25);
  assert.equal(verified.worksheets[0].getCell('B1').numFmt, '0.000%');
  assert.equal(verified.worksheets[0].getCell('B1').border.top.style, 'thin');
  const reread = await readSpreadsheet(output.data, 'again.xlsx');
  assert.equal(reread.sheets[0].cells.B1.note, '설명');
});

test('invalid format and broken input give clear errors', async () => {
  await assert.rejects(readSpreadsheet(new Uint8Array(), 'old.xls'), /XLSX/);
  await assert.rejects(readSpreadsheet(new Uint8Array(), 'broken.xlsx'), /XLSX/);
  await assert.rejects(readSpreadsheet(new TextEncoder().encode('"unclosed'), 'broken.csv'), /따옴표/);
  await assert.rejects(writeSpreadsheet(sample(), 'pdf'), /지원되지/);
});
