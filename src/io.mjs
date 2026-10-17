import ExcelJS from 'exceljs/dist/exceljs.min.js';
import { createEvaluator } from './engine.mjs';
import { attachNativeSource } from './native-plan.mjs';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const MAX_CELLS = 200_000;
const DAY = 86_400_000;
const formats = {
  general: 'General', number: '#,##0.00', currency: '₩#,##0.00',
  percent: '0.00%', date: 'yyyy-mm-dd', accounting: '₩* #,##0.00;[Red](₩* #,##0.00);₩* "-";@',
};
const errors = new Set(['#N/A', '#REF!', '#VALUE!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!']);
const numeric = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/;

const uniqueId = () => globalThis.crypto?.randomUUID?.() || `sheet-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
function newSheet(name = 'Sheet1') {
  return { id: uniqueId(), name, cells: {}, columnWidths: {}, rowHeights: {}, merges: [], charts: [] };
}
function address(col, row) {
  let letters = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + (n - 1) % 26) + letters;
  return letters + (row + 1);
}
function coordinates(value) {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(value);
  if (!match) return null;
  const col = [...match[1].toUpperCase()].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
  const row = Number(match[2]) - 1;
  return col < 16384 && row < 1048576 ? { col, row } : null;
}
function bytes(value) {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  return new Uint8Array(value);
}
function fileName(value) { return String(value || '통합 문서').replace(/\.(xlsx|xlsm|csv|tsv|sheet\.json|json)$/i, ''); }
function stringRaw(value) {
  const text = String(value);
  // Preserve XLSX text types, including numeric-looking identifiers and '=…'.
  return /^[=']/.test(text) || numeric.test(text.trim()) || /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$/.test(text.trim()) ||
    /^(TRUE|FALSE)$/i.test(text) || errors.has(text) || /^#(?:CYCLE!|SPILL!)$/.test(text) ? "'" + text : text;
}
function rgb(color) {
  if (color?.argb && /^[0-9a-f]{6,8}$/i.test(color.argb)) return '#' + color.argb.slice(-6);
  // Default Office theme colors. Theme-specific palette customizations are not retained.
  const theme = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
  return Number.isInteger(color?.theme) ? theme[color.theme] : undefined;
}
function argb(color) {
  if (typeof color !== 'string') return undefined;
  let hex = color.replace(/^#/, '');
  if (/^[0-9a-f]{3}$/i.test(hex)) hex = [...hex].map(char => char + char).join('');
  return /^[0-9a-f]{6}$/i.test(hex) ? 'FF' + hex.toUpperCase() : undefined;
}
function formatKind(format) {
  if (!format || /^general$/i.test(format)) return 'general';
  const cleaned = format.replace(/"[^"]*"|\\./g, '');
  if (/[ymdhis]/i.test(cleaned.replace(/\[[^\]]*\]/g, ''))) return 'date';
  if (cleaned.includes('%')) return 'percent';
  if (cleaned.includes('*')) return 'accounting';
  if (/[$₩€£¥]|\[\$/.test(cleaned)) return 'currency';
  return 'number';
}
function readStyle(cell) {
  const result = {};
  const { font = {}, alignment = {}, fill = {}, numFmt } = cell;
  if (font.bold) result.bold = true;
  if (font.italic) result.italic = true;
  if (font.underline) result.underline = true;
  if (font.size) result.fontSize = font.size;
  if (font.name) result.fontFamily = font.name;
  if (rgb(font.color)) result.color = rgb(font.color);
  if (fill.type === 'pattern' && fill.pattern === 'solid' && rgb(fill.fgColor)) result.background = rgb(fill.fgColor);
  if (alignment.horizontal) result.align = alignment.horizontal;
  if (alignment.vertical) result.verticalAlign = alignment.vertical;
  if (alignment.wrapText) result.wrap = true;
  if (Object.values(cell.border || {}).some(side => side?.style)) result.border = true;
  if (numFmt) {
    result.format = formatKind(numFmt);
    result.numberFormat = numFmt;
  }
  return result;
}

function decodeText(data) {
  if (data[0] === 0xff && data[1] === 0xfe) return new TextDecoder('utf-16le').decode(data).replace(/^\uFEFF/, '');
  if (data[0] === 0xfe && data[1] === 0xff) return new TextDecoder('utf-16be').decode(data).replace(/^\uFEFF/, '');
  return new TextDecoder('utf-8').decode(data).replace(/^\uFEFF/, '');
}
function parseDelimited(text, delimiter) {
  const rows = [];
  let row = [], field = '', quoted = false, cellCount = 0;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (char === '"') {
      if (quoted && text[i + 1] === '"') { field += '"'; i++; }
      else if (quoted || field === '') quoted = !quoted;
      else field += char;
    } else if (!quoted && (char === delimiter || char === '\n' || char === '\r')) {
      row.push(field); field = '';
      if (++cellCount > MAX_CELLS) throw new Error('200,000개 이하의 셀이 있는 파일을 선택하세요.');
      if (char !== delimiter) {
        rows.push(row); row = [];
        if (char === '\r' && text[i + 1] === '\n') i++;
      }
    } else field += char;
  }
  if (quoted) throw new Error('CSV의 따옴표가 닫히지 않았습니다.');
  if (field !== '' || row.length) {
    if (++cellCount > MAX_CELLS) throw new Error('200,000개 이하의 셀이 있는 파일을 선택하세요.');
    row.push(field); rows.push(row);
  }
  return rows;
}
function readDelimited(data, filename) {
  const text = decodeText(data);
  const delimiter = /\.tsv$/i.test(filename) ? '\t' : ',';
  const sheet = newSheet(fileName(filename).slice(0, 31) || 'Sheet1');
  const rows = parseDelimited(text, delimiter);
  rows.forEach((row, r) => row.forEach((value, c) => {
    if (!value) return;
    if (c >= 16384 || r >= 1048576) throw new Error('Excel의 최대 행 또는 열 수를 초과했습니다.');
    // CSV has no type metadata. Keep leading-zero IDs as text, ordinary numbers as numbers.
    const raw = /^[=+@]/.test(value) || /^0\d+/.test(value) || value.startsWith("'") ? "'" + value : value;
    sheet.cells[address(c, r)] = { raw, style: {} };
  }));
  return { name: fileName(filename), activeSheetId: sheet.id, sheets: [sheet] };
}

/** Read XLSX/CSV/TSV bytes into the app's sparse workbook model. */
export async function readSpreadsheet(arrayBuffer, filename = '통합 문서.xlsx') {
  const data = bytes(arrayBuffer);
  if (data.byteLength > 100 * 1024 * 1024) throw new Error('100 MB 이하의 파일을 선택하세요.');
  if (/\.(csv|tsv)$/i.test(filename)) return readDelimited(data, filename);
  if (!/\.xls[xm]$/i.test(filename)) throw new Error('XLSX, XLSM, CSV 또는 TSV 파일을 선택하세요.');
  const book = new ExcelJS.Workbook();
  try { await book.xlsx.load(data); }
  catch (error) { throw new Error('XLSX 파일을 읽을 수 없습니다. 암호화되었거나 손상된 파일인지 확인하세요.', { cause: error }); }
  let cellCount = 0;
  const sheets = book.worksheets.map(source => {
    const sheet = newSheet(source.name);
    source.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (row.height) sheet.rowHeights[rowNumber - 1] = row.height * 4 / 3;
      row.eachCell({ includeEmpty: false }, cell => {
        if (cell.isMerged && cell.master.address !== cell.address) return;
        if (++cellCount > MAX_CELLS) throw new Error('200,000개 이하의 셀이 있는 파일을 선택하세요.');
        const style = readStyle(cell);
        let raw = '';
        const value = cell.value;
        if (cell.formula) raw = '=' + cell.formula;
        else if (value instanceof Date) { raw = String(value.getTime() / DAY + 25569); style.format = 'date'; }
        else if (typeof value === 'number') raw = String(value);
        else if (typeof value === 'boolean') raw = String(value).toUpperCase();
        else if (typeof value === 'string') raw = stringRaw(value);
        else if (value?.richText) raw = stringRaw(value.richText.map(run => run.text).join(''));
        else if (value?.text !== undefined) raw = stringRaw(value.text);
        else if (value?.error) raw = value.error;
        const entry = { raw, style };
        if (typeof cell.note === 'string') entry.note = cell.note;
        else if (cell.note?.texts) entry.note = cell.note.texts.map(run => run.text).join('');
        if (cell.formula && cell.result !== undefined) entry.cachedValue = cell.result?.error || cell.result;
        sheet.cells[cell.address] = entry;
      });
    });
    // Rows with only a custom height can be omitted by eachRow(includeEmpty:false).
    for (const row of source.model.rows || []) if (row.height) sheet.rowHeights[row.number - 1] = row.height * 4 / 3;
    for (const column of source.columns || []) if (column.width) sheet.columnWidths[column.number - 1] = Math.round(column.width * 7 + 5);
    sheet.merges = [...(source.model.merges || [])];
    return sheet;
  });
  if (!sheets.length) sheets.push(newSheet());
  const activeIndex = book.views?.[0]?.activeTab || 0;
  const result = { name: fileName(filename), activeSheetId: sheets[activeIndex]?.id || sheets[0].id, sheets };
  return /\.xlsm$/i.test(filename) ? attachNativeSource(result, data, filename) : result;
}

async function evaluator() {
  const caches = new Map();
  return (workbook, sheetId, addr) => {
    if (!caches.has(sheetId)) caches.set(sheetId, createEvaluator(workbook, sheetId));
    return caches.get(sheetId)(addr);
  };
}
function toScalar(raw) {
  if (typeof raw !== 'string') return raw ?? null;
  if (raw.startsWith("'")) return raw.slice(1);
  if (raw === '') return null;
  if (numeric.test(raw.trim()) && Number.isFinite(Number(raw))) return Number(raw);
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$/.test(raw.trim())) return Number(raw.trim().slice(0, -1)) / 100;
  if (/^(TRUE|FALSE)$/i.test(raw)) return raw.toUpperCase() === 'TRUE';
  if (errors.has(raw)) return { error: raw };
  return raw;
}
function applyStyle(target, style = {}) {
  target.font = {
    name: style.fontFamily || 'Calibri', size: Number(style.fontSize) || 11,
    bold: !!style.bold, italic: !!style.italic, underline: !!style.underline,
    ...(argb(style.color) ? { color: { argb: argb(style.color) } } : {}),
  };
  if (argb(style.background)) target.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(style.background) } };
  target.alignment = {
    horizontal: ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'].includes(style.align) ? style.align : undefined,
    vertical: style.verticalAlign || 'middle', wrapText: !!style.wrap,
  };
  if (style.border) {
    const edge = { style: 'thin', color: { argb: 'FF8A8A8A' } };
    target.border = { top: edge, right: edge, bottom: edge, left: edge };
  }
  target.numFmt = style.numberFormat && formatKind(style.numberFormat) === (style.format || 'general')
    ? style.numberFormat : formats[style.format] || 'General';
  if (Number.isInteger(style.decimals) && ['number', 'currency', 'accounting', 'percent'].includes(style.format)) {
    const digits = Math.min(12, Math.max(0, style.decimals));
    const fraction = digits ? '.' + '0'.repeat(digits) : '';
    target.numFmt = style.format === 'percent' ? '0' + fraction + '%' :
      (style.format === 'currency' || style.format === 'accounting' ? '₩' : '') + '#,##0' + fraction;
  }
}
function safeSheetName(name, used) {
  const base = String(name || 'Sheet').replace(/[\\/*?:\[\]]/g, '_').replace(/^'|'$/g, '').slice(0, 31) || 'Sheet';
  let value = base;
  for (let suffix = 2; used.has(value.toLowerCase()); suffix++) value = base.slice(0, 31 - String(suffix).length - 1) + ' ' + suffix;
  used.add(value.toLowerCase());
  return value;
}
function cachedResult(evaluate, workbook, sheet, addr, cell) {
  let value;
  try { value = evaluate?.(workbook, sheet.id, addr); } catch { /* Excel can recalculate an unsupported formula. */ }
  if (value === undefined || value === '#NAME?' || value === '#ERROR!') value = cell.cachedValue;
  if (value === undefined || value === null) return undefined;
  if (typeof value === 'number') return Number.isFinite(value) ? value : { error: '#NUM!' };
  if (typeof value === 'boolean') return value;
  if (errors.has(value)) return { error: value };
  return String(value);
}

/** Export XLSX (all sheets) or CSV/TSV (the active sheet). */
export async function writeSpreadsheet(workbook, format = 'xlsx', options = {}) {
  const kind = String(format).toLowerCase().replace(/^\./, '');
  if (!workbook?.sheets?.length) throw new Error('통합 문서에 시트가 없습니다.');
  if (workbook.nativeSource && !options.allowMacroLoss) throw new Error('VBA 보존을 위해 XLSM 저장을 사용하세요. 일반 형식 변환에는 매크로 제외 확인이 필요합니다.');
  if (!['xlsx', 'csv', 'tsv'].includes(kind)) throw new Error('지원되지 않는 저장 형식입니다.');
  const evaluate = await evaluator();
  if (kind === 'csv' || kind === 'tsv') {
    const sheet = workbook.sheets.find(item => item.id === workbook.activeSheetId) || workbook.sheets[0];
    let maxRow = -1, maxCol = -1;
    for (const [addr, cell] of Object.entries(sheet.cells)) {
      if (cell.raw === '' || cell.raw === undefined) continue;
      const coordinate = coordinates(addr);
      if (!coordinate) continue;
      maxRow = Math.max(maxRow, coordinate.row); maxCol = Math.max(maxCol, coordinate.col);
    }
    if ((maxRow + 1) * (maxCol + 1) > 2_000_000) throw new Error('CSV 범위가 너무 큽니다. XLSX로 저장하세요.');
    const delimiter = kind === 'tsv' ? '\t' : ',';
    const rows = [];
    for (let r = 0; r <= maxRow; r++) {
      const row = [];
      for (let c = 0; c <= maxCol; c++) {
        const addr = address(c, r), cell = sheet.cells[addr];
        let value = cell?.raw?.startsWith?.('=') ? cachedResult(evaluate, workbook, sheet, addr, cell) : toScalar(cell?.raw);
        if (value && typeof value === 'object') value = value.error || '';
        let text = value === undefined || value === null ? '' : String(value);
        // Prevent plain text beginning with a formula marker from becoming a formula when opened elsewhere.
        if (typeof value === 'string' && /^[=+@\-\t\r]/.test(text)) text = "'" + text;
        if (text.includes(delimiter) || /["\r\n]/.test(text)) text = '"' + text.replace(/"/g, '""') + '"';
        row.push(text);
      }
      rows.push(row.join(delimiter));
    }
    return { data: new TextEncoder().encode('\uFEFF' + rows.join('\r\n')), mime: kind === 'csv' ? 'text/csv;charset=utf-8' : 'text/tab-separated-values;charset=utf-8', extension: kind };
  }
  const book = new ExcelJS.Workbook();
  book.creator = 'Sheet Studio';
  book.created = new Date(); book.modified = new Date();
  book.calcProperties.fullCalcOnLoad = true;
  book.views = [{ activeTab: Math.max(0, workbook.sheets.findIndex(sheet => sheet.id === workbook.activeSheetId)), visibility: 'visible' }];
  const names = new Set();
  for (const sheet of workbook.sheets) {
    const target = book.addWorksheet(safeSheetName(sheet.name, names));
    for (const [addr, cell] of Object.entries(sheet.cells || {})) {
      if (!coordinates(addr)) continue;
      const targetCell = target.getCell(addr);
      if (typeof cell.raw === 'string' && cell.raw.startsWith('=')) {
        const result = cachedResult(evaluate, workbook, sheet, addr, cell);
        targetCell.value = { formula: cell.raw.slice(1), ...(result === undefined ? {} : { result }) };
      } else targetCell.value = toScalar(cell.raw);
      applyStyle(targetCell, cell.style);
      if (typeof cell.raw === 'string' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$/.test(cell.raw.trim()) && (!cell.style?.format || cell.style.format === 'general')) targetCell.numFmt = '0.00%';
      if (typeof cell.note === 'string') targetCell.note = cell.note;
    }
    for (const [index, width] of Object.entries(sheet.columnWidths || {})) {
      if (/^\d+$/.test(index) && Number(index) < 16384 && Number.isFinite(Number(width))) target.getColumn(Number(index) + 1).width = Math.max(1, (Number(width) - 5) / 7);
    }
    for (const [index, height] of Object.entries(sheet.rowHeights || {})) {
      if (/^\d+$/.test(index) && Number(index) < 1048576 && Number.isFinite(Number(height))) target.getRow(Number(index) + 1).height = Math.max(1, Number(height) * 3 / 4);
    }
    for (const merge of sheet.merges || []) {
      if (typeof merge !== 'string' || !/^[A-Z]+\d+:[A-Z]+\d+$/i.test(merge)) continue;
      target.mergeCells(merge);
    }
  }
  const data = new Uint8Array(await book.xlsx.writeBuffer());
  return { data, mime: XLSX_MIME, extension: 'xlsx' };
}
