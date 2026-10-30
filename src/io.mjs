import {plainCopy,prepareWorksheetFeatures,addPreservedFeatures} from './xlsx-preservation.mjs';
import {cloneData} from './clone.mjs';
import {MAX_DOCUMENT_CELLS} from './limits.mjs';
import {cellBorders,borderSides,borderStyles} from './borders.mjs';
import {restoreWorksheetParts,addWorksheetParts} from './worksheet-xlsx.mjs';
import {importConditional,exportConditional,shiftWorksheetFeatures} from './conditional.mjs';
import {addFilterParts,restoreFilterParts} from './filter-xlsx.mjs';
import JSZip from 'jszip';
import {createColorResolver} from './spreadsheet-colors.mjs';
import ExcelJS from 'exceljs/dist/exceljs.min.js';
import { createEvaluator,transformFormulaReferences,formulaForXlsx,formulaForEditing } from './engine.mjs';
import { attachNativeSource } from './native-plan.mjs';
import {addPivotParts,restorePivotParts} from './pivot-xlsx.mjs';
import {applyPivot,sourceData} from './pivot.mjs';
import {restoreChartParts,addChartParts} from './chart-xlsx.mjs';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const MAX_CELLS = MAX_DOCUMENT_CELLS;
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
function readStyle(cell,rgb) {
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
  if (borderSides.some(side => cell.border?.[side]?.style)) {
    result.borders=Object.fromEntries(borderSides.map(side=>{const edge=cell.border[side];return [side,edge?.style?{style:borderStyles.includes(edge.style)?edge.style:'thin',color:rgb(edge.color)||'#8a8a8a'}:null];}));
  }
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
      if (++cellCount > MAX_CELLS) throw new Error('1,000,000개 이하의 셀이 있는 파일을 선택하세요.');
      if (char !== delimiter) {
        rows.push(row); row = [];
        if (char === '\r' && text[i + 1] === '\n') i++;
      }
    } else field += char;
  }
  if (quoted) throw new Error('CSV의 따옴표가 닫히지 않았습니다.');
  if (field !== '' || row.length) {
    if (++cellCount > MAX_CELLS) throw new Error('1,000,000개 이하의 셀이 있는 파일을 선택하세요.');
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
    const raw = value.startsWith("'") ? value : /^[=+@\t\r]/.test(value)||/^-/.test(value)&&!numeric.test(value)||/^0\d+/.test(value) ? "'"+value : value;
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
  let colorZip;
  try{colorZip=await JSZip.loadAsync(data);}catch(error){throw new Error('XLSX 파일을 읽을 수 없습니다. 암호화되었거나 손상된 파일인지 확인하세요.',{cause:error});}
  let expanded=0;
  // JSZip 3.10.2 exposes the central-directory size on its pinned compressed
  // entry. Check before ExcelJS inflates cell XML into its in-memory model.
  for(const file of Object.values(colorZip.files))if(!file.dir){const size=file._data?.uncompressedSize;if(!Number.isSafeInteger(size)||size<0)throw Error('XLSX 압축 정보를 읽을 수 없습니다.');expanded+=size;if(expanded>512*1024*1024)throw Error('압축 해제 후 512 MB 이하인 XLSX 파일을 선택하세요.');}
  const preserved=await prepareWorksheetFeatures(colorZip);
  const styleModels=[],styleIndexes=new Map(),styleObjects=new WeakMap();
  const rememberStyle=(cell,baseline)=>{if(!Object.keys(cell.style||{}).length)return undefined;let index=styleObjects.get(cell.style);if(index!==undefined)return index;const key=JSON.stringify(cell.style);index=styleIndexes.get(key);if(index===undefined){index=styleModels.length;styleIndexes.set(key,index);styleModels.push({style:plainCopy(cell.style),baseline:plainCopy(baseline)});}styleObjects.set(cell.style,index);return index;};
  const book = new ExcelJS.Workbook();
  try { await book.xlsx.load(preserved.data||data); }
  catch (error) { throw new Error('XLSX 파일을 읽을 수 없습니다. 암호화되었거나 손상된 파일인지 확인하세요.', { cause: error }); }
  const resolveColor=createColorResolver(await colorZip.file('xl/theme/theme1.xml')?.async('string'),await colorZip.file('xl/styles.xml')?.async('string'));
  const dateOffset=book.properties.date1904?1462:0;
  let cellCount = 0;
  const sheets = book.worksheets.map(source => {
    const sheet = newSheet(source.name);
    const defaultHeight=source.properties?.defaultRowHeight*4/3;
    if(Number.isFinite(defaultHeight)&&defaultHeight>=1&&defaultHeight<=10000)sheet.defaultRowHeight=defaultHeight;
    const importCell = cell => {
      if (cell.isMerged && cell.master.address !== cell.address) return;
      if (++cellCount > MAX_CELLS) throw new Error('1,000,000개 이하의 셀이 있는 파일을 선택하세요.');
      const style = readStyle(cell,resolveColor);
      let raw = '';
      const value = cell.value;
      if (cell.formula) raw = '='+formulaForEditing(cell.formula);
      else if (value instanceof Date) { raw = String(Math.round((value.getTime()/DAY+25569-(value.getTime()<Date.UTC(1900,2,1)?1:0)-dateOffset)*1e8)/1e8); style.format = 'date'; }
      else if (typeof value === 'number') raw = String(value);
      else if (typeof value === 'boolean') raw = String(value).toUpperCase();
      else if (typeof value === 'string') raw = stringRaw(value);
      else if (value?.richText) raw = stringRaw(value.richText.map(run => run.text).join(''));
      else if (value?.text !== undefined) raw = stringRaw(value.text);
      else if (value?.error) raw = value.error;
      const entry = { raw, style };
      const originalStyle=rememberStyle(cell,style);if(originalStyle!==undefined)entry.xlsxStyle=originalStyle;
      if(value?.richText||value?.hyperlink)entry.xlsxValue={raw,value:plainCopy(value)};
      if(cell.formula&&(value?.shareType==='array'||/_xlfn\.|_xlws\./i.test(cell.formula)))entry.xlsxFormula={raw,formula:cell.formula,...(value?.shareType==='array'?{shareType:'array',ref:value.ref}:{})};
      if (typeof cell.note === 'string') entry.note = cell.note;
      else if (cell.note?.texts) entry.note = cell.note.texts.map(run => run.text).join('');
      if (cell.formula && cell.result !== undefined) { entry.cachedValue = cell.result instanceof Date ? Math.round((cell.result.getTime()/DAY+25569-dateOffset)*1e8)/1e8 : cell.result?.error || cell.result; entry.cachedFormula = raw; }
      sheet.cells[cell.address] = entry;
    };
    source.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (row.height) sheet.rowHeights[rowNumber - 1] = row.height * 4 / 3;
      row.eachCell({ includeEmpty: false }, importCell);
    });
    // ExcelJS skips null-valued styled cells (and rows containing only those
    // cells) in eachRow/eachCell. Walk the sparse model to retain empty borders
    // and fills without expanding a sheet's potentially million-row extent.
    const sparseRows=source.model.rows||[];
    for (const row of sparseRows) for (const model of row.cells || []) {
      if (!sheet.cells[model.address] && Object.keys(model.style || {}).length) importCell(source.getCell(model.address));
    }
    // Rows with only a custom height can be omitted by eachRow(includeEmpty:false).
    for (const row of sparseRows) {
      const r=row.number-1;
      if(row.height)sheet.rowHeights[r]=row.height*4/3;
      if(row.hidden){sheet.hiddenSizes||={rows:{},cols:{}};sheet.hiddenSizes.rows[r]=sheet.rowHeights[r]||sheet.defaultRowHeight||25;sheet.rowHeights[r]=0;}
    }
    const protectionDefaults={rows:{},columns:{}};for(const row of sparseRows)if(row.style?.protection)protectionDefaults.rows[row.number-1]=plainCopy(row.style.protection);for(const col of source.columns||[])if(col.style?.protection)protectionDefaults.columns[col.number-1]=plainCopy(col.style.protection);if(Object.keys(protectionDefaults.rows).length||Object.keys(protectionDefaults.columns).length)sheet.xlsxProtectionDefaults=protectionDefaults;
    for (const column of source.columns || []) if (column.width) sheet.columnWidths[column.number - 1] = Math.round(column.width * 7 + 5);
    for(const column of source.columns||[])if(column.hidden){sheet.hiddenSizes||={rows:{},cols:{}};sheet.hiddenSizes.cols[column.number-1]=sheet.columnWidths[column.number-1]||90;sheet.columnWidths[column.number-1]=0;}
    const view=source.views?.find(v=>v.state==='frozen');if(view)sheet.freeze={rows:view.ySplit||0,cols:view.xSplit||0};
    importConditional(source,sheet,resolveColor);
    const features=preserved.features.get(source.name);if(features){for(const [ref,tip] of Object.entries(features.hyperlinkTips||{}))if(sheet.cells[ref]?.xlsxValue?.value?.hyperlink)sheet.cells[ref].xlsxValue.value.tooltip=tip;delete features.hyperlinkTips;sheet.xlsxFeatures=features;if(!sheet.unsupportedConditionalCount)delete features.conditionalFormatting;if(!Object.keys(features).length)delete sheet.xlsxFeatures;}
    sheet.merges = [...(source.model.merges || [])];
    return sheet;
  });
  if (!sheets.length) sheets.push(newSheet());
  const activeIndex = book.views?.[0]?.activeTab || 0;
  const result = { name: fileName(filename), activeSheetId: sheets[activeIndex]?.id || sheets[0].id, sheets };
  if(book.properties.date1904)result.date1904=true;
  if(styleModels.length)result.xlsxStyles=styleModels;
  if(sheets.some(s=>s.unsupportedConditionalCount)&&preserved.dxfs.length)result.xlsxDxfs=preserved.dxfs;
  await restoreWorksheetParts(data,result);
  await restoreFilterParts(data,result);
  await restoreChartParts(data,result);
  await restorePivotParts(data,result,{refresh:!/\.xlsm$/i.test(filename)});
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
  target.border = Object.fromEntries(Object.entries(cellBorders(style)).map(([side,edge])=>[side,edge?{style:edge.style,color:{argb:'FF'+edge.color.slice(1).toUpperCase()}}:{}]));
  target.numFmt = style.numberFormat && formatKind(style.numberFormat) === (style.format || 'general')
    ? style.numberFormat : formats[style.format] || 'General';
  if (Number.isInteger(style.decimals) && ['number', 'currency', 'accounting', 'percent'].includes(style.format)) {
    const digits = Math.min(12, Math.max(0, style.decimals));
    const fraction = digits ? '.' + '0'.repeat(digits) : '';
    target.numFmt = style.format === 'percent' ? '0' + fraction + '%' :
      (style.format === 'currency' || style.format === 'accounting' ? '₩' : '') + '#,##0' + fraction;
  }
}
const same=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
const fontFields={fontFamily:'name',fontSize:'size',bold:'bold',italic:'italic',underline:'underline',color:'color'};
function fontChanges(style,baseline,font){const changes={};for(const [key,field] of Object.entries(fontFields))if(!same(style[key],baseline[key]))changes[field]=font?.[field];return changes;}
function restoreCellStyle(target,cell,original){
  if(!original)return;const style=cell.style||{},baseline=original.baseline,generated=target.style,base=plainCopy(original.style);
  const changes=fontChanges(style,baseline,generated.font);if(Object.keys(changes).length)base.font={...base.font,...changes};
  if(!same(style.background,baseline.background))base.fill=generated.fill||{type:'pattern',pattern:'none'};
  for(const [key,field] of Object.entries({align:'horizontal',verticalAlign:'vertical',wrap:'wrapText'}))if(!same(style[key],baseline[key])){base.alignment||={};base.alignment[field]=generated.alignment?.[field];}
  if(!same(style.border,baseline.border)||!same(style.borders,baseline.borders))base.border={...base.border,...generated.border};
  if(['format','numberFormat','decimals'].some(key=>!same(style[key],baseline[key])))base.numFmt=generated.numFmt;
  target.style=base;
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
  if (value === undefined || value === '#NAME?' || value === '#ERROR!') value = !cell.cachedFormula || cell.cachedFormula === cell.raw ? cell.cachedValue : undefined;
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
  const evaluate = typeof options.evaluate === 'function' ? options.evaluate : await evaluator(),styleCache=new Map();
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
  // Export a consistent snapshot: cached result cells and native pivot cache use the same source.
  if(workbook.sheets.some(s=>s.pivot)){workbook=cloneData(workbook);for(const s of workbook.sheets)if(s.pivot){try{sourceData(workbook,s.pivot.sourceSheetId,s.pivot.sourceRange);}catch{delete s.pivot;continue;}applyPivot(workbook,s,s.pivot);}}
  const book = new ExcelJS.Workbook();
  book.properties.date1904=workbook.date1904===true;
  book.creator = 'Sheet Studio';
  book.created = new Date(); book.modified = new Date();
  book.calcProperties.fullCalcOnLoad = true;
  book.views = [{ activeTab: Math.max(0, workbook.sheets.findIndex(sheet => sheet.id === workbook.activeSheetId)), visibility: 'visible' }];
  const names = new Set(), exportNames = new Map();
  for (const sheet of workbook.sheets) {
    const target = book.addWorksheet(safeSheetName(sheet.name, names));
    if(sheet.defaultColumnWidth)target.properties.defaultColWidth=(sheet.defaultColumnWidth-5)/7;
    if(sheet.defaultRowHeight)target.properties.defaultRowHeight=sheet.defaultRowHeight*3/4;
    if(sheet.freeze?.rows||sheet.freeze?.cols)target.views=[{state:'frozen',xSplit:sheet.freeze.cols,ySplit:sheet.freeze.rows,topLeftCell:address(sheet.freeze.cols,sheet.freeze.rows),activeCell:'A1'}];
    exportNames.set(sheet.id,target.name);
    for(const [r,protection] of Object.entries(sheet.xlsxProtectionDefaults?.rows||{}))target.getRow(+r+1).protection=plainCopy(protection);
    for(const [c,protection] of Object.entries(sheet.xlsxProtectionDefaults?.columns||{}))target.getColumn(+c+1).protection=plainCopy(protection);
    const pageFilters=sheet.pivot?.filters.filter(f=>!sheet.pivot.rows.includes(f.field)&&!sheet.pivot.columns.includes(f.field))||[];
    const conditionalSheet=pageFilters.length?cloneData(sheet):sheet;if(pageFilters.length)for(let i=0;i<pageFilters.length+1;i++)shiftWorksheetFeatures(conditionalSheet,'r',0,1);exportConditional(target,conditionalSheet);
    pageFilters.forEach((f,i)=>{target.getCell(i+1,1).value=f.field;target.getCell(i+1,2).value=f.values.length===1?String(f.values[0]):'(여러 항목)';});
    for (const [addr, cell] of Object.entries(sheet.cells || {})) {
      if (!coordinates(addr)) continue;
      const pos=coordinates(addr), offset=pageFilters.length?pageFilters.length+1:0;
      const targetCell = target.getCell(offset?address(pos.col,pos.row+offset):addr);
      if (typeof cell.raw === 'string' && cell.raw.startsWith('=')) {
        const result = cachedResult(evaluate, workbook, sheet, addr, cell);
        let formula=cell.raw;
        for(const pivot of workbook.sheets.filter(s=>s.pivot)){const pages=pivot.pivot.filters.filter(f=>!pivot.pivot.rows.includes(f.field)&&!pivot.pivot.columns.includes(f.field)).length;for(let i=0;i<(pages?pages+1:0);i++)formula=transformFormulaReferences(formula,{axis:'r',index:0,delta:1,targetSheet:pivot.name,ownerSheet:sheet.name});}
        const original=cell.xlsxFormula,keepOriginal=original?.raw===formula;
        targetCell.value = { formula: keepOriginal?original.formula:formulaForXlsx(formula,original?.formula).slice(1),...(original?.shareType==='array'?{shareType:'array',ref:original.ref}:{}), ...(result === undefined ? {} : { result }) };
      } else {
        const original=cell.xlsxValue;targetCell.value=original?.raw===cell.raw?plainCopy(original.value):toScalar(cell.raw);
        if(original?.value?.hyperlink&&typeof targetCell.value==='string'&&targetCell.value)targetCell.value={...plainCopy(original.value),text:targetCell.value};
      }
      const autoPercent=typeof cell.raw==='string'&&/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$/.test(cell.raw.trim())&&(!cell.style?.format||cell.style.format==='general');
      const key=String(cell.xlsxStyle??'')+':'+JSON.stringify(targetCell.protection||{})+':'+JSON.stringify(cell.style||{})+(autoPercent?'%':''),cachedStyle=styleCache.get(key);
      if(cachedStyle)targetCell.style=cachedStyle;
      else{applyStyle(targetCell,cell.style);restoreCellStyle(targetCell,cell,workbook.xlsxStyles?.[cell.xlsxStyle]);if(autoPercent)targetCell.numFmt='0.00%';if(styleCache.size<10000)styleCache.set(key,targetCell.style);}
      if(cell.xlsxProtection)targetCell.style={...targetCell.style,protection:plainCopy(cell.xlsxProtection)};
      if(cell.xlsxValue?.value?.richText&&cell.xlsxValue.raw===cell.raw){const original=workbook.xlsxStyles?.[cell.xlsxStyle],delta=fontChanges(cell.style||{},original?.baseline||{},targetCell.font);if(Object.keys(delta).length)targetCell.value={richText:targetCell.value.richText.map(run=>({...run,font:{...run.font,...delta}}))};}
      if (typeof cell.note === 'string') targetCell.note = cell.note;
    }
    for (const [index, width] of Object.entries(sheet.columnWidths || {})) {
      if (/^\d+$/.test(index) && Number(index) < 16384 && Number.isFinite(Number(width))) target.getColumn(Number(index) + 1).width = Math.max(1, ((width===0?sheet.hiddenSizes?.cols?.[index]||90:Number(width)) - 5) / 7);
    }
    for (const [index, height] of Object.entries(sheet.rowHeights || {})) {
      if (/^\d+$/.test(index) && Number(index) < 1048576 && Number.isFinite(Number(height))) target.getRow(Number(index) + 1+(pageFilters.length?pageFilters.length+1:0)).height = Math.max(1, (height===0?sheet.hiddenSizes?.rows?.[index]||sheet.defaultRowHeight||25:Number(height)) * 3 / 4);
    }
    for(const [c,width] of Object.entries(sheet.columnWidths||{}))if(width===0)target.getColumn(+c+1).hidden=true;
    for(const r of sheet.autoFilter?.hiddenRows||[])target.getRow(r+1+(pageFilters.length?pageFilters.length+1:0)).hidden=true;
    for(const [r,height] of Object.entries(sheet.rowHeights||{}))if(height===0)target.getRow(+r+1+(pageFilters.length?pageFilters.length+1:0)).hidden=true;
    for (const merge of sheet.merges || []) {
      if (typeof merge !== 'string' || !/^[A-Z]+\d+:[A-Z]+\d+$/i.test(merge)) continue;
      let shifted=merge;for(let i=0;i<(pageFilters.length?pageFilters.length+1:0);i++)shifted=transformFormulaReferences(shifted,{axis:'r',index:0,delta:1,targetSheet:sheet.name,ownerSheet:sheet.name});target.mergeCells(shifted);
    }
  }
  const pivotBytes = await addPivotParts(new Uint8Array(await book.xlsx.writeBuffer()),workbook,exportNames);
  const data = await addWorksheetParts(await addFilterParts(await addChartParts(pivotBytes,workbook,exportNames),workbook,exportNames),workbook,exportNames);
  return { data:await addPreservedFeatures(data,workbook,exportNames), mime: XLSX_MIME, extension: 'xlsx' };
}
