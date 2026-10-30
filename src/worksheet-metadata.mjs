import {SaxesParser} from 'saxes';
const escape = value => String(value).replace(/[&<>"\r\n\t]/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\r':'&#13;','\n':'&#10;','\t':'&#9;'}[char]));

// Supplemental importers only need worksheet metadata and hidden row tags.
// Do not construct another million-cell DOM alongside ExcelJS's cell model.
export function worksheetMetadataXML(text, {wantCell = () => false, onCell = () => {}, onRow = () => {}, onAnyRow=()=>{},onColumn=()=>{},keepHiddenRows = false, onDataRange = () => {}} = {}) {
  const parser = new SaxesParser({xmlns: true}), output = [];
  let inData = false, cell = null, field = '', depth = 0, dataStart;
  const attrs = node => Object.fromEntries(Object.values(node.attributes).map(attr => [attr.name, attr.value]));
  const tag = node => '<' + node.name + Object.entries(attrs(node)).map(([key,value]) => ` ${key}="${escape(value)}"`).join('') + '>';
  parser.on('doctype', () => { throw Error('DOCTYPE이 포함된 워크시트는 지원하지 않습니다.'); });
  parser.on('opentag', node => {
    depth++;
    if (!inData) {
      if(node.local==='col')onColumn(attrs(node));
      output.push(tag(node));
      if (depth === 2 && node.local === 'sheetData') { inData = true; dataStart = text.lastIndexOf('<', parser.position - 1); }
      return;
    }
    if (node.local === 'row') {
      const row = attrs(node);
      onAnyRow(row);
      if (['1','true'].includes(row.hidden)) { onRow(row); if (keepHiddenRows) output.push(tag(node) + `</${node.name}>`); }
    }
    if (node.local === 'c') { const attributes = attrs(node); if (wantCell(attributes)) cell = {...attributes, value: '', formula: false}; }
    if (cell && ['v','f'].includes(node.local)) { field = node.local; if (field === 'f') cell.formula = true; }
  });
  parser.on('closetag', node => {
    if (inData && depth === 2 && node.local === 'sheetData') { inData = false; output.push(`</${node.name}>`); onDataRange({start: dataStart, end: parser.position}); }
    else if (!inData) output.push(`</${node.name}>`);
    else if (node.local === 'c') { if (cell) onCell(cell); cell = null; field = ''; }
    else if (node.local === field) field = '';
    depth--;
  });
  parser.on('text', value => { if (!inData) output.push(depth ? escape(value) : value); else if (cell && field === 'v') cell.value += value; });
  parser.on('cdata', value => { if (!inData) output.push(escape(value)); else if (cell && field === 'v') cell.value += value; });
  parser.write(text).close(); return output.join('');
}

// Edit the small metadata DOM while keeping every original cell byte intact.
export function splitWorksheetXML(text,options={}) {
  let data;
  const metadata = worksheetMetadataXML(text, {...options,onDataRange: range => { data = range; }});
  return {metadata, restore(serialized) {
    if (!data) return serialized;
    let destination;
    worksheetMetadataXML(serialized, {onDataRange: range => { destination = range; }});
    if (!destination) throw Error('워크시트 셀 영역을 보존할 수 없습니다.');
    return serialized.slice(0, destination.start) + text.slice(data.start, data.end) + serialized.slice(destination.end);
  }};
}
