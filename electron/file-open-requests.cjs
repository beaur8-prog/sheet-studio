'use strict';
const path = require('node:path');
const { fileURLToPath } = require('node:url');

// Only OS events and process arguments produce paths. Keep bytes out of the
// queue until the editor is ready, and coalesce duplicate pending requests.
class FileOpenRequests {
  constructor() { this.paths = new Set(); }
  get size() { return this.paths.size; }
  add(filename, directory = process.cwd()) {
    if (typeof filename !== 'string' || !filename || filename.startsWith('-') || filename.includes('\0')) return false;
    try {
      if (filename.startsWith('file:')) filename = fileURLToPath(filename);
      else if (/^[a-z][a-z0-9+.-]*:\/\//i.test(filename)) return false;
      filename = path.resolve(directory, filename);
      if (!/\.(xlsx|xlsm|csv|tsv|json)$/i.test(filename) || this.paths.has(filename) || this.paths.size >= 64) return false;
      this.paths.add(filename);
      return true;
    } catch { return false; }
  }
  addArguments(argv, directory) {
    for (const argument of argv.slice(1)) this.add(argument, directory);
  }
  take() {
    const filename = this.paths.values().next().value;
    if (filename) this.paths.delete(filename);
    return filename;
  }
}

// Alternate offers Open With without replacing an existing preferred editor.
const fileAssociations = [
  { ext: 'xlsx', name: 'Excel Workbook', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
  { ext: 'xlsm', name: 'Excel Macro-Enabled Workbook', mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12' },
  { ext: 'csv', name: 'CSV Spreadsheet', mimeType: 'text/csv' },
  { ext: 'tsv', name: 'TSV Spreadsheet', mimeType: 'text/tab-separated-values' },
].map(type => ({ ...type, role: 'Editor', rank: 'Alternate' }));

module.exports = { FileOpenRequests, fileAssociations };
