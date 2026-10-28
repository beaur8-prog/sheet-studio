'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { execFileSync } = require('node:child_process');
const { fileAssociations } = require('../electron/file-open-requests.cjs');
const { packagedExecutable } = require('./platform.cjs');
const version = require('../package.json').version;

(async () => {
  assert.equal(process.platform, 'darwin');
  const root = path.resolve(__dirname, '..');
  const executable = process.env.SHEET_STUDIO_EXECUTABLE || packagedExecutable(root, version);
  const bundle = path.resolve(path.dirname(executable), '../..');
  const metadata = JSON.parse(execFileSync('/usr/bin/plutil', ['-convert', 'json', '-o', '-', path.join(bundle, 'Contents/Info.plist')], { encoding: 'utf8' }));
  for (const type of fileAssociations) {
    const entry = metadata.CFBundleDocumentTypes?.find(entry => entry.CFBundleTypeExtensions?.includes(type.ext));
    assert.ok(entry, `Missing document type: ${type.ext}`);
    assert.equal(entry.CFBundleTypeRole, 'Editor'); assert.equal(entry.LSHandlerRank, 'Alternate');
  }
  const report = { version, metadata: true };
  if (!process.argv.includes('--metadata-only')) {
    assert.equal(process.env.CI, 'true', 'Run Launch Services registration checks only on an isolated CI Mac');
    const fixtures = await fs.mkdtemp(path.join(os.tmpdir(), 'sheet-file-support-'));
    const registration = '/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister';
    try {
      const ExcelJS = require('exceljs'), book = new ExcelJS.Workbook();
      book.addWorksheet('검증').getCell('A1').value = '지원 앱 등록 확인';
      const files = fileAssociations.map(type => path.join(fixtures, '한글 파일.' + type.ext));
      for (const file of files) {
        if (/\.xls[xm]$/.test(file)) await book.xlsx.writeFile(file);
        else await fs.writeFile(file, '지원 앱 등록 확인,42');
      }
      report.launchServices = JSON.parse(execFileSync('/usr/bin/swift', [path.join(__dirname, 'check-mac-file-support.swift'), bundle, ...files], { encoding: 'utf8', timeout: 90000 }));
    } finally {
      execFileSync(registration, ['-u', bundle]);
      await fs.rm(fixtures, { recursive: true, force: true });
    }
  }
  await fs.mkdir(path.join(root, 'test-results'), { recursive: true });
  await fs.writeFile(path.join(root, 'test-results/mac-file-support.json'), JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report, null, 2));
})().catch(error => { console.error(error); process.exitCode = 1; });
