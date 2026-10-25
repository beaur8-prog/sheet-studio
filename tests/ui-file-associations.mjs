import { _electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { pathToFileURL } from 'node:url';
import ExcelJS from 'exceljs';
import { createWorkbook } from '../src/engine.mjs';
import platform from '../scripts/platform.cjs';

const root = path.resolve(import.meta.dirname, '..');
const packaged = process.argv.includes('--packaged'), baseline = process.argv.includes('--baseline');
const version = createRequire(import.meta.url)('../package.json').version;
await fs.mkdir(path.join(root, '.package'), { recursive: true });
const fixture = await fs.mkdtemp(path.join(root, '.package/file-associations-'));
const documents = path.join(fixture, '한글과 공백이 있는 폴더');
await fs.mkdir(documents);
const profile = path.join(fixture, 'profile');
const env = { ...process.env, SHEET_STUDIO_DATA: profile }; delete env.ELECTRON_RUN_AS_NODE;
const executablePath = process.env.SHEET_STUDIO_EXECUTABLE || (packaged
  ? platform.packagedExecutable(root, version) : path.join(platform.findRuntime(root), 'dist', platform.runtimeSuffix()));
const args = files => [...(packaged ? [] : [root]), ...files];
const first = path.join(documents, '첫 번째 외부 문서.XLSX');
const second = path.join(documents, '두 번째 문서.csv');
const tsv = path.join(documents, '세 번째 문서.tsv');
const json = path.join(documents, '네 번째 문서.sheet.json');
const broken = path.join(documents, '깨진 문서.xlsx');
const oversized = path.join(documents, '너무 큰 문서.xlsx');
const xlsx = new ExcelJS.Workbook(), worksheet = xlsx.addWorksheet('한글 시트');
worksheet.getCell('A1').value = '외부 XLSX'; worksheet.getCell('A1').font = { bold: true };
worksheet.getCell('A2').value = 17; worksheet.getCell('B2').value = { formula: 'A2*2', result: 34 };
worksheet.getColumn(1).width = 24;
await xlsx.xlsx.writeFile(first);
await fs.writeFile(second, '외부 CSV,42\n두 번째 행,84');
await fs.writeFile(tsv, '외부 TSV\t53');
const book = createWorkbook(); book.name = '외부 앱 문서'; book.sheets[0].cells.A1 = { raw: '외부 JSON' };
await fs.writeFile(json, JSON.stringify(book));
await fs.writeFile(broken, 'not a workbook');
const large = await fs.open(oversized, 'w'); await large.truncate(100 * 1024 * 1024 + 1); await large.close();
let app, page; const errors = [], checks = [];
async function launch(files = []) {
  app = await _electron.launch({ executablePath, args: args(files), env, timeout: 45000 });
  page = await app.firstWindow(); page.setDefaultTimeout(20000);
  page.on('pageerror', error => errors.push(error.message));
  await page.waitForFunction(() => !!globalThis.sheetStudio);
}
const emit = (...files) => app.evaluate(({ app }, files) => files.map(file => app.emit('open-file', { preventDefault() {} }, file)), files);
const shown = async value => {
  await page.waitForFunction(value => globalThis.sheetStudio?.getCell('A1')?.raw === value, value);
  await page.locator('.workbook-busy').waitFor({ state: 'detached' });
  await page.evaluate(() => sheetStudio.whenIdle());
};
const menu = action => app.evaluate(({ BrowserWindow }, action) => BrowserWindow.getAllWindows()[0].webContents.send('sheet:menu-action', action), action);
async function edit(value) {
  await page.locator('#name-box').fill('A1'); await page.locator('#name-box').press('Enter');
  await page.keyboard.press('F2'); await page.locator('#editor').fill(value);
}
const triggerDialog = async kind => {
  await app.evaluate(({ dialog }, kind) => {
    globalThis.releaseFileDialog = null;
    dialog[kind] = () => new Promise(resolve => { globalThis.releaseFileDialog = resolve; });
  }, kind);
  await menu(kind === 'showOpenDialog' ? 'open' : 'save');
  const deadline = Date.now() + 10000;
  while (!await app.evaluate(() => !!globalThis.releaseFileDialog)) {
    if (Date.now() > deadline) throw Error('Native dialog did not open');
    await page.waitForTimeout(25);
  }
};
try {
  await launch([first]);
  if (baseline) {
    await page.evaluate(() => sheetStudio.whenIdle());
    assert.notEqual(await page.evaluate(() => sheetStudio.getCell('A1')?.raw), '외부 XLSX');
    assert.deepEqual(await emit(first), [false]);
    checks.push('reproduced: startup path ignored and macOS open-file has no listener');
  } else {
    await shown('외부 XLSX');
    assert.equal(await page.evaluate(() => sheetStudio.getCell('A1').style.bold), true);
    assert.equal(await page.evaluate(() => sheetStudio.evaluateAsync('B2')), 34);
    checks.push('cold launch opens XLSX with Korean/space path, formatting and formula');
    await edit('입력 중인 원본을 보존');
    await emit(second); await shown('외부 CSV');
    await menu('undo'); await shown('입력 중인 원본을 보존');
    await menu('redo'); await shown('외부 CSV');
    checks.push('running open commits editing; undo/redo preserves the previous workbook');
    await emit(broken); await page.locator('.toast').filter({ hasText: '파일을 열 수 없습니다:' }).waitFor();
    assert.equal(await page.evaluate(() => sheetStudio.getCell('A1').raw), '외부 CSV');
    await emit(path.join(documents, '없는 파일.xlsx'));
    await page.locator('.toast').filter({ hasText: 'ENOENT' }).waitFor();
    await emit(oversized); await page.locator('.toast').filter({ hasText: '100 MB' }).waitFor();
    assert.equal(await page.evaluate(() => sheetStudio.getCell('A1').raw), '외부 CSV');
    await emit(tsv, json); await shown('외부 JSON');
    await menu('undo'); await shown('외부 TSV');
    await menu('redo'); await shown('외부 JSON');
    checks.push('invalid/missing/oversized files preserve the workbook; queued TSV and JSON open in order');
    await app.evaluate(({ dialog }) => {
      dialog.showOpenDialog = async (_window, options) => { globalThis.fileDialogOptions = options; return { canceled: true }; };
    });
    await page.evaluate(() => desktop.openFile(localStorage.getItem('sheet-studio.document-location')));
    assert.equal(await app.evaluate(() => globalThis.fileDialogOptions.defaultPath), documents);
    assert.equal(await page.evaluate(async () => { try { await desktop.takeRequestedFile('arbitrary-path'); return false; } catch { return true; } }), false);
    // The public bridge accepts no path. Direct IPC rejects extra arguments.
    assert.equal(await app.evaluate(async ({ ipcMain, BrowserWindow }) => {
      const window = BrowserWindow.getAllWindows()[0];
      try { await ipcMain._invokeHandlers.get('sheet:take-requested-file')({ sender: window.webContents, senderFrame: window.webContents.mainFrame }, '/etc/passwd'); return false; }
      catch { return true; }
    }), true);
    checks.push('external document folder becomes chooser default; renderer cannot choose an arbitrary path');
    // A renderer modal must retain the original document until dismissed.
    await menu('find'); await page.getByRole('dialog').waitFor();
    await emit(first); await page.waitForTimeout(250);
    assert.equal(await page.evaluate(() => sheetStudio.getCell('A1').raw), '외부 JSON');
    await page.keyboard.press('Escape'); await shown('외부 XLSX');
    checks.push('open requests wait for an editor dialog to close');
    for (const kind of ['showOpenDialog', 'showSaveDialog']) {
      await triggerDialog(kind); await emit(second); await page.waitForTimeout(250);
      assert.equal(await page.evaluate(() => sheetStudio.getCell('A1').raw), '외부 XLSX');
      await app.evaluate(() => globalThis.releaseFileDialog({ canceled: true, filePaths: [] }));
      await shown('외부 CSV'); await emit(first); await shown('외부 XLSX');
    }
    checks.push('native open/save dialogs retain requests until cancellation, including busy save');
    await app.evaluate(({ app }, files) => app.emit('second-instance', {}, ['sheet-studio', '--ignored', ...files], process.cwd()), [pathToFileURL(first).href, second]);
    await shown('외부 CSV');
    // Launch a real second process, as Windows/Linux associations do.
    const child = spawn(executablePath, args([first]), { env, stdio: 'ignore' });
    await new Promise((resolve, reject) => { child.once('error', reject); child.once('exit', code => code === 0 ? resolve() : reject(Error('Second instance exit ' + code))); });
    await shown('외부 XLSX');
    assert.equal(await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length), 1);
    checks.push('second-instance file URLs and real second process reuse one editor');
    if (process.platform === 'darwin') {
      const closed = page.waitForEvent('close');
      await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close()); await closed;
      const reopened = app.waitForEvent('window'); await emit(second);
      page = await reopened; await shown('외부 CSV');
      checks.push('macOS file open recreates a closed editor and waits for recovery');
    }
    await page.evaluate(() => sheetStudio.flushRecovery());
    await app.close(); app = null;
    await launch([first]); await shown('외부 XLSX');
    checks.push('startup request takes precedence over the recovered workbook');
    await page.evaluate(() => sheetStudio.flushRecovery());
    await app.close(); app = null;
    await launch(); await shown('외부 XLSX');
    checks.push('plain launch retains normal recovery behavior');
    assert.deepEqual(errors, []);
  }
  await fs.mkdir(path.join(root, 'test-results'), { recursive: true });
  const result = { ok: true, baseline, packaged, platform: process.platform, checks };
  await fs.writeFile(path.join(root, 'test-results/file-associations' + (baseline ? '-baseline' : '') + '.json'), JSON.stringify(result, null, 2) + '\n');
  console.log(JSON.stringify(result, null, 2));
} finally {
  if (app) await app.close();
  await fs.rm(fixture, { recursive: true, force: true });
}
