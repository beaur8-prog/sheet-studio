import {assertBlankStart,recoverPrevious} from './recovery-ui.mjs';
import { _electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import platform from '../scripts/platform.cjs';

const root = path.resolve(import.meta.dirname, '..'), packaged = process.argv.includes('--packaged');
const version = createRequire(import.meta.url)('../package.json').version;
await fs.mkdir(path.join(root, '.package'), { recursive: true });
const fixture = await fs.mkdtemp(path.join(root, '.package/file-dialogs-'));
const dataDir = path.join(fixture, 'profile'), stateFile = path.join(dataDir, 'file-locations.json');
const [a, b, c, d, bad] = ['A 한글', 'B', 'C 저장', 'D', 'invalid'].map(name => path.join(fixture, name));
for (const directory of [a, b, c, d, bad]) await fs.mkdir(directory);
await fs.writeFile(path.join(a, 'first.csv'), '이름,금액\n첫 문서,17');
await fs.writeFile(path.join(b, 'second.csv'), '이름,금액\n두 번째 문서,42');
await fs.writeFile(path.join(bad, 'invalid.sheet.json'), '{broken');
const env = { ...process.env, SHEET_STUDIO_DATA: dataDir }; delete env.ELECTRON_RUN_AS_NODE;
const executablePath = process.env.SHEET_STUDIO_EXECUTABLE || (packaged ? platform.packagedExecutable(root, version) : path.join(platform.findRuntime(root), 'dist', platform.runtimeSuffix()));
let app, page, documents; const errors = [], checks = [];
async function launch() {
  app = await _electron.launch({ executablePath, args: packaged ? [] : [root], env });
  page = await app.firstWindow(); page.setDefaultTimeout(20000);
  page.on('pageerror', error => errors.push(error.message));
  await page.locator('.cell').first().waitFor();
  documents = await app.evaluate(({ app }) => app.getPath('documents'));
}
const id = () => page.evaluate(() => localStorage.getItem('sheet-studio.document-location'));
const menu = action => app.evaluate(({ BrowserWindow }, action) => BrowserWindow.getAllWindows()[0].webContents.send('sheet:menu-action', action), action);
async function mock(kind, result) {
  await app.evaluate(({ dialog }, { kind, result }) => {
    globalThis.fileDialogOptions = null;
    dialog[kind === 'open' ? 'showOpenDialog' : 'showSaveDialog'] = async (_window, options) => {
      globalThis.fileDialogOptions = options; return result;
    };
  }, { kind, result });
  await page.evaluate(() => document.querySelector('.toast')?.remove());
}
const options = () => app.evaluate(() => globalThis.fileDialogOptions);
async function inspectDirectory(expected, label) {
  await mock('open', { canceled: true, filePaths: [] });
  assert.equal(await page.evaluate(documentId => desktop.openFile(documentId), await id()), null);
  assert.equal((await options()).defaultPath, expected, label + ': open');
  await mock('save', { canceled: true });
  assert.equal(await page.evaluate(documentId => desktop.saveFile({ documentId, name: '다음.xlsx', data: [] }), await id()), null);
  assert.equal((await options()).defaultPath, path.join(expected, '다음.xlsx'), label + ': save');
  checks.push(label);
}
async function preferences(expected) {
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) {
    try { const state = JSON.parse(await fs.readFile(stateFile, 'utf8')); if (state.lastDirectory === expected) return state; } catch {}
    await new Promise(resolve => setTimeout(resolve, 20));
  }
  throw Error('폴더 기록을 기다리는 중 시간 초과: ' + expected);
}
async function open(file, content) {
  await mock('open', { canceled: false, filePaths: [file] });
  await menu('open');
  await page.waitForFunction(value => sheetStudio.getCell('A2')?.raw === value, content);
  await page.locator('.toast').filter({ hasText: '파일을 열었습니다.' }).waitFor();
  await preferences(path.dirname(file));
}
async function save(file, format = 'xlsx') {
  await mock('save', { canceled: false, filePath: file });
  if (format === 'xlsx') await menu('save');
  else {
    await page.getByRole('button', { name: '파일', exact: true }).click();
    await page.getByRole('button', { name: format === 'csv' ? 'CSV 저장' : '앱 문서 저장 (.sheet.json)', exact: true }).click();
  }
  await page.locator('.toast').filter({ hasText: '파일을 저장했습니다.' }).waitFor();
  assert.ok((await fs.stat(file)).size > 0);
  return options();
}
try {
  await launch();
  if (process.platform === 'linux') {
    assert.equal(await app.evaluate(({ app }) => app.commandLine.getSwitchValue('xdg-portal-required-version')), '4');
    checks.push('older Linux portals fall back to a chooser that supports defaultPath');
  }
  await inspectDirectory(documents, 'first use falls back to Documents');
  await open(path.join(a, 'first.csv'), '첫 문서'); const firstId = await id();
  await inspectDirectory(a, 'opened workbook folder and canceled dialogs');
  await mock('open', { canceled: false, filePaths: [path.join(bad, 'invalid.sheet.json')] }); await menu('open');
  await page.locator('.toast').filter({ hasText: '파일을 열 수 없습니다:' }).waitFor();
  assert.equal(await id(), firstId); assert.equal(await page.evaluate(() => sheetStudio.getCell('A2').raw), '첫 문서');
  await inspectDirectory(a, 'invalid workbook does not change folder');
  await mock('open', { canceled: false, filePaths: [path.join(bad, 'missing.xlsx')] }); await menu('open');
  await page.locator('.toast').waitFor();
  await inspectDirectory(a, 'unreadable file does not change folder');
  await mock('save', { canceled: false, filePath: path.join(bad, 'missing', 'failed.xlsx') }); await menu('save');
  await page.locator('.toast').filter({ hasText: '저장 실패:' }).waitFor();
  await inspectDirectory(a, 'failed write does not change folder');
  assert.equal((await preferences(a)).lastDirectory, a);

  await open(path.join(b, 'second.csv'), '두 번째 문서'); const secondId = await id(); assert.notEqual(firstId, secondId);
  await menu('undo'); await page.waitForFunction(() => sheetStudio.getCell('A2')?.raw === '첫 문서');
  assert.equal(await id(), firstId);
  await inspectDirectory(a, 'undo restores current document folder before most recent folder');
  await menu('redo'); await page.waitForFunction(() => sheetStudio.getCell('A2')?.raw === '두 번째 문서');
  await inspectDirectory(b, 'redo restores second document folder');
  const saved = await save(path.join(c, 'second.xlsx'));
  assert.equal(saved.defaultPath, path.join(b, 'second.xlsx'));
  await inspectDirectory(c, 'Save As updates current folder');
  const csv = await save(path.join(c, 'second.csv'), 'csv'); assert.equal(csv.defaultPath, path.join(c, 'second.csv'));
  const json = await save(path.join(c, 'second.sheet.json'), 'json'); assert.equal(json.defaultPath, path.join(c, 'second.sheet.json'));
  const exported = await fs.readFile(path.join(c, 'second.sheet.json'), 'utf8');
  assert.equal(exported.includes(fixture), false); assert.equal(exported.includes(secondId), false);
  checks.push('XLSX, CSV and app-document save actions retain folder without exporting paths or IDs');

  await app.close(); app = null; await launch();
  await assertBlankStart(page);await recoverPrevious(page);
  assert.equal(await id(), secondId); assert.equal(await page.evaluate(() => sheetStudio.getCell('A2').raw), '두 번째 문서');
  await inspectDirectory(c, 'restart retains recovered workbook folder');
  await menu('new'); await page.getByRole('button', { name: '새로 만들기', exact: true }).click();
  await page.waitForFunction(() => sheetStudio.getWorkbook().name === '통합 문서1'); assert.notEqual(await id(), secondId);
  await inspectDirectory(c, 'new workbook starts in last successful folder');
  await save(path.join(d, 'new.xlsx')); await inspectDirectory(d, 'new workbook remembers its saved folder');
  await menu('undo'); await page.waitForFunction(() => sheetStudio.getCell('A2')?.raw === '두 번째 문서');
  await inspectDirectory(c, 'undo new workbook restores original folder after another save');
  await page.reload(); await page.locator('.cell').first().waitFor(); await inspectDirectory(c, 'reload preserves document location');
  await fs.rm(c, { recursive: true }); await inspectDirectory(d, 'missing current folder falls back to last folder');
  await fs.rm(d, { recursive: true }); await inspectDirectory(documents, 'missing current and last folders fall back to Documents');
  await app.close(); app = null;
  await fs.writeFile(stateFile, '{invalid preferences'); await launch();
  await inspectDirectory(documents, 'damaged preferences do not block file dialogs');
  assert.deepEqual(errors, []);
  await fs.mkdir(path.join(root, 'test-results'), { recursive: true });
  await fs.writeFile(path.join(root, 'test-results/file-dialogs.json'), JSON.stringify({ ok: true, platform: process.platform, packaged, checks }, null, 2));
  console.log(JSON.stringify({ ok: true, platform: process.platform, packaged, checks }, null, 2));
} finally {
  if (app) await app.close();
  await fs.rm(fixture, { recursive: true, force: true });
}
