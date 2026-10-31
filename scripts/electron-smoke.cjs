'use strict';
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const assert = require('node:assert/strict');
const { _electron } = require('playwright');
const root = path.resolve(__dirname, '..');
const locations = [root, path.dirname(root), path.join(path.dirname(root), 'word-studio'), path.join(path.dirname(root), 'slide-studio')];
const {findRuntime,runtimeSuffix,packagedExecutable}=require('./platform.cjs');
const runtime=path.join(findRuntime(root),'dist',runtimeSuffix());
const packaged=packagedExecutable(root,require('../package.json').version);

async function smoke() {
  const environment = { ...process.env, SHEET_STUDIO_DATA: path.join(root, '.package/smoke-data') };
  delete environment.ELECTRON_RUN_AS_NODE;
  const usingPackage = process.argv.includes('--packaged');
  const application = await _electron.launch({
    executablePath: process.env.SHEET_STUDIO_EXECUTABLE || (usingPackage ? packaged : runtime),
    args: usingPackage ? [] : [root], env: environment,
  });
  try {
    const page = await application.firstWindow();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.waitForSelector('.cell', { timeout: 20_000 });
    await page.waitForFunction(() => !!window.desktop?.openFile);
    const state = await page.evaluate(() => ({
      title: document.title, cells: document.querySelectorAll('.cell').length,
      nodeIntegration: typeof window.require, bridge: Object.keys(window.desktop).sort(),
      body: document.body.innerText.slice(0, 350),
    }));
    const expectedVersion = require('../package.json').version;
    assert.equal(await application.evaluate(({ app }) => app.getVersion()), expectedVersion);
    assert.equal(await page.evaluate(() => desktop.getAppVersion()), expectedVersion);
    assert.equal(await page.locator('#app-version').innerText(), 'v'+expectedVersion);
    assert.ok(state.title.endsWith(' — Sheet Studio '+expectedVersion));
    assert.equal(await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].getTitle()), state.title);
    await page.locator('[data-tab="도움말"]').click();
    await page.locator('[data-action="about"]').click();
    assert.equal(await page.locator('[role="dialog"] .pill').innerText(), expectedVersion);
    await page.locator('[data-close]').click();
    await page.locator('[data-tab="홈"]').click();
    assert.equal(state.nodeIntegration, 'undefined');
    assert.deepEqual(state.bridge, ['confirmFile', 'editText', 'getAppVersion', 'onBeforeClose', 'onMenuAction', 'onOpenRequested', 'onWindowBlur', 'openFile', 'pasteClipboard', 'saveFile', 'setTitle', 'takeRequestedFile', 'vbaCancel', 'vbaCommand', 'vbaProbe', 'vbaResult', 'vbaStart', 'vbaStatus', 'writeClipboard']);
    assert.equal(typeof (await page.evaluate(() => window.desktop.vbaProbe())).available, 'boolean');
    assert.ok(state.cells > 0, 'Spreadsheet grid rendered');
    await page.evaluate(() => window.desktop.setTitle('Sheet Studio – smoke test'));
    assert.match(await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].getTitle()), /smoke test/);
    const fixtures = path.join(root, '.package/smoke-files');
    await fsp.mkdir(fixtures, { recursive: true });
    const source = path.join(fixtures, 'source.csv'), destination = path.join(fixtures, 'saved.csv');
    await fsp.writeFile(source, '이름,금액\r\n테스트,42', 'utf8');
    await application.evaluate(({ dialog }, paths) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [paths.source] });
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: paths.destination });
    }, { source, destination });
    const opened = await page.evaluate(async () => {const file=await window.desktop.openFile();return {...file,data:Array.from(file.data)};});
    assert.equal(opened.name, 'source.csv');
    assert.equal(Buffer.from(opened.data).toString('utf8'), '이름,금액\r\n테스트,42');
    const saved = await page.evaluate(data => window.desktop.saveFile({ name: 'saved.csv', data }), opened.data);
    assert.equal(saved, destination);
    assert.equal(await fsp.readFile(destination, 'utf8'), '이름,금액\r\n테스트,42');
    const invalid = await page.evaluate(async () => {
      try { await window.desktop.saveFile({ name: 'invalid.csv', data: [999] }); return false; }
      catch { return true; }
    });
    assert.equal(invalid, true);
    await application.evaluate(({ dialog }) => {
      dialog.showOpenDialog = async () => ({ canceled: true, filePaths: [] });
      dialog.showSaveDialog = async () => ({ canceled: true });
    });
    assert.equal(await page.evaluate(() => window.desktop.openFile()), null);
    assert.equal(await page.evaluate(() => window.desktop.saveFile({ name: 'cancel.csv', data: [] })), null);
    // Exercise the same menu IPC used by macOS, on every test host.
    const menuAction = action => application.evaluate(({ BrowserWindow }, action) => BrowserWindow.getAllWindows()[0].webContents.send('sheet:menu-action', action), action);
    await page.locator('#book-title').focus();
    await menuAction('select-all');
    await page.waitForFunction(() => { const el=document.querySelector('#book-title');return el.selectionStart===0&&el.selectionEnd===el.value.length; });
    await page.locator('#name-box').fill('A1');await page.locator('#name-box').press('Enter');
    await page.keyboard.press('F2');await page.locator('#editor').fill('17');await page.locator('#editor').press('Enter');
    await page.locator('#name-box').fill('A1');await page.locator('#name-box').press('Enter');
    // Keep the user's OS clipboard untouched while exercising menu -> preload
    // -> edit-text IPC -> paste event and the real renderer's paste handling.
    await application.evaluate(({ BrowserWindow }) => {
      const contents=BrowserWindow.getAllWindows()[0].webContents;
      contents.paste=()=>contents.executeJavaScript('(()=>{const data=new DataTransfer();data.setData("text/plain","42");document.querySelector("#grid").dispatchEvent(new ClipboardEvent("paste",{bubbles:true,cancelable:true,clipboardData:data}));})()');
    });
    await menuAction('paste');await page.waitForFunction(() => sheetStudio.evaluate('A1')===42);
    await menuAction('undo');await page.waitForFunction(() => sheetStudio.evaluate('A1')===17);
    await menuAction('redo');await page.waitForFunction(() => sheetStudio.evaluate('A1')===42);
    await menuAction('find');await page.getByRole('dialog').waitFor();
    await menuAction('new');assert.equal(await page.getByRole('dialog').count(),1);
    assert.equal(await page.evaluate(async () => {try {await desktop.editText('executeJavaScript');return false;}catch{return true;}}),true);
    await page.keyboard.press('Escape');
    const output = path.join(root, 'test-results');
    await fsp.mkdir(output, { recursive: true });
    await page.screenshot({ path: path.join(output, usingPackage ? 'desktop-packaged.png' : 'desktop.png') });
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({ ok: true, packaged: usingPackage, ...state }, null, 2));
  } finally { await application.close(); }
}

smoke().catch(error => { console.error(error); process.exitCode = 1; });
