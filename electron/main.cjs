'use strict';
const { app, BrowserWindow, ipcMain, dialog, Menu, clipboard, shell, Notification } = require('electron');
const fs = require('node:fs/promises');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const vba = require('./vba.cjs');
const { applicationMenu } = require('./menu.cjs');
const { createFileLocations } = require('./file-locations.cjs');
const { createStudioUpdater } = require('./studio-updater.cjs');

const MAX_FILE_BYTES = 100 * 1024 * 1024;
const filters = [
  { name: '스프레드시트', extensions: ['xlsx', 'xlsm', 'csv', 'tsv', 'json'] },
  { name: 'Excel 통합 문서', extensions: ['xlsx', 'xlsm'] },
  { name: 'CSV', extensions: ['csv'] },
  { name: 'Sheet Studio', extensions: ['sheet.json', 'json'] },
];
let win, quitting = false, closing = null, recoveryReady = false;
const rendererPath = path.join(__dirname, '../dist/index.html');
const rendererURL = pathToFileURL(rendererPath).href;
app.setName('Sheet Studio');
// GNOME matches this ID to sheet-studio.desktop for its icon and Dash pinning.
// It sets the Wayland app_id and the X11 WM_CLASS before any windows are made.
if (process.platform === 'linux') {
  app.setDesktopName('sheet-studio.desktop');
  // Older portal file choosers ignore defaultPath; use GTK/KDE there instead.
  app.commandLine.appendSwitch('xdg-portal-required-version', '4');
}
// Match the installer shortcut identity so Windows can group and pin this app.
if (process.platform === 'win32') app.setAppUserModelId('local.lattice.sheetstudio');
if (process.env.SHEET_STUDIO_DATA) app.setPath('userData', path.resolve(process.env.SHEET_STUDIO_DATA));
const fileLocations = createFileLocations({ stateFile: path.join(app.getPath('userData'), 'file-locations.json'), documents: app.getPath('documents') });
const updates = createStudioUpdater({ app, dialog, shell, Notification, id: 'sheet', getWindow: () => win });
function installMenu() {
  const template = applicationMenu(process.platform, action => win?.webContents.send('sheet:menu-action', action)) || [];
  template.push({ label: '도움말', submenu: updates.menuTemplate() });
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}
updates.onChange(installMenu);

function trusted(event) {
  if (event.sender !== win?.webContents || event.senderFrame !== win.webContents.mainFrame ||
      event.senderFrame.url.split('#')[0] !== rendererURL) throw new Error('허용되지 않는 앱 요청입니다.');
}

ipcMain.handle('sheet:open-file', async (event, documentId) => {
  trusted(event);
  const result = await dialog.showOpenDialog(win, { title: '통합 문서 열기', defaultPath: await fileLocations.directory(documentId), properties: ['openFile'], filters });
  if (result.canceled || !result.filePaths.length) return null;
  const file = result.filePaths[0];
  if ((await fs.stat(file)).size > MAX_FILE_BYTES) throw new Error('100 MB 이하의 통합 문서를 선택하세요.');
  const data=await fs.readFile(file);
  const originToken=/\.xlsm$/i.test(file)?await vba.rememberFile(file,data):undefined;
  return { name: path.basename(file), data: new Uint8Array(data), locationToken: fileLocations.stage(file), ...(originToken?{originToken}:{}) };
});

ipcMain.handle('sheet:confirm-file', async (event, payload) => { trusted(event); await fileLocations.confirm(payload); });

ipcMain.handle('sheet:save-file', async (event, payload) => {
  trusted(event);
  if (!payload || typeof payload.name !== 'string' || !(payload.data instanceof Uint8Array || Array.isArray(payload.data)) ||
      payload.data.length > MAX_FILE_BYTES || Array.isArray(payload.data) && payload.data.some(value => !Number.isInteger(value) || value < 0 || value > 255)) {
    throw new Error('저장할 파일 데이터가 올바르지 않습니다.');
  }
  const basename = path.basename(payload.name).replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_') || '통합 문서.xlsx';
  const extension = basename.toLowerCase().endsWith('.sheet.json') ? 'json' : path.extname(basename).slice(1);
  const result = await dialog.showSaveDialog(win, {
    title: '통합 문서 저장', defaultPath: path.join(await fileLocations.directory(payload.documentId), basename),
    filters: [{ name: extension === 'xlsx' ? 'Excel 통합 문서' : extension.toUpperCase(), extensions: [extension || 'xlsx'] }],
  });
  if (result.canceled || !result.filePath) return null;
  const tempPath = path.join(path.dirname(result.filePath), '.' + path.basename(result.filePath) + '.' + randomUUID() + '.tmp');
  try {
    await fs.writeFile(tempPath, Buffer.from(payload.data), { flag: 'wx' });
    // XLSM exports keep their internet-origin marker when one is present.
    if(extension==='xlsm')await vba.markExport(tempPath,Buffer.from(payload.data),payload.originToken);
    await fs.rename(tempPath, result.filePath);
  } catch (error) {
    await fs.unlink(tempPath).catch(() => {});
    throw error;
  }
  await fileLocations.remember(payload.documentId, result.filePath);
  return result.filePath;
});

for(const [name,handler] of Object.entries({vbaProbe:()=>vba.probe(),vbaStart:payload=>vba.start(payload),vbaStatus:id=>vba.status(id),vbaCommand:payload=>vba.command(payload.id,payload.command),vbaResult:id=>vba.result(id),vbaCancel:id=>vba.cancel(id)})){
  ipcMain.handle('sheet:'+name,async(event,payload)=>{trusted(event);return handler(payload);});
}

ipcMain.on('sheet:set-title', (event, title) => {
  trusted(event);
  if (typeof title === 'string') win.setTitle(title.slice(0, 250).replace(/[\r\n]/g, ' '));
});

ipcMain.on('sheet:close-ready', (event, result) => {
  trusted(event);
  if (!closing || result?.id !== closing.id) return;
  clearTimeout(closing.timer); const mode = closing.mode; closing = null;
  if (result.error) {
    quitting = false;
    void updates.cancelInstall();
    dialog.showErrorBox('자동 복구 저장 실패', '문서를 파일로 저장한 뒤 다시 종료하세요.\n' + String(result.error).slice(0, 500));
    return;
  }
  if (mode === 'reload') { win.webContents.reload(); return; }
  win.destroy();
  if (quitting) app.quit();
});
ipcMain.on('sheet:recovery-ready', event => { trusted(event); recoveryReady = true; });

ipcMain.handle('sheet:write-clipboard', (event, payload) => {
  trusted(event);
  if (!payload || typeof payload.text !== 'string' || typeof payload.html !== 'string' || Buffer.byteLength(payload.text) > 10*1024*1024 || Buffer.byteLength(payload.html) > 10*1024*1024) throw new Error('복사할 표 데이터가 너무 큽니다.');
  clipboard.write({ text: payload.text, html: payload.html });
});
ipcMain.handle('sheet:paste-clipboard', event => { trusted(event); win.webContents.paste(); });

ipcMain.handle('sheet:edit-text', (event, action) => {
  trusted(event);
  const methods = { undo: 'undo', redo: 'redo', cut: 'cut', copy: 'copy', paste: 'paste', 'select-all': 'selectAll' };
  if (!Object.hasOwn(methods, action)) throw new Error('허용되지 않는 편집 요청입니다.');
  win.webContents[methods[action]]();
});

async function createWindow() {
  recoveryReady = false;
  win = new BrowserWindow({
    width: 1540, height: 970, minWidth: 1000, minHeight: 650,
    title: 'Sheet Studio', backgroundColor: '#202020', show: false, autoHideMenuBar: true,
    icon: path.join(__dirname, 'app.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true,
      nodeIntegration: false, sandbox: true, webviewTag: false,
    },
  });
  installMenu();
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  win.webContents.on('will-navigate', event => event.preventDefault());
  win.webContents.session.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
  win.webContents.session.setPermissionCheckHandler(() => false);
  const window = win;
  const flushBefore = mode => {
    if (closing) return;
    const id = randomUUID();
    const timer = setTimeout(() => {
      closing = null; quitting = false;
      void updates.cancelInstall();
      dialog.showErrorBox('자동 복구 저장 대기', '저장을 마치지 못해 창을 유지했습니다. 문서를 파일로 저장한 뒤 다시 종료하세요.');
    }, 60000);
    closing = {id, timer, mode}; window.webContents.send('sheet:before-close', id);
  };
  window.webContents.on('will-prevent-unload', () => { flushBefore('reload'); });
  window.on('close', event => {
    if (window.webContents.isCrashed() || !window.webContents.getURL() || !recoveryReady) return;
    event.preventDefault(); flushBefore('close');
  });
  window.on('blur', () => { if (!window.webContents.isDestroyed()) window.webContents.send('sheet:window-blur'); });
  win.once('ready-to-show', () => win.show());
  await win.loadFile(rendererPath);
  win.on('closed', () => { win = null; });
}

app.whenReady().then(async () => { await createWindow(); await updates.start(); }).catch(error => {
  console.error(error);
  dialog.showErrorBox('Sheet Studio를 시작할 수 없습니다', error.message);
  app.exit(1);
});
app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
app.on('before-quit', () => { quitting = true; });
app.on('window-all-closed', () => { if(process.platform!=='darwin')vba.closeAll().finally(()=>app.quit()); });
