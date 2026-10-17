'use strict';
const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const fs = require('node:fs/promises');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const vba = require('./vba.cjs');
const { applicationMenu } = require('./menu.cjs');

const MAX_FILE_BYTES = 100 * 1024 * 1024;
const filters = [
  { name: '스프레드시트', extensions: ['xlsx', 'xlsm', 'csv', 'tsv', 'json'] },
  { name: 'Excel 통합 문서', extensions: ['xlsx', 'xlsm'] },
  { name: 'CSV', extensions: ['csv'] },
  { name: 'Sheet Studio', extensions: ['sheet.json', 'json'] },
];
let win;
const rendererPath = path.join(__dirname, '../dist/index.html');
const rendererURL = pathToFileURL(rendererPath).href;
app.setName('Sheet Studio');
if (process.env.SHEET_STUDIO_DATA) app.setPath('userData', path.resolve(process.env.SHEET_STUDIO_DATA));

function trusted(event) {
  if (event.sender !== win?.webContents || event.senderFrame !== win.webContents.mainFrame ||
      event.senderFrame.url.split('#')[0] !== rendererURL) throw new Error('허용되지 않는 앱 요청입니다.');
}

ipcMain.handle('sheet:open-file', async event => {
  trusted(event);
  const result = await dialog.showOpenDialog(win, { title: '통합 문서 열기', properties: ['openFile'], filters });
  if (result.canceled || !result.filePaths.length) return null;
  const file = result.filePaths[0];
  if ((await fs.stat(file)).size > MAX_FILE_BYTES) throw new Error('100 MB 이하의 통합 문서를 선택하세요.');
  const data=await fs.readFile(file);
  const originToken=/\.xlsm$/i.test(file)?await vba.rememberFile(file,data):undefined;
  return { name: path.basename(file), data: Array.from(data), ...(originToken?{originToken}:{}) };
});

ipcMain.handle('sheet:save-file', async (event, payload) => {
  trusted(event);
  if (!payload || typeof payload.name !== 'string' || !Array.isArray(payload.data) ||
      payload.data.length > MAX_FILE_BYTES || payload.data.some(value => !Number.isInteger(value) || value < 0 || value > 255)) {
    throw new Error('저장할 파일 데이터가 올바르지 않습니다.');
  }
  const basename = path.basename(payload.name).replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_') || '통합 문서.xlsx';
  const extension = basename.toLowerCase().endsWith('.sheet.json') ? 'json' : path.extname(basename).slice(1);
  const result = await dialog.showSaveDialog(win, {
    title: '통합 문서 저장', defaultPath: path.join(app.getPath('documents'), basename),
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
  return result.filePath;
});

for(const [name,handler] of Object.entries({vbaProbe:()=>vba.probe(),vbaStart:payload=>vba.start(payload),vbaStatus:id=>vba.status(id),vbaCommand:payload=>vba.command(payload.id,payload.command),vbaResult:id=>vba.result(id),vbaCancel:id=>vba.cancel(id)})){
  ipcMain.handle('sheet:'+name,async(event,payload)=>{trusted(event);return handler(payload);});
}

ipcMain.on('sheet:set-title', (event, title) => {
  trusted(event);
  if (typeof title === 'string') win.setTitle(title.slice(0, 250).replace(/[\r\n]/g, ' '));
});

ipcMain.handle('sheet:edit-text', (event, action) => {
  trusted(event);
  const methods = { undo: 'undo', redo: 'redo', cut: 'cut', copy: 'copy', paste: 'paste', 'select-all': 'selectAll' };
  if (!Object.hasOwn(methods, action)) throw new Error('허용되지 않는 편집 요청입니다.');
  win.webContents[methods[action]]();
});

async function createWindow() {
  win = new BrowserWindow({
    width: 1540, height: 970, minWidth: 1000, minHeight: 650,
    title: 'Sheet Studio', backgroundColor: '#202020', show: false, autoHideMenuBar: true,
    icon: path.join(__dirname, 'app.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true,
      nodeIntegration: false, sandbox: true, webviewTag: false,
    },
  });
  const menu = applicationMenu(process.platform, action => win?.webContents.send('sheet:menu-action', action));
  Menu.setApplicationMenu(menu ? Menu.buildFromTemplate(menu) : null);
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  win.webContents.on('will-navigate', event => event.preventDefault());
  win.webContents.session.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
  win.webContents.session.setPermissionCheckHandler(() => false);
  win.once('ready-to-show', () => win.show());
  await win.loadFile(rendererPath);
  win.on('closed', () => { win = null; });
}

app.whenReady().then(createWindow).catch(error => {
  console.error(error);
  dialog.showErrorBox('Sheet Studio를 시작할 수 없습니다', error.message);
  app.exit(1);
});
app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
app.on('window-all-closed', () => { if(process.platform!=='darwin')vba.closeAll().finally(()=>app.quit()); });
