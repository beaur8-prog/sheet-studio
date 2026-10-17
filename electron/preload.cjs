'use strict';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('desktop', Object.freeze({
  openFile: () => ipcRenderer.invoke('sheet:open-file'),
  saveFile: payload => ipcRenderer.invoke('sheet:save-file', payload),
  setTitle: title => ipcRenderer.send('sheet:set-title', String(title)),
}));
