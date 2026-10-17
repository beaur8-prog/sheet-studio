'use strict';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('desktop', Object.freeze({
  openFile: () => ipcRenderer.invoke('sheet:open-file'),
  saveFile: payload => ipcRenderer.invoke('sheet:save-file', payload),
  setTitle: title => ipcRenderer.send('sheet:set-title', String(title)),
  vbaProbe: () => ipcRenderer.invoke('sheet:vbaProbe'),
  vbaStart: payload => ipcRenderer.invoke('sheet:vbaStart', payload),
  vbaStatus: id => ipcRenderer.invoke('sheet:vbaStatus', id),
  vbaCommand: (id,command) => ipcRenderer.invoke('sheet:vbaCommand', {id,command}),
  vbaResult: id => ipcRenderer.invoke('sheet:vbaResult', id),
  vbaCancel: id => ipcRenderer.invoke('sheet:vbaCancel', id),
}));
