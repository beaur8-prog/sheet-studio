'use strict';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('desktop', Object.freeze({
  openFile: () => ipcRenderer.invoke('sheet:open-file'),
  saveFile: payload => ipcRenderer.invoke('sheet:save-file', payload),
  setTitle: title => ipcRenderer.send('sheet:set-title', String(title)),
  editText: action => ipcRenderer.invoke('sheet:edit-text', action),
  onMenuAction: listener => {
    if (typeof listener !== 'function') throw new TypeError('A menu listener is required.');
    const allowed = new Set(['new', 'open', 'save', 'undo', 'redo', 'cut', 'copy', 'paste', 'select-all', 'find']);
    const handler = (_event, action) => { if (allowed.has(action)) listener(action); };
    ipcRenderer.on('sheet:menu-action', handler);
    return () => ipcRenderer.removeListener('sheet:menu-action', handler);
  },
  vbaProbe: () => ipcRenderer.invoke('sheet:vbaProbe'),
  vbaStart: payload => ipcRenderer.invoke('sheet:vbaStart', payload),
  vbaStatus: id => ipcRenderer.invoke('sheet:vbaStatus', id),
  vbaCommand: (id,command) => ipcRenderer.invoke('sheet:vbaCommand', {id,command}),
  vbaResult: id => ipcRenderer.invoke('sheet:vbaResult', id),
  vbaCancel: id => ipcRenderer.invoke('sheet:vbaCancel', id),
}));
