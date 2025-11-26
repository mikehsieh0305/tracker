const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('excelAPI', {
  importAllSheets: () => ipcRenderer.invoke('excel:importAllSheets'),
  exportSummary: (payload) => ipcRenderer.invoke('excel:exportSummary', payload)
});
