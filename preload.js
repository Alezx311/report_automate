const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  selectPSTFile: () => ipcRenderer.invoke('select-pst-file'),
  parsePST: options => ipcRenderer.invoke('parse-pst', options),
  exportCSV: issues => ipcRenderer.invoke('export-csv', issues),
})
