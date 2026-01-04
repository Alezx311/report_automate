const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  selectPSTFile: () => ipcRenderer.invoke('select-pst-file'),
  loadConfig: () => ipcRenderer.invoke('load-config'),
  parsePST: options => ipcRenderer.invoke('parse-pst', options),
  exportCSV: issues => ipcRenderer.invoke('export-csv', issues),
  // IMAP functions
  testIMAPConnection: credentials => ipcRenderer.invoke('test-imap-connection', credentials),
  getIMAPFolders: credentials => ipcRenderer.invoke('get-imap-folders', credentials),
  parseIMAP: options => ipcRenderer.invoke('parse-imap', options),
})
