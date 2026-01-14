const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  selectPSTFile: () => ipcRenderer.invoke('select-pst-file'),
  loadConfig: () => ipcRenderer.invoke('load-config'),
  parsePST: options => ipcRenderer.invoke('parse-pst', options),
  exportCSV: issues => ipcRenderer.invoke('export-csv', issues),
  // Graph API functions
  testGraphConnection: credentials => ipcRenderer.invoke('test-graph-connection', credentials),
  getGraphFolders: credentials => ipcRenderer.invoke('get-graph-folders', credentials),
  parseGraph: options => ipcRenderer.invoke('parse-graph', options),
  // Jira functions
  connectJira: config => ipcRenderer.invoke('connect-jira', config),
  fetchJiraIssues: options => ipcRenderer.invoke('fetch-jira-issues', options),
  exportToJira: options => ipcRenderer.invoke('export-to-jira', options),
  onJiraProgress: callback => ipcRenderer.on('jira-export-progress', (event, data) => callback(data)),
  // Cache functions
  saveToCache: options => ipcRenderer.invoke('save-to-cache', options),
  listCacheFiles: () => ipcRenderer.invoke('list-cache-files'),
  loadFromCache: fileName => ipcRenderer.invoke('load-from-cache', fileName),
  deleteCacheFile: fileName => ipcRenderer.invoke('delete-cache-file', fileName),
  clearAllCache: () => ipcRenderer.invoke('clear-all-cache'),
  getCacheStats: () => ipcRenderer.invoke('get-cache-stats'),
})
