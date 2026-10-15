const { contextBridge, ipcRenderer } = require('electron')
contextBridge.exposeInMainWorld('sheetly', {
  openDialog: () => ipcRenderer.invoke('open'),
  save: (data, saveAs) => ipcRenderer.invoke('save', data, saveAs),
  setCurrent: (p) => ipcRenderer.send('current', p),
  directory: () => ipcRenderer.sendSync('directory'),
  osversion: () => ipcRenderer.sendSync('osversion'),
  onMenu: (cb) => ipcRenderer.on('menu', (_e, cmd) => cb(cmd)),
})
