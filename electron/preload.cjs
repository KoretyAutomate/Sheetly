const { contextBridge, ipcRenderer } = require('electron')
contextBridge.exposeInMainWorld('sheetly', {
  openDialog: () => ipcRenderer.invoke('open'),
  save: (data, saveAs) => ipcRenderer.invoke('save', data, saveAs),
  setCurrent: (p) => ipcRenderer.send('current', p),
  onMenu: (cb) => ipcRenderer.on('menu', (_e, cmd) => cb(cmd)),
})
