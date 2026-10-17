const { contextBridge, ipcRenderer, webFrame } = require('electron')
contextBridge.exposeInMainWorld('sheetly', {
  openDialog: () => ipcRenderer.invoke('open'),
  initialFile: () => ipcRenderer.invoke('initial-file'),
  openInNewWindow: (p) => ipcRenderer.invoke('open-in-new-window', p),
  save: (data, saveAs) => ipcRenderer.invoke('save', data, saveAs),
  setCurrent: (p) => ipcRenderer.send('current', p),
  loadFailed: () => ipcRenderer.send('load-failed'),
  directory: () => ipcRenderer.sendSync('directory'),
  osversion: () => ipcRenderer.sendSync('osversion'),
  webservice: (url) => ipcRenderer.invoke('webservice', url),
  typeText: (text) => ipcRenderer.invoke('type-text', text),
  sendKeys: (keys) => ipcRenderer.invoke('send-keys', keys),
  print: (html, opts) => ipcRenderer.invoke('print', html, opts),
  appAction: (name) => ipcRenderer.invoke('app-action', name),
  // F7 Spelling uses Chromium's spellchecker (en-US).
  spell: { isMisspelled: (w) => webFrame.isWordMisspelled(w), suggestions: (w) => webFrame.getWordSuggestions(w) },
  onMenu: (cb) => ipcRenderer.on('menu', (_e, cmd) => cb(cmd)),
})
