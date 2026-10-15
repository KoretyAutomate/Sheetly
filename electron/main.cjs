const { app, BrowserWindow, Menu, dialog, ipcMain } = require('electron')
const fs = require('fs'), path = require('path'), os = require('os')
let win, current = null

function send(cmd) { win?.webContents.send('menu', cmd) }
function createWindow() {
  win = new BrowserWindow({ width: 1400, height: 900, title: 'Sheetly',
    webPreferences: { preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true } })
  // Menu has no Ctrl+T/W/N-style accelerators that Excel uses differently.
  Menu.setApplicationMenu(Menu.buildFromTemplate([
    { label: 'File', submenu: [
      { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => send('open') },
      { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => send('save') },
      { label: 'Save As…', accelerator: 'F12', click: () => send('saveAs') },
      { type: 'separator' }, { role: 'quit' } ] },
    { label: 'View', submenu: [{ role: 'toggleDevTools' }, { role: 'togglefullscreen' }] } ]))
  const dev = process.env.SHEETLY_DEV_URL
  dev ? win.loadURL(dev) : win.loadFile(path.join(__dirname, '../dist/index.html'))
}
ipcMain.handle('open', async () => {
  const r = await dialog.showOpenDialog(win, { filters: [{ name: 'Excel', extensions: ['xlsx'] }], properties: ['openFile'] })
  if (r.canceled) return null
  const p = r.filePaths[0]; current = p
  return { path: p, name: path.basename(p), data: new Uint8Array(fs.readFileSync(p)) }
})
ipcMain.handle('save', async (_e, data, saveAs) => {
  let p = current
  if (!p || saveAs) {
    const r = await dialog.showSaveDialog(win, { filters: [{ name: 'Excel', extensions: ['xlsx'] }], defaultPath: 'Book1.xlsx' })
    if (r.canceled) return null
    p = r.filePath
  }
  fs.writeFileSync(p, Buffer.from(data)); current = p; return p
})
// INFO("directory"): folder of the open workbook (else home), with a trailing separator like Excel.
ipcMain.on('directory', (e) => { e.returnValue = (current ? path.dirname(current) : os.homedir()) + path.sep })
ipcMain.on('osversion', (e) => { e.returnValue = `${os.type()} ${os.release()}` })
ipcMain.on('current', (_e, p) => { current = p })
app.whenReady().then(createWindow)
app.on('window-all-closed', () => app.quit())
