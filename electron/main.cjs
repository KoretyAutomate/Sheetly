const { app, BrowserWindow, Menu, dialog, ipcMain, net } = require('electron')
const fs = require('fs'), path = require('path'), os = require('os')

// One window per workbook, like Excel. Each window remembers its file.
const files = new Map() // webContents.id -> path
const fileOf = (e) => files.get(e.sender.id) ?? null
const pending = new Map() // webContents.id -> path to load once the page asks for it

// Workbooks named on the command line (sheetly.sh book.xlsx, the file manager's "Open with").
const workbookArgs = (argv, cwd) => argv.slice(1).filter((a) => !a.startsWith('-') && /\.xlsx$/i.test(a))
  .map((a) => path.resolve(cwd ?? process.cwd(), a)).filter((p) => fs.existsSync(p))
// Opening a workbook that is already open brings its window forward, as in Excel.
function openWorkbook(p) {
  for (const [id, f] of [...files, ...pending]) if (f === p) { const w = BrowserWindow.getAllWindows().find((x) => x.webContents.id === id); if (w) { if (w.isMinimized()) w.restore(); w.focus(); return w } }
  return createWindow(p)
}

function createWindow(file) {
  const win = new BrowserWindow({ width: 1400, height: 900, title: 'Sheetly',
    webPreferences: { preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true, spellcheck: true } })
  win.webContents.session.setSpellCheckerLanguages(['en-US'])
  const id = win.webContents.id
  if (file) pending.set(id, file)
  win.on('closed', () => { files.delete(id); pending.delete(id) })
  const dev = process.env.SHEETLY_DEV_URL
  dev ? win.loadURL(dev) : win.loadFile(path.join(__dirname, '../dist/index.html'))
  return win
}
const focused = () => BrowserWindow.getFocusedWindow() ?? BrowserWindow.getAllWindows()[0]
const send = (cmd) => focused()?.webContents.send('menu', cmd)
function cycleWindow(step) {
  const all = BrowserWindow.getAllWindows()
  if (all.length < 2) return
  const i = all.indexOf(focused())
  all[(i + step + all.length) % all.length].focus()
}

// The menu shows Excel's File/Window keys, but the keys themselves are handled by the renderer's
// shortcut system (registerAccelerator: false), which calls the 'app-action' IPC below. One code path
// for keyboard and tests, and no double handling.
const actions = {
  new: () => createWindow(),
  open: () => send('open'),
  save: () => send('save'),
  saveAs: () => send('saveAs'),
  close: () => focused()?.close(),
  nextWindow: () => cycleWindow(1),
  prevWindow: () => cycleWindow(-1),
  minimize: () => focused()?.minimize(),
  maximize: () => { const w = focused(); w?.isMaximized() ? w.unmaximize() : w?.maximize() },
}
const item = (label, accelerator, action) => ({ label, accelerator, registerAccelerator: false, click: actions[action] })
Menu.setApplicationMenu(Menu.buildFromTemplate([
  { label: 'File', submenu: [item('New', 'CmdOrCtrl+N', 'new'), item('Open…', 'CmdOrCtrl+O', 'open'), item('Save', 'CmdOrCtrl+S', 'save'),
    item('Save As…', 'F12', 'saveAs'), { type: 'separator' }, item('Close', 'CmdOrCtrl+W', 'close'), { role: 'quit' }] },
  { label: 'Window', submenu: [item('Next Window', 'CmdOrCtrl+Tab', 'nextWindow'), item('Previous Window', 'CmdOrCtrl+Shift+Tab', 'prevWindow'),
    item('Minimize', 'CmdOrCtrl+F9', 'minimize'), item('Maximize / Restore', 'CmdOrCtrl+F10', 'maximize'),
    { role: 'togglefullscreen' }, { role: 'toggleDevTools' }] }]))
// Calls are counted so tests can verify a key reached its action (native dialogs cannot be driven).
global.__appActions = []
ipcMain.handle('app-action', (e, name) => {
  global.__appActions.push(name)
  if (process.env.SHEETLY_TEST_NO_DIALOGS && ['open', 'save', 'saveAs'].includes(name)) return
  BrowserWindow.fromWebContents(e.sender)?.focus()
  actions[name]?.()
})

ipcMain.handle('open', async (e) => {
  const r = await dialog.showOpenDialog(BrowserWindow.fromWebContents(e.sender), { filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }], properties: ['openFile'] })
  if (r.canceled) return null
  // The window takes the file only once it loads it (the page calls 'current'); opening into a new window leaves
  // this window's own file as it was.
  const p = r.filePaths[0]
  return { path: p, name: path.basename(p), data: new Uint8Array(fs.readFileSync(p)) }
})
ipcMain.handle('save', async (e, data, saveAs) => {
  let p = fileOf(e)
  if (!p || saveAs) {
    const r = await dialog.showSaveDialog(BrowserWindow.fromWebContents(e.sender), { filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }], defaultPath: 'Book1.xlsx' })
    if (r.canceled) return null
    p = r.filePath
  }
  fs.writeFileSync(p, Buffer.from(data)); files.set(e.sender.id, p); return p
})
ipcMain.on('current', (e, p) => files.set(e.sender.id, p))
// File > Open over an edited workbook: Excel keeps it and opens the file in a window of its own.
ipcMain.handle('open-in-new-window', (_e, p) => { openWorkbook(p) })
// The page asks once it is ready whether this window was opened for a file.
ipcMain.handle('initial-file', (e) => {
  const p = pending.get(e.sender.id)
  pending.delete(e.sender.id)
  if (!p) return null
  // The window takes the file as its own only once the page has loaded it ('current'): a file that fails to open
  // must never become the target of Save.
  try { return { path: p, name: path.basename(p), data: new Uint8Array(fs.readFileSync(p)) } } catch { return null }
})
// INFO("directory"): folder of the open workbook (else home), with a trailing separator like Excel.
ipcMain.on('directory', (e) => { const p = fileOf(e); e.returnValue = (p ? path.dirname(p) : os.homedir()) + path.sep })
ipcMain.on('osversion', (e) => { e.returnValue = `${os.type()} ${os.release()}` })
// WEBSERVICE(url), following Excel's limits: http(s) only, URL <= 2048 chars, result <= 32767 chars;
// anything else, or a failed request, is #VALUE!.
ipcMain.handle('webservice', async (_e, url) => {
  if (url.length > 2048 || !/^https?:\/\//i.test(url)) return '#VALUE!'
  try {
    const res = await net.fetch(url, { signal: AbortSignal.timeout(30000) })
    if (!res.ok) return '#VALUE!'
    const text = await res.text()
    return text.length > 32767 ? '#VALUE!' : text
  } catch { return '#VALUE!' }
})
// Typing into the grid as a user would (used by Ctrl+' and Ctrl+Shift+").
ipcMain.handle('type-text', (e, text) => e.sender.insertText(text))

// Key chords for editor actions, e.g. [['End', ['shift']], ['Delete', []]].
ipcMain.handle('send-keys', (e, keys) => {
  for (const [keyCode, modifiers] of keys) {
    e.sender.sendInputEvent({ type: 'keyDown', keyCode, modifiers })
    e.sender.sendInputEvent({ type: 'keyUp', keyCode, modifiers })
  }
})
// Print: render the preview HTML in a hidden window and open the system print dialog.
ipcMain.handle('print', async (_e, html, opts) => {
  const w = new BrowserWindow({ show: false })
  await w.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html))
  if (process.env.SHEETLY_TEST_PRINT) { const out = await w.webContents.printToPDF({ landscape: !!opts.landscape, pageSize: opts.pageSize }); w.close(); return out.length }
  return new Promise((resolve) => w.webContents.print({ silent: false, landscape: !!opts.landscape, pageSize: opts.pageSize }, (ok) => { w.close(); resolve(ok) }))
})

// One Sheetly process: a second launch hands its files to this one (each opens in its own window).
// Tests run their own instances, so they skip the lock.
// SHEETLY_USER_DATA gives a test its own profile (and so its own single-instance lock).
if (process.env.SHEETLY_USER_DATA) app.setPath('userData', process.env.SHEETLY_USER_DATA)
const single = process.env.SHEETLY_TEST_NO_DIALOGS || app.requestSingleInstanceLock()
if (!single) app.quit()
else {
  app.on('second-instance', (_e, argv, cwd) => {
    const list = workbookArgs(argv, cwd)
    if (list.length) list.forEach(openWorkbook); else createWindow()
  })
  app.whenReady().then(() => {
    const list = workbookArgs(process.argv)
    if (list.length) list.forEach(openWorkbook); else createWindow()
  })
}
app.on('window-all-closed', () => app.quit())
