import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import enUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import { IFunctionService } from '@univerjs/engine-formula'
import { IShortcutService } from '@univerjs/ui'
import { installExcelShortcuts } from './excel-shortcuts.js'
import { xlsxToSnapshot, snapshotToXlsx } from './xlsx.js'

const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: { [LocaleType.EN_US]: mergeLocales(enUS) },
  presets: [UniverSheetsCorePreset({ container: 'app' })],
})
window.univerAPI = univerAPI
const shortcutSvc = univerAPI._injector.get(IShortcutService)
window.__shortcuts = () => shortcutSvc.getAllShortcuts().map((x) => ({ id: x.id, b: x.binding, w: x.win, l: x.linux, d: shortcutSvc.getShortcutDisplay(x) }))
window.__functions = () => [...univerAPI._injector.get(IFunctionService).getExecutors().keys()]
installExcelShortcuts(univerAPI, shortcutSvc)
// Excel's grid size: 1,048,576 rows x 16,384 columns (A..XFD)
export const EXCEL_ROWS = 1048576, EXCEL_COLS = 16384
univerAPI.createWorkbook({ id: 'wb1', name: 'Book1', sheetOrder: ['sheet1'],
  sheets: { sheet1: { id: 'sheet1', name: 'Sheet1', rowCount: EXCEL_ROWS, columnCount: EXCEL_COLS } } })

async function openFile() {
  const f = await window.sheetly.openDialog()
  if (!f) return
  const snap = await xlsxToSnapshot(f.data)
  univerAPI.disposeUnit(univerAPI.getActiveWorkbook().getId())
  univerAPI.createWorkbook({ ...snap, name: f.name })
  window.sheetly.setCurrent(f.path)
}
async function saveFile(saveAs) {
  const wb = univerAPI.getActiveWorkbook()
  const data = await snapshotToXlsx(wb.save())
  await window.sheetly.save(data, saveAs)
}
window.sheetly.onMenu((cmd) => ({ open: openFile, save: () => saveFile(false), saveAs: () => saveFile(true) })[cmd]?.())
