import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import enUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import findReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US'
import '@univerjs/preset-sheets-find-replace/lib/index.css'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import filterEnUS from '@univerjs/preset-sheets-filter/locales/en-US'
import '@univerjs/preset-sheets-filter/lib/index.css'
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link'
import hyperLinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US'
import '@univerjs/preset-sheets-hyper-link/lib/index.css'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import noteEnUS from '@univerjs/preset-sheets-note/locales/en-US'
import '@univerjs/preset-sheets-note/lib/index.css'
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import sortEnUS from '@univerjs/preset-sheets-sort/locales/en-US'
import '@univerjs/preset-sheets-sort/lib/index.css'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import dataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import conditionalFormattingEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import drawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import tableEnUS from '@univerjs/preset-sheets-table/locales/en-US'
import '@univerjs/preset-sheets-table/lib/index.css'
import { UniverSheetsThreadCommentPreset } from '@univerjs/preset-sheets-thread-comment'
import threadCommentEnUS from '@univerjs/preset-sheets-thread-comment/locales/en-US'
import '@univerjs/preset-sheets-thread-comment/lib/index.css'
import { IFunctionService } from '@univerjs/engine-formula'
import { IShortcutService } from '@univerjs/ui'
import { installExcelFunctions } from './excel-functions.js'
import { installExcelShortcuts } from './excel-shortcuts.js'
import { xlsxToSnapshot, snapshotToXlsx } from './xlsx.js'

const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: { [LocaleType.EN_US]: mergeLocales(enUS, findReplaceEnUS, filterEnUS, hyperLinkEnUS, noteEnUS, sortEnUS, dataValidationEnUS, conditionalFormattingEnUS, drawingEnUS, tableEnUS, threadCommentEnUS) },
  presets: [
    UniverSheetsCorePreset({ container: 'app' }),
    UniverSheetsFindReplacePreset(),
    UniverSheetsFilterPreset(),
    UniverSheetsHyperLinkPreset(),
    UniverSheetsNotePreset(),
    UniverSheetsSortPreset(),
    UniverSheetsDataValidationPreset(),
    UniverSheetsConditionalFormattingPreset(),
    UniverSheetsDrawingPreset(),
    UniverSheetsTablePreset(),
    UniverSheetsThreadCommentPreset(),
  ],
})
window.univerAPI = univerAPI
const shortcutSvc = univerAPI._injector.get(IShortcutService)
window.__shortcuts = () => shortcutSvc.getAllShortcuts().map((x) => ({ id: x.id, b: x.binding, w: x.win, l: x.linux, d: shortcutSvc.getShortcutDisplay(x) }))
window.__functions = () => [...univerAPI._injector.get(IFunctionService).getExecutors().keys()]
installExcelShortcuts(univerAPI, shortcutSvc)
installExcelFunctions(univerAPI, { directory: () => window.sheetly.directory(), osversion: () => window.sheetly.osversion() })
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
