// KeyTips first: its key listener must be registered before Univer's own (see keytips.js).
import { keyTipState, setKeyTipTree } from './keytips.js'
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
import { IContextMenuService, IShortcutService } from '@univerjs/ui'
import { ICommandService, IContextService } from '@univerjs/core'
import { IRenderManagerService } from '@univerjs/engine-render'
import { HeaderUnhideRenderController, SheetSkeletonManagerService, whenSheetEditorFocused } from '@univerjs/sheets-ui'
import { buildKeyTipTree } from './keytips-tree.js'
import { pivotFunctions } from './groupby.js'
import { excelFunctions } from './excel-functions.js'
import { installExcelShortcuts } from './excel-shortcuts.js'
import { createDialogs } from './excel-dialogs.js'
import { createCharts } from './charts.js'
import { createExtras } from './excel-extras.js'
import { createOutline } from './outline.js'
import { installPanes } from './panes.js'
import { createQuickAnalysis } from './quick-analysis.js'
import { setDialogShortcutService } from './ui/dialog.js'
import { xlsxToSnapshot, snapshotToXlsx, formulaResolver } from './xlsx.js'

let univerAPI
const getApi = () => univerAPI
const extra = [pivotFunctions(getApi), excelFunctions(getApi, { directory: () => window.sheetly.directory(), osversion: () => window.sheetly.osversion(),
  webservice: (url) => window.sheetly.webservice(url) })]
;({ univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: { [LocaleType.EN_US]: mergeLocales(enUS, findReplaceEnUS, filterEnUS, hyperLinkEnUS, noteEnUS, sortEnUS, dataValidationEnUS, conditionalFormattingEnUS, drawingEnUS, tableEnUS, threadCommentEnUS) },
  presets: [
    UniverSheetsCorePreset({ container: 'app', formula: { function: extra.flatMap((x) => x.functions), description: extra.flatMap((x) => x.descriptions) } }),
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
}))
window.univerAPI = univerAPI
window.__ICommandService = ICommandService
window.__IShortcutService = IShortcutService
window.__IContextMenuService = IContextMenuService
window.__IRenderManagerService = IRenderManagerService
window.__SheetSkeletonManagerService = SheetSkeletonManagerService
window.__HeaderUnhideRenderController = HeaderUnhideRenderController
const shortcutSvc = univerAPI._injector.get(IShortcutService)
window.__shortcuts = () => shortcutSvc.getAllShortcuts().map((x) => ({ id: x.id, b: x.binding, w: x.win, l: x.linux, d: shortcutSvc.getShortcutDisplay(x) }))
window.__functions = () => [...univerAPI._injector.get(IFunctionService).getExecutors().keys()]
const dialogs = createDialogs(univerAPI)
dialogs.setShortcutService(shortcutSvc)
setDialogShortcutService(shortcutSvc)
const charts = createCharts(univerAPI, dialogs)
const extras = createExtras(univerAPI)
const outline = createOutline(univerAPI)
const quickAnalysis = createQuickAnalysis(univerAPI, { dialogs, charts })
const panes = installPanes(univerAPI)
installExcelShortcuts(univerAPI, shortcutSvc, dialogs, charts, extras, outline, quickAnalysis)
window.__sheetly = { dialogs, charts, extras, outline, keyTips: keyTipState, panes, quickAnalysis, xlsx: { xlsxToSnapshot, snapshotToXlsx, save: () => workbookToXlsx(univerAPI.getActiveWorkbook()) } }
const contextService = univerAPI._injector.get(IContextService)
setKeyTipTree(buildKeyTipTree(univerAPI, { dialogs, extras, charts, outline }), { canStart: () => whenSheetEditorFocused(contextService) })
// Every sheet gets Excel's grid size, including ones added later (sheet tab +, Shift+F11, F11), which
// Univer otherwise creates at 1,000 x 20.
const commandService = univerAPI._injector.get(ICommandService)
const fitExcelGrid = () => {
  for (const s of univerAPI.getActiveWorkbook()?.getSheets() ?? []) {
    if (s.getMaxRows() < EXCEL_ROWS) s.setRowCount(EXCEL_ROWS)
    if (s.getMaxColumns() < EXCEL_COLS) s.setColumnCount(EXCEL_COLS)
  }
}
commandService.onCommandExecuted((cmd) => { if (cmd.id === 'sheet.command.insert-sheet' || cmd.id === 'sheet.command.copy-sheet') fitExcelGrid() })
// Excel's grid size: 1,048,576 rows x 16,384 columns (A..XFD)
export const EXCEL_ROWS = 1048576, EXCEL_COLS = 16384
// Workbooks carry their own locale for number/date formatting; Univer's default is zh-CN.
univerAPI.createWorkbook({ id: 'wb1', name: 'Book1', locale: LocaleType.EN_US, sheetOrder: ['sheet1'],
  sheets: { sheet1: { id: 'sheet1', name: 'Sheet1', rowCount: EXCEL_ROWS, columnCount: EXCEL_COLS } } })

async function openFile() {
  const f = await window.sheetly.openDialog()
  if (!f) return
  const snap = await xlsxToSnapshot(f.data)
  univerAPI.disposeUnit(univerAPI.getActiveWorkbook().getId())
  univerAPI.createWorkbook({ ...snap, name: f.name, locale: LocaleType.EN_US })
  window.sheetly.setCurrent(f.path)
}
// Filled-down formulas are kept by Univer as shared-formula ids; their text comes from the live workbook.
const workbookToXlsx = (wb) => snapshotToXlsx(wb.save(), { formulaOf: formulaResolver(wb) })
async function saveFile(saveAs) {
  const wb = univerAPI.getActiveWorkbook()
  const data = await workbookToXlsx(wb)
  await window.sheetly.save(data, saveAs)
}
window.sheetly.onMenu((cmd) => ({ open: openFile, save: () => saveFile(false), saveAs: () => saveFile(true) })[cmd]?.())
