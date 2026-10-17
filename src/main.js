// KeyTips first: its key listener must be registered before Univer's own (see keytips.js).
import { keyTipState, setKeyTipTree } from './keytips.js'
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import './values.js'
import enUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import './fonts/fonts.css'
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
import { FormulaDataModel, IFunctionService } from '@univerjs/engine-formula'
import { IContextMenuService, IShortcutService } from '@univerjs/ui'
import { ICommandService, IContextService, IUndoRedoService } from '@univerjs/core'
import { IRenderManagerService } from '@univerjs/engine-render'
import { IDrawingManagerService } from '@univerjs/drawing'
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
import { createPivots } from './pivot.js'
import { createSparklines } from './sparklines.js'
import { createRefs } from './refs.js'
import { el, openDialog, setDialogShortcutService } from './ui/dialog.js'
import { xlsxToSnapshot, snapshotToXlsx, formulaResolver } from './xlsx.js'

let univerAPI
const getApi = () => univerAPI
let pivots = null
const extra = [pivotFunctions(getApi), excelFunctions(getApi, { directory: () => window.sheetly.directory(), osversion: () => window.sheetly.osversion(),
  webservice: (url) => window.sheetly.webservice(url), pivotLookup: (...a) => pivots?.lookup(...a) ?? { error: '#REF!' } })]
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
window.__IDrawingManagerService = IDrawingManagerService
window.__IUndoRedoService = IUndoRedoService
const shortcutSvc = univerAPI._injector.get(IShortcutService)
window.__shortcuts = () => shortcutSvc.getAllShortcuts().map((x) => ({ id: x.id, b: x.binding, w: x.win, l: x.linux, d: shortcutSvc.getShortcutDisplay(x) }))
window.__functions = () => [...univerAPI._injector.get(IFunctionService).getExecutors().keys()]
const dialogs = createDialogs(univerAPI)
dialogs.setShortcutService(shortcutSvc)
setDialogShortcutService(shortcutSvc)
const refs = createRefs(univerAPI)
const charts = createCharts(univerAPI, dialogs, refs)
const extras = createExtras(univerAPI)
const outline = createOutline(univerAPI)
const sparklines = createSparklines(univerAPI, refs)
const quickAnalysis = createQuickAnalysis(univerAPI, { dialogs, charts, sparklines })
pivots = createPivots(univerAPI, { dialogs, refs })
const panes = installPanes(univerAPI)
installExcelShortcuts(univerAPI, shortcutSvc, dialogs, charts, extras, outline, quickAnalysis, pivots)
window.__sheetly = { untouched: () => untouched(), refs, pivots, sparklines, dialogs, charts, extras, outline, keyTips: keyTipState, panes, quickAnalysis, xlsx: { xlsxToSnapshot, snapshotToXlsx, save: () => workbookToXlsx(univerAPI.getActiveWorkbook()), load: (data, name = 'Book.xlsx') => loadWorkbook({ data, name, path: null }) } }
const contextService = univerAPI._injector.get(IContextService)
setKeyTipTree(buildKeyTipTree(univerAPI, { dialogs, extras, charts, outline, pivots, sparklines }), { canStart: () => whenSheetEditorFocused(contextService) || !!charts.selectedNow() })
// Every sheet gets Excel's grid size, including ones added later (sheet tab +, Shift+F11, F11), which
// Univer otherwise creates at 1,000 x 20.
const commandService = univerAPI._injector.get(ICommandService)
const fitExcelGrid = () => {
  for (const s of univerAPI.getActiveWorkbook()?.getSheets() ?? []) {
    if (s.getMaxRows() < EXCEL_ROWS) s.setRowCount(EXCEL_ROWS)
    if (s.getMaxColumns() < EXCEL_COLS) s.setColumnCount(EXCEL_COLS)
    // New sheets get Excel's default sizes and font (Univer's are 24 px rows, 88 px columns, Arial).
    const cfg = s.getSheet().getConfig()
    if (!cfg.custom?.sheetlyDefaults) {
      Object.assign(cfg, { defaultRowHeight: EXCEL_SHEET.defaultRowHeight, defaultColumnWidth: EXCEL_SHEET.defaultColumnWidth })
      if (!cfg.defaultStyle) cfg.defaultStyle = { ...EXCEL_SHEET.defaultStyle }
      cfg.custom = { ...(cfg.custom ?? {}), sheetlyDefaults: true }
    }
  }
}
commandService.onCommandExecuted((cmd) => { if (cmd.id === 'sheet.command.insert-sheet' || cmd.id === 'sheet.command.copy-sheet') fitExcelGrid() })
// Objects bound to their own sheet (sparkline cells, PivotTable places, charts over their own sheet's data) are
// pointed back at it whenever sheets or drawings are inserted: a copied sheet (Copy Sheet, and its redo, which
// replays the saved sheet) arrives with references to the original.
let fixPending = false
commandService.onCommandExecuted((cmd) => {
  if (cmd.id !== 'sheet.mutation.insert-sheet' && cmd.id !== 'sheet.mutation.set-drawing-apply') return
  if (fixPending) return
  fixPending = true
  setTimeout(() => { fixPending = false; sparklines.fixLocal(); pivots.fixLocal(); charts.fixLocal() }, 0)
})
// Excel's grid size: 1,048,576 rows x 16,384 columns (A..XFD)
export const EXCEL_ROWS = 1048576, EXCEL_COLS = 16384
// Excel's sheet defaults: Calibri 11, rows 15 pt (20 px), columns 8.43 characters (64 px).
export const EXCEL_SHEET = { defaultRowHeight: 20, defaultColumnWidth: 64, defaultStyle: { ff: 'Calibri', fs: 11 } }
// Workbooks carry their own locale for number/date formatting; Univer's default is zh-CN.
univerAPI.createWorkbook({ id: 'wb1', name: 'Book1', locale: LocaleType.EN_US, sheetOrder: ['sheet1'],
  sheets: { sheet1: { id: 'sheet1', name: 'Sheet1', rowCount: EXCEL_ROWS, columnCount: EXCEL_COLS, ...structuredClone(EXCEL_SHEET), custom: { sheetlyDefaults: true } } } })
// Univer measures text when it lays out the sheet; redraw once Calibri (Carlito) has loaded.
document.fonts?.load('11pt Calibri').then(() => document.fonts.load('bold 11pt Calibri')).then(() => window.dispatchEvent(new Event('resize'))).catch(() => {})

// Excel opens a file in place of Book1 only while Book1 is new and untouched; otherwise in a window of its own.
let fromFile = false
const untouched = () => {
  const id = univerAPI.getActiveWorkbook()?.getId()
  const st = id ? univerAPI._injector.get(IUndoRedoService).getUndoRedoStatus(id) : { undos: 0, redos: 0 }
  return !fromFile && !st.undos && !st.redos
}
async function openFile() {
  const f = await window.sheetly.openDialog()
  if (!f) return
  if (untouched()) await loadWorkbook(f)
  else await window.sheetly.openInNewWindow(f.path)
}
async function loadWorkbook(f) {
  let snap
  try { snap = await xlsxToSnapshot(f.data) } catch (e) {
    // Unreadable (damaged, encrypted, not a workbook): nothing changes, and the window does not take the file.
    console.warn('open', f.name, e)
    openDialog({ title: 'Sheetly', width: 420, body: el('p', { style: { margin: 0 } }, `Sheetly cannot open the file '${f.name}' because the file format or file extension is not valid, or the file is protected with a password.`), cancelLabel: 'Close' })
    window.sheetly.loadFailed?.()
    return false
  }
  // The file is readable: it is this window's from here, before the first calculation (INFO("directory") is its
  // folder).
  window.sheetly.setCurrent(f.path)
  fromFile = true
  univerAPI.disposeUnit(univerAPI.getActiveWorkbook().getId())
  const { sheetlyTables = [], sheetlyLegacyArrays = [], ...data } = snap
  const book = univerAPI.createWorkbook({ ...data, name: f.name, locale: LocaleType.EN_US })
  fitExcelGrid()
  // Recalculate everything, as Excel does for a workbook saved by another program: dynamic arrays spill again
  // (their spilled values are not kept from the file) and cached results are brought up to date.
  univerAPI.getFormula().executeCalculation()
  // Legacy array formulas (Ctrl+Shift+Enter) that turn out to give a single value fill their whole range with the
  // formula, as Excel shows them; array results spill. Opening a file is not an undo step.
  if (sheetlyLegacyArrays.length) {
    await univerAPI.getFormula().onCalculationResultApplied?.(10000).catch(() => {})
    const spills = univerAPI._injector.get(FormulaDataModel).getArrayFormulaRange?.() ?? {}
    for (const a of sheetlyLegacyArrays) {
      if (spills[book.getId()]?.[a.sheetId]?.[a.row]?.[a.col]) continue
      const sheet = book.getSheetBySheetId(a.sheetId)
      for (let r = a.range.startRow; r <= a.range.endRow; r++) for (let c = a.range.startColumn; c <= a.range.endColumn; c++) if (r !== a.row || c !== a.col) sheet?.getRange(r, c).setFormula(a.formula)
    }
  }
  // Tables from the file, through Univer's table API.
  // A table Univer refuses (a name that is also a sheet's, which Excel allows) keeps its definition and is written
  // back as it was.
  for (const t of sheetlyTables) {
    const sheet = book.getSheetBySheetId(t.sheetId)
    let ok = false
    try { ok = (await sheet?.addTable(t.name, t.range, undefined, { showHeader: t.showHeader, showFooter: t.showFooter })) !== false } catch (e) { console.warn('table', t.name, e) }
    if (!ok && sheet) {
      const w = sheet.getSheet(), custom = { ...(w.getCustomMetadata?.() ?? {}) }
      custom.sheetlyTablesKept = [...(custom.sheetlyTablesKept ?? []), { name: t.name, range: t.range, showHeader: t.showHeader, showFooter: t.showFooter, columns: t.columns,
        rangeRef: refs.track(t.sheetId, t.range, book) ?? undefined }]
      w.setCustomMetadata(custom)
    }
  }
  // Opening a file is not something to undo.
  univerAPI._injector.get(IUndoRedoService).clearUndoRedo(book.getId())
}
// Filled-down formulas are kept by Univer as shared-formula ids; their text comes from the live workbook.
// Formulas that spill are written as Excel dynamic arrays over their spill range.
const arrayRangeOf = (wb) => {
  const model = univerAPI._injector.get(FormulaDataModel).getArrayFormulaRange?.() ?? {}
  return (sheetId, r, c) => model[wb.getId()]?.[sheetId]?.[r]?.[c] ?? null
}
const workbookToXlsx = (wb) => snapshotToXlsx(wb.save(), { formulaOf: formulaResolver(wb), arrayRangeOf: arrayRangeOf(wb) })
async function saveFile(saveAs) {
  const wb = univerAPI.getActiveWorkbook()
  const data = await workbookToXlsx(wb)
  await window.sheetly.save(data, saveAs)
}
// A window opened for a file (command line, file manager, second launch) loads it at start.
window.sheetly.initialFile?.().then((f) => { if (f) loadWorkbook(f) })
window.sheetly.onMenu((cmd) => ({ open: openFile, save: () => saveFile(false), saveAs: () => saveFile(true) })[cmd]?.())
