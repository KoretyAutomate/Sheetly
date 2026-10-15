// Excel keyboard shortcuts on top of Univer: adds the ones Univer lacks and overrides the ones whose
// Univer meaning differs from Excel (e.g. Ctrl+0 zoom -> hide columns, Ctrl+- zoom -> delete cells,
// Shift+Backspace clear -> collapse selection, Ctrl+Shift+0 hide -> unhide columns).
import { BorderStyleTypes, BorderType, CommandType, ICommandService } from '@univerjs/core'
import { whenSheetEditorActivated, whenSheetEditorFocused } from '@univerjs/sheets-ui'
import { KeyCode, MetaKeys } from '@univerjs/ui'

const { CTRL_COMMAND: CTRL, SHIFT, ALT } = MetaKeys
const CONTEXT_MENU_KEY = 93, PAGE_UP = 33, PAGE_DOWN = 34, SEMICOLON = 186, QUOTE = 222, BACKQUOTE = 192, BRACKET_L = 219, BRACKET_R = 221
const EQUAL = KeyCode.EQUAL, MINUS = KeyCode.MINUS, PERIOD = KeyCode.PERIOD, BACK_SLASH = KeyCode.BACK_SLASH
const DIGIT = (n) => KeyCode.Digit0 + n
const excelSerial = (d) => Math.floor((d - new Date(1899, 11, 30)) / 864e5)
const dayFraction = (d) => (d - new Date(d.getFullYear(), d.getMonth(), d.getDate())) / 864e5
const PRIORITY = 200

export function installExcelShortcuts(api, shortcuts, dialogs, charts, extras) {
  const commands = api._injector.get(ICommandService)
  const wb = () => api.getActiveWorkbook()
  const sheet = () => wb().getActiveSheet()
  const cur = () => sheet().getActiveRange()
  const sel = () => cur().getRange()
  const activeCell = () => { const c = sheet().getSelection?.()?.getCurrentCell?.(); return c ? [c.actualRow ?? c.startRow, c.actualColumn ?? c.startColumn] : [sel().startRow, sel().startColumn] }
  const goto = (r, c) => { sheet().getRange(Math.max(0, r), Math.max(0, c)).activate(); return true }
  const selectRect = (r0, c0, r1, c1) => {
    const g = { startRow: Math.min(r0, r1), endRow: Math.max(r0, r1), startColumn: Math.min(c0, c1), endColumn: Math.max(c0, c1) }
    dialogs.select([g])
    // keep the original cell as the active cell, like Excel
    return true
  }
  const visible = () => sheet().getVisibleRange?.() ?? { startRow: 0, endRow: 20, startColumn: 0, endColumn: 9 }
  const pageRows = () => { const v = visible(); return Math.max(1, v.endRow - v.startRow) }
  const pageCols = () => { const v = visible(); return Math.max(1, v.endColumn - v.startColumn) }
  const fmt = (p) => { cur().setNumberFormat(p); return true }
  const lastCell = () => { const u = dialogs.usedRange(); return [u.endRow, u.endColumn] }
  const app = (name) => { window.sheetly?.appAction?.(name); return true }

  // Modes: F8 Extend Selection, End mode.
  const mode = { extend: false, end: false }
  const status = document.createElement('div')
  Object.assign(status.style, { position: 'fixed', left: '8px', bottom: '4px', font: '12px system-ui, sans-serif', color: '#605e5c', pointerEvents: 'none', zIndex: 5 })
  status.setAttribute('aria-live', 'polite')
  document.body.append(status)
  const showMode = () => { status.textContent = mode.extend ? 'Extend Selection' : mode.end ? 'End Mode' : '' }
  const inExtend = (ctx) => mode.extend && whenSheetEditorFocused(ctx)
  const inEnd = (ctx) => mode.end && whenSheetEditorFocused(ctx)

  const defs = [
    // ---- navigation
    ['go-top-left', CTRL | KeyCode.HOME, () => goto(0, 0)],
    ['go-last-cell', CTRL | KeyCode.END, () => goto(...lastCell())],
    ['go-row-start', KeyCode.HOME, () => goto(activeCell()[0], 0)],
    ['page-down', PAGE_DOWN, () => { const [r, c] = activeCell(); return goto(r + pageRows(), c) }],
    ['page-up', PAGE_UP, () => { const [r, c] = activeCell(); return goto(r - pageRows(), c) }],
    ['page-right', ALT | PAGE_DOWN, () => { const [r, c] = activeCell(); return goto(r, c + pageCols()) }],
    ['page-left', ALT | PAGE_UP, () => { const [r, c] = activeCell(); return goto(r, c - pageCols()) }],
    ['next-sheet', CTRL | PAGE_DOWN, () => switchSheet(1)],
    ['prev-sheet', CTRL | PAGE_UP, () => switchSheet(-1)],
    ['scroll-to-active', CTRL | KeyCode.BACKSPACE, () => { const [r, c] = activeCell(); sheet().scrollToCell?.(r, c); return true }],
    ['next-corner', CTRL | PERIOD, () => nextCorner()],
    ['goto', KeyCode.F5, () => dialogs.goTo()],
    ['goto-g', CTRL | KeyCode.G, () => dialogs.goTo()],
    // ---- selection
    ['select-column', CTRL | KeyCode.SPACE, () => { const s = sheet(), g = sel(); dialogs.select([{ startRow: 0, endRow: s.getMaxRows() - 1, startColumn: g.startColumn, endColumn: g.endColumn }]); return true }],
    ['select-row', SHIFT | KeyCode.SPACE, () => { const s = sheet(), g = sel(); dialogs.select([{ startRow: g.startRow, endRow: g.endRow, startColumn: 0, endColumn: s.getMaxColumns() - 1 }]); return true }],
    ['select-all', CTRL | SHIFT | KeyCode.SPACE, () => { const s = sheet(); dialogs.select([{ startRow: 0, endRow: s.getMaxRows() - 1, startColumn: 0, endColumn: s.getMaxColumns() - 1 }]); return true }],
    ['select-to-last', CTRL | SHIFT | KeyCode.END, () => { const [r, c] = activeCell(), [lr, lc] = lastCell(); return selectRect(r, c, lr, lc) }],
    ['select-to-start', CTRL | SHIFT | KeyCode.HOME, () => { const [r, c] = activeCell(); return selectRect(0, 0, r, c) }],
    ['select-row-start', SHIFT | KeyCode.HOME, () => { const [r, c] = activeCell(); return selectRect(r, 0, r, c) }],
    ['current-region', CTRL | SHIFT | DIGIT(8), () => { const [r, c] = activeCell(); dialogs.select([dialogs.currentRegion(r, c)]); return true }],
    ['collapse-selection', SHIFT | KeyCode.BACKSPACE, () => { const [r, c] = activeCell(); return goto(r, c) }],
    ['extend-mode', KeyCode.F8, () => { mode.extend = !mode.extend; mode.end = false; showMode(); return true }],
    ['end-mode', KeyCode.END, () => { mode.end = !mode.end; mode.extend = false; showMode(); return true }],
    ['select-notes', CTRL | SHIFT | KeyCode.O, () => dialogs.runGoToSpecial('notes') || true],
    ['select-visible', ALT | SEMICOLON, () => dialogs.runGoToSpecial('visible') || true],
    ['row-differences', CTRL | BACK_SLASH, () => dialogs.runGoToSpecial('rowdiff') || true],
    ['column-differences', CTRL | SHIFT | BACK_SLASH, () => dialogs.runGoToSpecial('coldiff') || true],
    ['precedents', CTRL | BRACKET_L, () => dialogs.runGoToSpecial('precedents') || true],
    ['dependents', CTRL | BRACKET_R, () => dialogs.runGoToSpecial('dependents') || true],
    // ---- editing
    ['insert-date', CTRL | SEMICOLON, () => { const r = cur(); r.setValue(excelSerial(new Date())); r.setNumberFormat('m/d/yyyy'); return true }],
    ['insert-time', CTRL | SHIFT | SEMICOLON, () => { const r = cur(); r.setValue(dayFraction(new Date())); r.setNumberFormat('h:mm AM/PM'); return true }],
    ['copy-formula-above', CTRL | QUOTE, () => copyAbove(true)],
    ['copy-value-above', CTRL | SHIFT | QUOTE, () => copyAbove(false)],
    ['autosum', ALT | EQUAL, () => autosum(sheet(), cur())],
    ['insert-cells', CTRL | SHIFT | EQUAL, () => dialogs.insertCells()],
    ['delete-cells', CTRL | MINUS, () => dialogs.deleteCells()],
    ['paste-special', CTRL | ALT | KeyCode.V, () => dialogs.pasteSpecial()],
    ['hide-rows', CTRL | DIGIT(9), () => { const g = sel(); sheet().hideRows(g.startRow, g.endRow - g.startRow + 1); return true }],
    ['unhide-rows', CTRL | SHIFT | DIGIT(9), () => { const g = sel(); sheet().showRows(g.startRow, g.endRow - g.startRow + 1); return true }],
    ['hide-columns', CTRL | DIGIT(0), () => { const g = sel(); sheet().hideColumns(g.startColumn, g.endColumn - g.startColumn + 1); return true }],
    ['unhide-columns', CTRL | SHIFT | DIGIT(0), () => { const g = sel(); sheet().showColumns(g.startColumn, g.endColumn - g.startColumn + 1); return true }],
    ['insert-note', SHIFT | KeyCode.F2, () => commands.executeCommand('sheet.operation.add-note-popup')],
    ['create-names', CTRL | SHIFT | KeyCode.F3, () => dialogs.createNames()],
    ['name-manager', CTRL | KeyCode.F3, () => commands.executeCommand('sidebar.operation.defined-name', { value: 'open' })],
    // Ctrl+K (Insert Hyperlink) is Univer's own binding from the hyper-link preset.
    ['table', CTRL | KeyCode.T, () => dialogs.createTable()],
    ['table-l', CTRL | KeyCode.L, () => dialogs.createTable()],
    ['chart-sheet', KeyCode.F11, () => charts.chartSheet()],
    ['chart-embedded', ALT | KeyCode.F1, () => charts.embeddedChart()],
    ['new-sheet', SHIFT | KeyCode.F11, () => newSheet()],
    ['new-sheet-alt', ALT | SHIFT | KeyCode.F1, () => newSheet()],
    ['print', CTRL | KeyCode.P, () => dialogs.printPreview()],
    ['print-preview', CTRL | KeyCode.F2, () => dialogs.printPreview()],
    ['help', KeyCode.F1, () => dialogs.help(SHORTCUT_DOCS)],
    ['calculate', KeyCode.F9, () => recalc()],
    ['calculate-sheet', SHIFT | KeyCode.F9, () => recalc()],
    ['calculate-all', CTRL | ALT | KeyCode.F9, () => recalc()],
    ['pick-from-list', ALT | KeyCode.ARROW_DOWN, () => pickFromList()],
    // ---- formatting
    ['strikethrough', CTRL | DIGIT(5), () => commands.executeCommand('sheet.command.set-range-stroke')],
    ['format-cells', CTRL | DIGIT(1), () => dialogs.formatCells('Number')],
    ['format-font', CTRL | SHIFT | KeyCode.F, () => dialogs.formatCells('Font')],
    ['format-font-p', CTRL | SHIFT | KeyCode.P, () => dialogs.formatCells('Font')],
    ['outline-border', CTRL | SHIFT | DIGIT(7), () => { cur().setBorder(BorderType.OUTSIDE, BorderStyleTypes.THIN, '#000000'); return true }],
    ['remove-borders', CTRL | SHIFT | MINUS, () => { cur().setBorder(BorderType.NONE, BorderStyleTypes.NONE, '#000000'); return true }],
    ['fmt-general', CTRL | SHIFT | BACKQUOTE, () => fmt('General')],
    ['fmt-number', CTRL | SHIFT | DIGIT(1), () => fmt('#,##0.00')],
    ['fmt-time', CTRL | SHIFT | DIGIT(2), () => fmt('h:mm AM/PM')],
    ['fmt-date', CTRL | SHIFT | DIGIT(3), () => fmt('d-mmm-yy')],
    ['fmt-currency', CTRL | SHIFT | DIGIT(4), () => fmt('$#,##0.00_);($#,##0.00)')],
    ['fmt-percent', CTRL | SHIFT | DIGIT(5), () => fmt('0%')],
    ['fmt-scientific', CTRL | SHIFT | DIGIT(6), () => fmt('0.00E+00')],
    ['zoom-in', CTRL | ALT | EQUAL, () => zoomBy(0.1)],
    ['show-formulas', CTRL | BACKQUOTE, () => extras.toggleShowFormulas()],
    ['insert-function', SHIFT | KeyCode.F3, () => extras.insertFunction()],
    ['paste-name', KeyCode.F3, () => extras.pasteName()],
    ['flash-fill', CTRL | KeyCode.E, () => extras.flashFill()],
    ['spelling', KeyCode.F7, () => extras.spelling()],
    ['context-menu', SHIFT | KeyCode.F10, () => extras.contextMenu()],
    ['context-menu-key', CONTEXT_MENU_KEY, () => extras.contextMenu()],
    ['toggle-objects', CTRL | DIGIT(6), () => extras.toggleObjects()],
    ['formula-bar', CTRL | SHIFT | KeyCode.U, () => extras.toggleFormulaBar()],
    // ---- workbook / window (actions run in the Electron main process)
    ['file-new', CTRL | KeyCode.N, () => app('new')],
    ['file-open', CTRL | KeyCode.O, () => app('open')],
    ['file-open-f12', CTRL | KeyCode.F12, () => app('open')],
    ['file-save', CTRL | KeyCode.S, () => app('save')],
    ['file-save-f12', SHIFT | KeyCode.F12, () => app('save')],
    ['file-save-as', KeyCode.F12, () => app('saveAs')],
    ['file-close', CTRL | KeyCode.W, () => app('close')],
    ['file-close-f4', CTRL | KeyCode.F4, () => app('close')],
    ['window-next', CTRL | KeyCode.TAB, () => app('nextWindow')],
    ['window-next-f6', CTRL | KeyCode.F6, () => app('nextWindow')],
    ['window-prev', CTRL | SHIFT | KeyCode.TAB, () => app('prevWindow')],
    ['window-prev-f6', CTRL | SHIFT | KeyCode.F6, () => app('prevWindow')],
    ['window-minimize', CTRL | KeyCode.F9, () => app('minimize')],
    ['window-maximize', CTRL | KeyCode.F10, () => app('maximize')],
    ['zoom-out', CTRL | ALT | MINUS, () => zoomBy(-0.1)],
  ]

  function switchSheet(step) {
    const w = wb(), s = w.getSheets().filter((x) => !x.isSheetHidden())
    const i = s.findIndex((x) => x.getSheetId() === sheet().getSheetId())
    w.setActiveSheet(s[(i + step + s.length) % s.length])
    return true
  }
  function nextCorner() {
    const g = sel(), [r, c] = activeCell()
    const corners = [[g.startRow, g.startColumn], [g.startRow, g.endColumn], [g.endRow, g.endColumn], [g.endRow, g.startColumn]]
    const i = corners.findIndex(([y, x]) => y === r && x === c)
    const [nr, nc] = corners[(i + 1) % 4]
    commands.executeCommand('sheet.operation.set-selections', { unitId: wb().getId(), subUnitId: sheet().getSheetId(),
      selections: [{ range: { ...g, rangeType: 0 }, style: null, primary: { startRow: nr, startColumn: nc, endRow: nr, endColumn: nc, actualRow: nr, actualColumn: nc, isMerged: false, isMergedMainCell: false } }] })
    return true
  }
  function newSheet() {
    const w = wb(), names = new Set(w.getSheets().map((s) => s.getSheetName()))
    let n = 1
    while (names.has(`Sheet${n}`)) n++
    const index = w.getSheets().findIndex((s) => s.getSheetId() === sheet().getSheetId())
    const s = w.insertSheet(`Sheet${n}`, { index })
    w.setActiveSheet(s)
    return true
  }
  // Ctrl+' / Ctrl+Shift+": put the formula (or value) from the cell above into the active cell and
  // leave it in edit mode. Done synchronously inside the key event so a following Enter commits it;
  // Esc restores what the cell held before, as in Excel.
  function copyAbove(formula) {
    const [r, c] = activeCell()
    if (r === 0) return true
    const above = sheet().getRange(r - 1, c), range = sheet().getRange(r, c)
    // getCellData() is the live cell object: copy it before the cell is overwritten.
    const original = structuredClone(range.getCellData() ?? null)
    if (formula && above.getFormula()) range.setFormula(above.getFormula())
    else { const v = above.getValue(); if (v == null || v === '') range.clearContent(); else range.setValue(v) }
    commands.executeCommand('sheet.operation.set-cell-edit-visible-f2', { visible: true, eventType: 4, keycode: KeyCode.F2 })
    const sub = commands.onCommandExecuted((cmd) => {
      // Only user-initiated closes carry a keycode (Esc, Enter, Tab, ...); internal closes do not.
      if (cmd.id !== 'sheet.operation.set-cell-edit-visible' || cmd.params?.visible !== false || cmd.params?.keycode == null) return
      sub.dispose()
      if (cmd.params?.keycode === KeyCode.ESC) { if (original) sheet().getRange(r, c).setValues([[original]]); else sheet().getRange(r, c).clearContent() }
    })
    return true
  }
  function recalc() {
    api.getFormula().executeCalculation()
    return true
  }
  function zoomBy(d) {
    const z = sheet().getZoom?.() ?? 1
    sheet().zoom?.(Math.min(4, Math.max(0.1, Math.round((z + d) * 10) / 10)))
    return true
  }
  // Alt+Down: data-validation dropdown if the cell has one, else Excel's "Pick From Drop-down List"
  // of the distinct text entries above and below in the column.
  function pickFromList() {
    const [r, c] = activeCell()
    const s = sheet(), seen = new Set()
    for (let y = r - 1; y >= 0 && s.getRange(y, c).getValue() != null && s.getRange(y, c).getValue() !== ''; y--) seen.add(String(s.getRange(y, c).getValue()))
    for (let y = r + 1; s.getRange(y, c).getValue() != null && s.getRange(y, c).getValue() !== ''; y++) seen.add(String(s.getRange(y, c).getValue()))
    const options = [...seen].filter((v) => Number.isNaN(Number(v))).sort((a, b) => a.localeCompare(b))
    if (!options.length) return true
    const hide = s.getRange(r, c).showDropdown({
      location: { unitId: wb().getId(), subUnitId: s.getSheetId(), workbook: wb().getWorkbook(), worksheet: s.getSheet(), row: r, col: c },
      type: 'list', closeOnOutSide: true,
      props: { options: options.map((v) => ({ label: v, value: v })), onChange: (v) => { if (v[0] != null) s.getRange(r, c).setValue(v[0]); hide?.dispose?.(); return true } },
    })
    return true
  }

  // Register commands + shortcuts.
  for (const [name, binding, handler] of defs) {
    const id = `sheetly.${name}`
    commands.registerCommand({ id, type: CommandType.COMMAND, handler: async () => (await handler()) !== false })
    shortcuts.registerShortcut({ id, binding, priority: PRIORITY, preconditions: (ctx) => !mode.extend && !mode.end && whenSheetEditorFocused(ctx) || (name === 'extend-mode' && mode.extend) || (name === 'end-mode' && mode.end) })
  }

  // Re-bind existing Univer commands for Excel modes by cloning Univer's own shortcut items. Univer
  // registers its sheet shortcuts lazily, so clone whenever the shortcut list changes.
  const cloned = new Set()
  let cloning = false
  const cloneBindings = () => {
    if (cloning) return
    cloning = true
    try {
      for (const item of shortcuts.getAllShortcuts()) {
        const p = item.staticParameters ?? {}
        const key = `${item.id}|${item.binding}|${JSON.stringify(p)}`
        if (cloned.has(key) || String(item.id).startsWith('sheetly.')) continue
        // F8 Extend Selection: plain arrows extend, like Shift+arrows.
        if (item.id === 'sheet.command.expand-selection' && item.binding >= (SHIFT | KeyCode.ARROW_LEFT) && item.binding <= (SHIFT | KeyCode.ARROW_DOWN) && p.jumpOver == null) {
          cloned.add(key)
          shortcuts.registerShortcut({ ...item, binding: item.binding & ~SHIFT, priority: PRIORITY + 10, preconditions: inExtend })
        }
        // End mode: End then arrow = Ctrl+arrow.
        if (item.id === 'sheet.command.move-selection' && p.jumpOver === 1 && (item.binding & ~CTRL) >= KeyCode.ARROW_LEFT && (item.binding & ~CTRL) <= KeyCode.ARROW_DOWN) {
          cloned.add(key)
          const arrow = item.binding & ~CTRL
          const id = `sheetly.end-arrow-${arrow}`
          commands.registerCommand({ id, type: CommandType.COMMAND, handler: async () => { mode.end = false; showMode(); return commands.executeCommand(item.id, p) } })
          shortcuts.registerShortcut({ id, binding: arrow, priority: PRIORITY + 10, preconditions: inEnd })
        }
        // Ctrl+Shift+Enter commits the cell like Enter (legacy array entry; results spill as in Excel 365).
        if (item.id === 'sheet.operation.set-cell-edit-visible' && item.binding === KeyCode.ENTER) {
          cloned.add(key)
          shortcuts.registerShortcut({ ...item, binding: CTRL | SHIFT | KeyCode.ENTER, priority: PRIORITY })
        }
      }
    } finally { cloning = false }
  }
  cloneBindings()
  shortcuts.shortcutChanged$.subscribe(cloneBindings)

  // Leaving the modes: Esc ends both.
  commands.registerCommand({ id: 'sheetly.leave-modes', type: CommandType.COMMAND, handler: () => { mode.extend = mode.end = false; showMode(); return true } })
  shortcuts.registerShortcut({ id: 'sheetly.leave-modes', binding: KeyCode.ESC, priority: PRIORITY + 10, preconditions: () => mode.extend || mode.end })

  // While editing a cell: Shift+F3 Insert Function, F3 Paste Name, Ctrl+Shift+A argument names.
  for (const [name, binding, handler] of [['edit-insert-function', SHIFT | KeyCode.F3, () => extras.insertFunction()],
    ['edit-paste-name', KeyCode.F3, () => extras.pasteName()], ['edit-argument-names', CTRL | SHIFT | KeyCode.A, () => extras.insertArgumentNames()]]) {
    const id = `sheetly.${name}`
    commands.registerCommand({ id, type: CommandType.COMMAND, handler: async () => (await handler()) !== false })
    shortcuts.registerShortcut({ id, binding, priority: PRIORITY, preconditions: whenSheetEditorActivated })
  }

  // In-cell editing: Ctrl+Delete deletes to the end of the line (Excel), not just the next word.
  commands.registerCommand({ id: 'sheetly.delete-to-line-end', type: CommandType.COMMAND, handler: () => { window.sheetly?.sendKeys?.([['End', ['shift']], ['Delete', []]]); return true } })
  shortcuts.registerShortcut({ id: 'sheetly.delete-to-line-end', binding: CTRL | KeyCode.DELETE, priority: PRIORITY, preconditions: whenSheetEditorActivated })
}

function autosum(ws, r) {
  const row = r.getRow(), col = r.getColumn(), num = (a, b) => typeof ws.getRange(a, b).getValue() === 'number'
  let top = row - 1
  while (top >= 0 && num(top, col)) top--
  const target = ws.getRange(row, col)
  if (top < row - 1) { target.setFormula(`=SUM(${ws.getRange(top + 1, col, row - 1 - top, 1).getA1Notation()})`); return true }
  let left = col - 1
  while (left >= 0 && num(row, left)) left--
  if (left < col - 1) { target.setFormula(`=SUM(${ws.getRange(row, left + 1, 1, col - 1 - left).getA1Notation()})`); return true }
  // Nothing to sum: Excel leaves =SUM() in edit mode for the user to pick a range.
  target.setFormula('=SUM()')
  return true
}

/** Shown in the F1 Help pane. */
export const SHORTCUT_DOCS = [
  ['Ctrl+Home', 'Go to cell A1'], ['Ctrl+End', 'Go to the last used cell'], ['Home', 'Go to the start of the row'],
  ['Ctrl+Arrow', 'Go to the edge of the data region'], ['PageUp / PageDown', 'Move one screen up / down'],
  ['Alt+PageUp / Alt+PageDown', 'Move one screen left / right'], ['Ctrl+PageUp / Ctrl+PageDown', 'Previous / next sheet'],
  ['F5 / Ctrl+G', 'Go To'], ['Ctrl+F / Ctrl+H', 'Find / Replace'], ['Ctrl+Backspace', 'Scroll to the active cell'],
  ['Ctrl+.', 'Move to the next corner of the selection'],
  ['Ctrl+A', 'Select the current region, then the whole sheet'], ['Ctrl+Shift+Space', 'Select the whole sheet'],
  ['Ctrl+Space / Shift+Space', 'Select column / row'], ['Shift+Arrow', 'Extend the selection'],
  ['Ctrl+Shift+Arrow', 'Extend to the edge of the data'], ['Ctrl+Shift+Home / End', 'Extend to A1 / the last used cell'],
  ['Ctrl+Shift+*', 'Select the current region'], ['Shift+Backspace', 'Select only the active cell'],
  ['F8', 'Extend Selection mode'], ['End, Arrow', 'End mode: move to the edge of the data'],
  ['Ctrl+Shift+O', 'Select cells with notes'], ['Alt+;', 'Select visible cells only'], ['Ctrl+\\ / Ctrl+Shift+|', 'Row / column differences'],
  ['Ctrl+[ / Ctrl+]', 'Select precedents / dependents'],
  ['F2', 'Edit the active cell'], ['Esc', 'Cancel editing'], ['Enter / Shift+Enter', 'Commit and move down / up'],
  ['Tab / Shift+Tab', 'Commit and move right / left'], ['Alt+Enter', 'New line in a cell'], ['Ctrl+Enter', 'Fill the selection with the entry'],
  ['Ctrl+Shift+Enter', 'Enter as an array formula'], ['F4', 'Toggle absolute reference / repeat last action'],
  ['Ctrl+D / Ctrl+R', 'Fill down / right'], ['Ctrl+; / Ctrl+Shift+;', 'Insert date / time'],
  ["Ctrl+' / Ctrl+Shift+\"", 'Copy formula / value from the cell above'], ['Alt+=', 'AutoSum'],
  ['Ctrl+C / Ctrl+X / Ctrl+V', 'Copy / cut / paste'], ['Ctrl+Alt+V', 'Paste Special'], ['Ctrl+Z / Ctrl+Y', 'Undo / redo'],
  ['Ctrl+Shift++ / Ctrl+-', 'Insert / delete cells'], ['Ctrl+9 / Ctrl+Shift+9', 'Hide / unhide rows'], ['Ctrl+0 / Ctrl+Shift+0', 'Hide / unhide columns'],
  ['Shift+F2', 'Insert or edit a note'], ['Ctrl+F3 / Ctrl+Shift+F3', 'Name Manager / create names from selection'],
  ['Ctrl+K', 'Insert hyperlink'], ['Ctrl+T / Ctrl+L', 'Create table'], ['Ctrl+Shift+L', 'Toggle filter'], ['Alt+Down', 'Open drop-down list'],
  ['F11 / Alt+F1', 'Chart sheet / embedded chart'], ['Shift+F11 / Alt+Shift+F1', 'Insert a worksheet'],
  ['F9 / Shift+F9 / Ctrl+Alt+F9', 'Calculate'], ['Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+5', 'Bold / italic / underline / strikethrough'],
  ['Ctrl+1', 'Format Cells'], ['Ctrl+Shift+F / Ctrl+Shift+P', 'Format Cells, Font tab'],
  ['Ctrl+Shift+~ ! @ # $ % ^', 'General / number / time / date / currency / percent / scientific format'],
  ['Ctrl+Shift+& / Ctrl+Shift+_', 'Add / remove outline border'], ['Ctrl+Shift+> / <', 'Increase / decrease font size'],
  ['Ctrl+Alt+= / Ctrl+Alt+-', 'Zoom in / out'], ['Ctrl+`', 'Show formulas'], ['Shift+F3', 'Insert Function'], ['F3', 'Paste Name'],
  ['Ctrl+Shift+A', 'Insert argument names (while editing)'], ['Ctrl+E', 'Flash Fill'], ['F7', 'Spelling'], ['Shift+F10', 'Context menu'],
  ['Ctrl+6', 'Hide / show objects'], ['Ctrl+Shift+U', 'Expand / collapse the formula bar'],
  ['Ctrl+N / Ctrl+O / Ctrl+S / F12', 'New / open / save / save as'], ['Ctrl+W / Ctrl+F4', 'Close the workbook'],
  ['Ctrl+P / Ctrl+F2', 'Print'], ['F1', 'Help'],
]
