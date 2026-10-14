// Adds Excel keyboard shortcuts that Univer lacks, and overrides ones that conflict with Excel
// (Ctrl+0 hides columns in Excel; Univer binds it to zoom reset).
import { ICommandService, CommandType } from '@univerjs/core'
import { whenSheetEditorFocused } from '@univerjs/sheets-ui'
import { KeyCode, MetaKeys } from '@univerjs/ui'

const { CTRL_COMMAND: CTRL, SHIFT, ALT } = MetaKeys
const PAGE_UP = 33, PAGE_DOWN = 34, SEMICOLON = 186, BACKQUOTE = 192
const PAGE_ROWS = 20
const excelSerial = (d) => Math.floor((d - new Date(1899, 11, 30)) / 864e5) + (d - new Date(d.getFullYear(), d.getMonth(), d.getDate())) / 864e5

export function installExcelShortcuts(api, shortcuts) {
  const commands = api._injector.get(ICommandService)
  const sheet = () => api.getActiveWorkbook().getActiveSheet()
  const cur = () => sheet().getActiveRange()
  const goto = (r, c) => { sheet().getRange(Math.max(0, r), Math.max(0, c)).activate(); return true }
  const lastCell = () => {
    const m = sheet().getSheet().getCellMatrix(); let lr = 0, lc = 0
    m.forValue((r, c, cell) => { if (cell && (cell.v != null || cell.f)) { lr = Math.max(lr, r); lc = Math.max(lc, c) } })
    return [lr, lc]
  }
  const defs = [
    ['go-top-left', CTRL | KeyCode.HOME, () => goto(0, 0)],
    ['go-last-cell', CTRL | KeyCode.END, () => goto(...lastCell())],
    ['go-row-start', KeyCode.HOME, () => goto(cur().getRow(), 0)],
    ['page-down', PAGE_DOWN, () => goto(cur().getRow() + PAGE_ROWS, cur().getColumn())],
    ['page-up', PAGE_UP, () => goto(cur().getRow() - PAGE_ROWS, cur().getColumn())],
    ['next-sheet', CTRL | PAGE_DOWN, () => { const wb = api.getActiveWorkbook(), s = wb.getSheets(), i = s.findIndex((x) => x.getSheetId() === sheet().getSheetId()); wb.setActiveSheet(s[(i + 1) % s.length]); return true }],
    ['prev-sheet', CTRL | PAGE_UP, () => { const wb = api.getActiveWorkbook(), s = wb.getSheets(), i = s.findIndex((x) => x.getSheetId() === sheet().getSheetId()); wb.setActiveSheet(s[(i - 1 + s.length) % s.length]); return true }],
    ['select-column', CTRL | KeyCode.SPACE, () => { const s = sheet(); s.getRange(0, cur().getColumn(), s.getMaxRows(), 1).activate(); return true }],
    ['select-row', SHIFT | KeyCode.SPACE, () => { const s = sheet(); s.getRange(cur().getRow(), 0, 1, s.getMaxColumns()).activate(); return true }],
    ['insert-date', CTRL | SEMICOLON, () => { const r = cur(); r.setValue(Math.floor(excelSerial(new Date()))); r.setNumberFormat('m/d/yyyy'); return true }],
    ['insert-time', CTRL | SHIFT | SEMICOLON, () => { const r = cur(), n = new Date(); r.setValue((n - new Date(n.getFullYear(), n.getMonth(), n.getDate())) / 864e5); r.setNumberFormat('h:mm AM/PM'); return true }],
    ['autosum', ALT | KeyCode.EQUAL, () => autosum(sheet(), cur())],
    ['hide-columns', CTRL | KeyCode.Digit0, () => { const r = cur(); sheet().hideColumns(r.getColumn(), r.getWidth()); return true }],
    ['unhide-rows', CTRL | SHIFT | KeyCode.Digit9, () => { const r = cur(); sheet().showRows(Math.max(0, r.getRow() - 1), r.getHeight() + 2); return true }],
    ['fmt-general', CTRL | SHIFT | 192, () => fmt('General')],
    ['fmt-number', CTRL | SHIFT | KeyCode.Digit1, () => fmt('#,##0.00')],
    ['fmt-time', CTRL | SHIFT | KeyCode.Digit2, () => fmt('h:mm AM/PM')],
    ['fmt-date', CTRL | SHIFT | KeyCode.Digit3, () => fmt('d-mmm-yy')],
    ['fmt-currency', CTRL | SHIFT | KeyCode.Digit4, () => fmt('$#,##0.00')],
    ['fmt-percent', CTRL | SHIFT | KeyCode.Digit5, () => fmt('0%')],
    ['fmt-scientific', CTRL | SHIFT | KeyCode.Digit6, () => fmt('0.00E+00')],
  ]
  const fmt = (p) => { cur().setNumberFormat(p); return true }
  for (const [name, binding, handler] of defs) {
    const id = `sheetly.${name}`
    commands.registerCommand({ id, type: CommandType.COMMAND, handler })
    shortcuts.registerShortcut({ id, binding, priority: 100, preconditions: whenSheetEditorFocused, description: name })
  }
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
  target.setFormula('=SUM()'); return true
}
