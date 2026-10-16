// Excel's ribbon KeyTips (Microsoft 365, English) and the Office 2003 access keys Excel still accepts, mapped to
// Sheetly actions. Only commands Sheetly implements are listed; a letter for a command it lacks does nothing,
// like a disabled ribbon button. Letters are from Microsoft's documentation and cross-checked lists (2026-10-16
// research); entries one source supports are marked "unverified" in PLAN.md.
import { BorderStyleTypes, BorderType, ICommandService } from '@univerjs/core'
import { el, field, openDialog, radios } from './ui/dialog.js'
import { autosum } from './excel-shortcuts.js'

const ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
const COMMA = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
const H_ALIGN = { left: 1, center: 2, right: 3 }, V_ALIGN = { top: 1, middle: 2, bottom: 3 }
const WRAP = 3, OVERFLOW = 1

export function buildKeyTipTree(api, { dialogs, extras, charts, outline }) {
  const commands = api._injector.get(ICommandService)
  const wb = () => api.getActiveWorkbook(), ws = () => wb().getActiveSheet(), cur = () => ws().getActiveRange(), sel = () => cur().getRange()
  const cmd = (id, params) => () => commands.executeCommand(id, params)
  const app = (name) => () => window.sheetly?.appAction?.(name)
  const n = (k, label, run) => ({ k, label, run })
  const menu = (k, label, ...items) => ({ k, label, items })

  // ---- small helpers for commands that have no Univer command of their own
  const border = (...parts) => () => { for (const [type, style] of parts) cur().setBorder(type, style, '#000000') }
  const fmt = (pattern) => () => cur().setNumberFormat(pattern)
  const wrapToggle = () => { const wrapped = cur().getCellStyleData?.()?.tb === WRAP; commands.executeCommand('sheet.command.set-text-wrap', { value: wrapped ? OVERFLOW : WRAP }) }
  const mergeCenter = async () => { await commands.executeCommand('sheet.command.add-worksheet-merge-all'); await commands.executeCommand('sheet.command.set-horizontal-text-align', { value: H_ALIGN.center }) }
  const sum = (fn) => () => autosum(ws(), cur(), fn)
  const fill = (dr, dc) => () => {
    const g = sel(), s = ws()
    if (dr) { const src = dr > 0 ? g.startRow : g.endRow; for (let c = g.startColumn; c <= g.endColumn; c++) copyCell(s, src, c, dr > 0 ? [src + 1, g.endRow] : [g.startRow, src - 1], 'row') }
    else { const src = dc > 0 ? g.startColumn : g.endColumn; for (let r = g.startRow; r <= g.endRow; r++) copyCell(s, r, src, dc > 0 ? [src + 1, g.endColumn] : [g.startColumn, src - 1], 'col') }
  }
  function copyCell(s, r, c, [a, b], axis) {
    const v = s.getRange(r, c).getValue(), f = s.getRange(r, c).getFormula()
    for (let i = a; i <= b; i++) { const t = axis === 'row' ? s.getRange(i, c) : s.getRange(r, i); if (f) t.setFormula(f); else t.setValue(v) }
  }
  // Insert as many rows / columns as are selected, before the selection (Excel's Insert Sheet Rows / Columns).
  const insertRows = () => { const g = sel(); ws().insertRowsBefore(g.startRow, g.endRow - g.startRow + 1) }
  const insertCols = () => { const g = sel(); ws().insertColumnsBefore(g.startColumn, g.endColumn - g.startColumn + 1) }
  const clearNotes = () => { const g = sel(); for (const x of ws().getNotes?.() ?? []) if (x.row >= g.startRow && x.row <= g.endRow && x.col >= g.startColumn && x.col <= g.endColumn) ws().getRange(x.row, x.col).deleteNote() }
  const prompt = (title, label, value, apply) => {
    const input = el('input', { type: 'text', value: String(value), 'aria-label': label })
    openDialog({ title, width: 260, initialFocus: input, body: field(label, input), onOk: () => apply(input.value) })
    setTimeout(() => input.select(), 0)
  }
  const rowHeight = () => prompt('Row Height', 'Row height:', Math.round(ws().getRowHeight(sel().startRow) * 0.75 * 100) / 100, (v) => {
    const pt = Number(v); if (!(pt >= 0 && pt <= 409)) return false
    const g = sel(); ws().setRowHeights(g.startRow, g.endRow - g.startRow + 1, Math.round(pt / 0.75))
  })
  const colWidth = () => prompt('Column Width', 'Column width:', Math.round(ws().getColumnWidth(sel().startColumn) / 7 * 100) / 100, (v) => {
    const ch = Number(v); if (!(ch >= 0 && ch <= 255)) return false
    const g = sel(); ws().setColumnWidths(g.startColumn, g.endColumn - g.startColumn + 1, Math.round(ch * 7))
  })
  const renameSheet = () => prompt('Rename Sheet', 'Name:', ws().getSheetName(), (v) => { if (!v.trim()) return false; ws().setName(v.trim()) })
  const unhideSheet = () => {
    const hidden = wb().getSheets().filter((s) => s.isSheetHidden())
    if (!hidden.length) return
    const r = radios('unhide-sheet', hidden.map((s) => [s.getSheetId(), s.getSheetName()]), hidden[0].getSheetId())
    openDialog({ title: 'Unhide', width: 260, body: el('div', {}, 'Unhide sheet:', ...r.nodes), onOk: () => { commands.executeCommand('sheet.command.set-worksheet-show', { unitId: wb().getId(), subUnitId: r.value }) } })
  }
  const zoom = (z) => ws().zoom(z)
  const zoomDialog = () => {
    const r = radios('zoom', [['2', '&200%'], ['1', '&100%'], ['0.75', '&75%'], ['0.5', '&50%'], ['0.25', '25%'], ['custom', 'Custo&m:']], String(ws().getZoom?.() ?? 1))
    const custom = el('input', { type: 'text', value: Math.round((ws().getZoom?.() ?? 1) * 100), 'aria-label': 'Custom zoom percent', style: { width: '60px' } })
    openDialog({ title: 'Zoom', width: 220, body: el('div', {}, ...r.nodes, custom), onOk: () => {
      const z = r.value === 'custom' ? Number(custom.value) / 100 : Number(r.value)
      if (!(z >= 0.1 && z <= 4)) return false
      zoom(z)
    } })
  }
  const freezeToggle = () => {
    const f = ws().getSheet().getConfig().freeze
    commands.executeCommand(f && (f.xSplit || f.ySplit) ? 'sheet.command.cancel-frozen' : 'sheet.command.set-selection-frozen')
  }
  // Toolbar drop-downs (fill colour, font colour, font, size): open Univer's own control, as Excel opens its gallery.
  const toolbar = (id) => () => {
    const host = [...document.querySelectorAll(`[data-u-command="${id}"]`)].find((x) => x.offsetParent)
    if (!host) return
    // Univer's toolbar menus are Radix drop-downs: they open on pointerdown.
    const trigger = host.matches('[data-slot="dropdown-menu-trigger"]') ? host : host.querySelector('[data-slot="dropdown-menu-trigger"]')
    trigger?.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, button: 0, pointerType: 'mouse' }))
    const input = host.querySelector('input')
    if (input) { input.focus(); input.select() }
  }
  const paste = (what, op = 'none', skip = false, transpose = false) => () => dialogs.runPasteSpecial(what, op, skip, transpose)
  const pasteValuesSourceFormat = async () => { await dialogs.runPasteSpecial('values', 'none', false, false); await dialogs.runPasteSpecial('formats', 'none', false, false) }
  const pasteKeepWidths = async () => { await dialogs.runPasteSpecial('all', 'none', false, false); await dialogs.runPasteSpecial('widths', 'none', false, false) }

  // ---- shared sub-menus
  const pasteMenu = menu('V', 'Paste',
    n('P', 'Paste', cmd('univer.command.paste')), n('F', 'Formulas', paste('formulas')), n('O', 'Formulas & Number Formatting', paste('formulasnum')),
    n('K', 'Keep Source Formatting', paste('all')), n('B', 'No Borders', paste('noborders')), n('W', 'Keep Source Column Widths', pasteKeepWidths),
    n('T', 'Transpose', paste('all', 'none', false, true)), n('V', 'Values', paste('values')), n('A', 'Values & Number Formatting', paste('valuesnum')),
    n('E', 'Values & Source Formatting', pasteValuesSourceFormat), n('R', 'Formatting', paste('formats')), n('N', 'Paste Link', () => dialogs.pasteLink()),
    n('S', 'Paste Special…', () => dialogs.pasteSpecial()))
  const T = BorderType, S = BorderStyleTypes
  const bordersMenu = menu('B', 'Borders',
    n('O', 'Bottom Border', border([T.BOTTOM, S.THIN])), n('P', 'Top Border', border([T.TOP, S.THIN])), n('L', 'Left Border', border([T.LEFT, S.THIN])),
    n('R', 'Right Border', border([T.RIGHT, S.THIN])), n('N', 'No Border', border([T.NONE, S.NONE])), n('A', 'All Borders', border([T.ALL, S.THIN])),
    n('S', 'Outside Borders', border([T.OUTSIDE, S.THIN])), n('T', 'Thick Outside Borders', border([T.OUTSIDE, S.THICK])),
    n('B', 'Bottom Double Border', border([T.BOTTOM, S.DOUBLE])), n('H', 'Thick Bottom Border', border([T.BOTTOM, S.THICK])),
    n('D', 'Top and Bottom Border', border([T.TOP, S.THIN], [T.BOTTOM, S.THIN])), n('C', 'Top and Thick Bottom Border', border([T.TOP, S.THIN], [T.BOTTOM, S.THICK])),
    n('U', 'Top and Double Bottom Border', border([T.TOP, S.THIN], [T.BOTTOM, S.DOUBLE])), n('M', 'More Borders…', () => dialogs.formatCells('Border')))
  const autoSumMenu = (k) => menu(k, 'AutoSum', n('S', 'Sum', sum('SUM')), n('A', 'Average', sum('AVERAGE')), n('C', 'Count Numbers', sum('COUNT')),
    n('M', 'Max', sum('MAX')), n('I', 'Min', sum('MIN')), n('F', 'More Functions…', () => extras.insertFunction()))
  const sortFilter = [n('S', 'Sort A to Z', cmd('sheet.command.sort-range-asc')), n('O', 'Sort Z to A', cmd('sheet.command.sort-range-desc')),
    n('U', 'Custom Sort…', cmd('sheet.command.sort-range-custom')), n('F', 'Filter', cmd('sheet.command.smart-toggle-filter')),
    n('C', 'Clear', cmd('sheet.command.clear-filter-criteria')), n('Y', 'Reapply', cmd('sheet.command.re-calc-filter'))]

  const tabs = [
    menu('F', 'File', n('N', 'New', app('new')), n('O', 'Open', app('open')), n('S', 'Save', app('save')), n('A', 'Save As', app('saveAs')),
      n('P', 'Print', () => dialogs.printPreview()), n('C', 'Close', app('close'))),
    menu('H', 'Home',
      pasteMenu, n('X', 'Cut', cmd('univer.command.cut')), n('C', 'Copy', cmd('univer.command.copy')),
      n('FP', 'Format Painter', cmd('sheet.command.set-once-format-painter')),
      n('FF', 'Font', toolbar('sheet.command.set-range-font-family')), n('FS', 'Font Size', toolbar('sheet.command.set-range-fontsize')),
      n('FG', 'Increase Font Size', cmd('sheet.command.set-range-font-increase')), n('FK', 'Decrease Font Size', cmd('sheet.command.set-range-font-decrease')),
      n('1', 'Bold', cmd('sheet.command.set-range-bold')), n('2', 'Italic', cmd('sheet.command.set-range-italic')),
      menu('3', 'Underline', n('U', 'Underline', cmd('sheet.command.set-range-underline'))),
      n('FC', 'Font Color', toolbar('sheet.command.set-range-text-color')), n('H', 'Fill Color', toolbar('sheet.command.set-background-color')),
      n('FN', 'Format Cells: Number', () => dialogs.formatCells('Number')), bordersMenu,
      menu('A', 'Align',
        n('T', 'Top Align', cmd('sheet.command.set-vertical-text-align', { value: V_ALIGN.top })),
        n('M', 'Middle Align', cmd('sheet.command.set-vertical-text-align', { value: V_ALIGN.middle })),
        n('B', 'Bottom Align', cmd('sheet.command.set-vertical-text-align', { value: V_ALIGN.bottom })),
        n('L', 'Align Left', cmd('sheet.command.set-horizontal-text-align', { value: H_ALIGN.left })),
        n('C', 'Center', cmd('sheet.command.set-horizontal-text-align', { value: H_ALIGN.center })),
        n('R', 'Align Right', cmd('sheet.command.set-horizontal-text-align', { value: H_ALIGN.right })),
        n('N', 'Accounting Number Format', fmt(ACCOUNTING))),
      n('W', 'Wrap Text', wrapToggle),
      menu('FQ', 'Orientation', n('O', 'Angle Counterclockwise', cmd('sheet.command.set-text-rotation', { value: 45 })),
        n('L', 'Angle Clockwise', cmd('sheet.command.set-text-rotation', { value: -45 })), n('V', 'Vertical Text', cmd('sheet.command.set-text-rotation', { value: 'v' })),
        n('U', 'Rotate Text Up', cmd('sheet.command.set-text-rotation', { value: 90 })), n('D', 'Rotate Text Down', cmd('sheet.command.set-text-rotation', { value: -90 })),
        n('M', 'Format Cell Alignment', () => dialogs.formatCells('Alignment'))),
      menu('M', 'Merge', n('C', 'Merge & Center', mergeCenter), n('A', 'Merge Across', cmd('sheet.command.add-worksheet-merge-horizontal')),
        n('M', 'Merge Cells', cmd('sheet.command.add-worksheet-merge-all')), n('U', 'Unmerge Cells', cmd('sheet.command.remove-worksheet-merge'))),
      n('N', 'Number Format', cmd('sheet.operation.open.numfmt.panel')), n('P', 'Percent Style', fmt('0%')), n('K', 'Comma Style', fmt(COMMA)),
      n('0', 'Increase Decimal', cmd('sheet.command.numfmt.add.decimal.command')), n('9', 'Decrease Decimal', cmd('sheet.command.numfmt.subtract.decimal.command')),
      menu('L', 'Conditional Formatting', n('N', 'New Rule…', cmd('sheet.operation.open.conditional.formatting.panel')),
        n('R', 'Manage Rules…', cmd('sheet.operation.open.conditional.formatting.panel')),
        menu('C', 'Clear Rules', n('S', 'Clear Rules from Selected Cells', cmd('sheet.command.clear-range-conditional-rule')),
          n('E', 'Clear Rules from Entire Sheet', cmd('sheet.command.clear-worksheet-conditional-rule')))),
      n('T', 'Format as Table', () => dialogs.createTable()),
      menu('I', 'Insert', n('I', 'Insert Cells…', () => dialogs.insertCells()), n('R', 'Insert Sheet Rows', insertRows),
        n('C', 'Insert Sheet Columns', insertCols), n('S', 'Insert Sheet', cmd('sheetly.new-sheet'))),
      menu('D', 'Delete', n('D', 'Delete Cells…', () => dialogs.deleteCells()), n('R', 'Delete Sheet Rows', cmd('sheet.command.remove-row')),
        n('C', 'Delete Sheet Columns', cmd('sheet.command.remove-col')), n('S', 'Delete Sheet', cmd('sheet.command.remove-sheet-confirm'))),
      menu('O', 'Format', n('H', 'Row Height…', rowHeight), n('A', 'AutoFit Row Height', cmd('sheet.command.set-row-is-auto-height')),
        n('W', 'Column Width…', colWidth), n('I', 'AutoFit Column Width', cmd('sheet.command.set-col-auto-width')),
        menu('U', 'Hide & Unhide', n('R', 'Hide Rows', cmd('sheetly.hide-rows')), n('C', 'Hide Columns', cmd('sheetly.hide-columns')),
          n('S', 'Hide Sheet', cmd('sheet.command.set-worksheet-hidden')), n('O', 'Unhide Rows', cmd('sheetly.unhide-rows')),
          n('L', 'Unhide Columns', cmd('sheetly.unhide-columns')), n('H', 'Unhide Sheet…', unhideSheet)),
        n('R', 'Rename Sheet', renameSheet), n('E', 'Format Cells…', () => dialogs.formatCells('Number'))),
      autoSumMenu('U'),
      menu('FI', 'Fill', n('D', 'Down', fill(1, 0)), n('R', 'Right', fill(0, 1)), n('U', 'Up', fill(-1, 0)), n('L', 'Left', fill(0, -1)),
        n('F', 'Flash Fill', () => extras.flashFill())),
      menu('E', 'Clear', n('A', 'Clear All', cmd('sheet.command.clear-selection-all')), n('F', 'Clear Formats', cmd('sheet.command.clear-selection-format')),
        n('C', 'Clear Contents', cmd('sheet.command.clear-selection-content')), n('M', 'Clear Comments and Notes', clearNotes)),
      menu('S', 'Sort & Filter', ...sortFilter),
      menu('FD', 'Find & Select', n('F', 'Find…', cmd('ui.operation.open-find-dialog')), n('R', 'Replace…', cmd('ui.operation.open-replace-dialog')),
        n('G', 'Go To…', () => dialogs.goTo()), n('S', 'Go To Special…', () => dialogs.goToSpecial()),
        n('U', 'Formulas', () => dialogs.runGoToSpecial('formulas')), n('M', 'Notes', () => dialogs.runGoToSpecial('notes')),
        n('N', 'Constants', () => dialogs.runGoToSpecial('constants')))),
    menu('N', 'Insert', n('T', 'Table', () => dialogs.createTable()), n('R', 'Recommended Charts', () => charts.embeddedChart()),
      n('C1', 'Column Chart', () => charts.embeddedChart()), n('I', 'Link', cmd('sheet.operation.open-hyper-link-edit-panel'))),
    menu('P', 'Page Layout', n('SP', 'Page Setup', () => dialogs.printPreview()), menu('V', 'View', n('G', 'Gridlines', cmd('sheet.command.toggle-gridlines')))),
    menu('M', 'Formulas', n('F', 'Insert Function', () => extras.insertFunction()), autoSumMenu('U'),
      n('N', 'Name Manager', cmd('sidebar.operation.defined-name', { value: 'open' })), n('C', 'Create from Selection', () => dialogs.createNames()),
      n('H', 'Show Formulas', () => extras.toggleShowFormulas()), n('B', 'Calculate Now', () => api.getFormula().executeCalculation()),
      n('J', 'Calculate Sheet', () => api.getFormula().executeCalculation())),
    menu('A', 'Data', menu('S', 'Sort', n('S', 'Sort…', cmd('sheet.command.sort-range-custom')), n('A', 'Sort A to Z', cmd('sheet.command.sort-range-asc')),
      n('D', 'Sort Z to A', cmd('sheet.command.sort-range-desc'))),
    n('T', 'Filter', cmd('sheet.command.smart-toggle-filter')), n('C', 'Clear', cmd('sheet.command.clear-filter-criteria')),
    n('Y3', 'Reapply', cmd('sheet.command.re-calc-filter')), n('E', 'Text to Columns', cmd('sheet.command.split-text-to-columns')),
    menu('F', 'Flash Fill', n('F', 'Flash Fill', () => extras.flashFill())),
    menu('V', 'Data Validation', n('V', 'Data Validation…', cmd('data-validation.operation.open-validation-panel'))),
    menu('G', 'Group', n('G', 'Group…', () => outline.group())), menu('U', 'Ungroup', n('U', 'Ungroup…', () => outline.ungroup()), n('C', 'Clear Outline', () => outline.clearOutline())),
    n('J', 'Show Detail', () => outline.showDetail()), n('H', 'Hide Detail', () => outline.hideDetail())),
    menu('R', 'Review', n('S', 'Spelling', () => extras.spelling()), n('C', 'New Comment', cmd('sheet.operation.show-comment-modal'))),
    menu('W', 'View', menu('V', 'Show', n('G', 'Gridlines', cmd('sheet.command.toggle-gridlines'))),
      n('Q', 'Zoom…', zoomDialog), n('J', '100%', () => zoom(1)),
      menu('F', 'Freeze Panes', n('F', 'Freeze Panes', freezeToggle), n('R', 'Freeze Top Row', cmd('sheet.command.set-first-row-frozen')),
        n('C', 'Freeze First Column', cmd('sheet.command.set-first-column-frozen')))),
  ]

  // Office 2003 menus (letters no ribbon tab uses).
  const legacy = [
    menu('E', 'Edit', n('S', 'Paste Special…', () => dialogs.pasteSpecial()), n('D', 'Delete…', () => dialogs.deleteCells()),
      n('L', 'Delete Sheet', cmd('sheet.command.remove-sheet-confirm')),
      menu('A', 'Clear', n('A', 'All', cmd('sheet.command.clear-selection-all')), n('F', 'Formats', cmd('sheet.command.clear-selection-format')),
        n('C', 'Contents', cmd('sheet.command.clear-selection-content'))),
      menu('I', 'Fill', n('D', 'Down', fill(1, 0)), n('R', 'Right', fill(0, 1)), n('U', 'Up', fill(-1, 0)), n('L', 'Left', fill(0, -1))),
      n('F', 'Find…', cmd('ui.operation.open-find-dialog')), n('E', 'Replace…', cmd('ui.operation.open-replace-dialog')), n('G', 'Go To…', () => dialogs.goTo())),
    menu('I', 'Insert', n('R', 'Rows', insertRows), n('C', 'Columns', insertCols),
      n('W', 'Worksheet', cmd('sheetly.new-sheet')), n('E', 'Cells…', () => dialogs.insertCells())),
    menu('O', 'Format', n('E', 'Cells…', () => dialogs.formatCells('Number')),
      menu('R', 'Row', n('E', 'Height…', rowHeight), n('A', 'AutoFit', cmd('sheet.command.set-row-is-auto-height')),
        n('H', 'Hide', cmd('sheetly.hide-rows')), n('U', 'Unhide', cmd('sheetly.unhide-rows'))),
      menu('C', 'Column', n('W', 'Width…', colWidth), n('A', 'AutoFit Selection', cmd('sheet.command.set-col-auto-width')),
        n('H', 'Hide', cmd('sheetly.hide-columns')), n('U', 'Unhide', cmd('sheetly.unhide-columns'))),
      menu('H', 'Sheet', n('R', 'Rename', renameSheet), n('H', 'Hide', cmd('sheet.command.set-worksheet-hidden')), n('U', 'Unhide…', unhideSheet))),
    menu('D', 'Data', n('S', 'Sort…', cmd('sheet.command.sort-range-custom')), menu('F', 'Filter', n('F', 'AutoFilter', cmd('sheet.command.smart-toggle-filter'))),
      menu('G', 'Group and Outline', n('G', 'Group…', () => outline.group()), n('U', 'Ungroup…', () => outline.ungroup()),
        n('H', 'Hide Detail', () => outline.hideDetail()), n('S', 'Show Detail', () => outline.showDetail()), n('C', 'Clear Outline', () => outline.clearOutline())),
      n('L', 'Validation…', cmd('data-validation.operation.open-validation-panel')), n('E', 'Text to Columns…', cmd('sheet.command.split-text-to-columns'))),
    menu('T', 'Tools', n('S', 'Spelling…', () => extras.spelling())),
    menu('V', 'View', n('Z', 'Zoom…', zoomDialog)),
  ]
  return { tabs, legacy }
}
