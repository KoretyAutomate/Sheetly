// Excel dialogs reached by keyboard shortcuts: Go To / Go To Special, Insert, Delete, Create Names,
// Create Table, Paste Special, Format Cells, Print preview, and the Help pane.
import { BorderStyleTypes, BorderType, ICommandService, numfmt, TextDecoration } from '@univerjs/core'
import { button, check, el, field, group, openDialog, radios } from './ui/dialog.js'

// ---------- references
export const colName = (c) => { let s = ''; for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s }
export const colIndex = (s) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
export const a1 = (r, c, abs = false) => (abs ? `$${colName(c)}$${r + 1}` : `${colName(c)}${r + 1}`)
export const rangeA1 = (g, abs = false) => (g.startRow === g.endRow && g.startColumn === g.endColumn ? a1(g.startRow, g.startColumn, abs)
  : `${a1(g.startRow, g.startColumn, abs)}:${a1(g.endRow, g.endColumn, abs)}`)
const quoteSheet = (n) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`)

/** Parses "A1", "A1:B2", "B:B", "3:5", "Sheet 2!A1", "'My Sheet'!A1:C3". Returns {sheetName?, range} or null. */
export function parseRef(text, maxRows, maxCols) {
  let t = String(text).trim().replace(/^=/, '')
  let sheetName
  const bang = t.lastIndexOf('!')
  if (bang > 0) { sheetName = t.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'"); t = t.slice(bang + 1) }
  t = t.replace(/\$/g, '').toUpperCase()
  let m
  if ((m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(t))) {
    const r1 = +m[2] - 1, c1 = colIndex(m[1]), r2 = m[4] ? +m[4] - 1 : r1, c2 = m[3] ? colIndex(m[3]) : c1
    return { sheetName, range: { startRow: Math.min(r1, r2), endRow: Math.max(r1, r2), startColumn: Math.min(c1, c2), endColumn: Math.max(c1, c2) } }
  }
  if ((m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(t))) {
    const c1 = colIndex(m[1]), c2 = colIndex(m[2])
    return { sheetName, range: { startRow: 0, endRow: maxRows - 1, startColumn: Math.min(c1, c2), endColumn: Math.max(c1, c2) } }
  }
  if ((m = /^(\d+):(\d+)$/.exec(t))) {
    const r1 = +m[1] - 1, r2 = +m[2] - 1
    return { sheetName, range: { startRow: Math.min(r1, r2), endRow: Math.max(r1, r2), startColumn: 0, endColumn: maxCols - 1 } }
  }
  return null
}

/** Shifts the relative row/column parts of A1 references in a formula (used for transpose and paste link). */
export function shiftFormula(formula, dRow, dCol) {
  let out = '', inStr = false
  for (let i = 0; i < formula.length;) {
    const ch = formula[i]
    if (ch === '"') { inStr = !inStr; out += ch; i++; continue }
    if (inStr) { out += ch; i++; continue }
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])/.exec(formula.slice(i))
    const prev = i > 0 ? formula[i - 1] : ''
    if (m && !/[\w.]/.test(prev)) {
      const c = m[1] ? colIndex(m[2]) : colIndex(m[2]) + dCol, r = m[3] ? +m[4] - 1 : +m[4] - 1 + dRow
      out += r < 0 || c < 0 ? '#REF!' : `${m[1]}${colName(c)}${m[3]}${r + 1}`
      i += m[0].length
      continue
    }
    out += ch; i++
  }
  return out
}

// ---------- helpers bound to the running app
export function createDialogs(api) {
  const commands = api._injector.get(ICommandService)
  const wb = () => api.getActiveWorkbook()
  const ws = () => wb().getActiveSheet()
  const sel = () => ws().getActiveRange().getRange()
  const cellRaw = (r, c) => ws().getSheet().getCellRaw(r, c)
  const hasContent = (r, c) => { const x = cellRaw(r, c); return !!x && (x.f != null || x.si != null || (x.v != null && x.v !== '') || !!x.p) }
  const usedRange = () => {
    let er = 0, ec = 0
    ws().getSheet().getCellMatrix().forValue((r, c) => { if (hasContent(r, c)) { er = Math.max(er, r); ec = Math.max(ec, c) } })
    return { startRow: 0, startColumn: 0, endRow: er, endColumn: ec }
  }
  /** Excel's CurrentRegion: the block bounded by empty rows and columns around the cell. */
  const currentRegion = (r, c) => {
    let g = { startRow: r, endRow: r, startColumn: c, endColumn: c }
    for (let grown = true; grown;) {
      grown = false
      const rowHas = (row) => { for (let x = Math.max(0, g.startColumn - 1); x <= g.endColumn + 1; x++) if (row >= 0 && hasContent(row, x)) return true; return false }
      const colHas = (col) => { for (let y = Math.max(0, g.startRow - 1); y <= g.endRow + 1; y++) if (col >= 0 && hasContent(y, col)) return true; return false }
      if (g.startRow > 0 && rowHas(g.startRow - 1)) { g.startRow--; grown = true }
      if (rowHas(g.endRow + 1)) { g.endRow++; grown = true }
      if (g.startColumn > 0 && colHas(g.startColumn - 1)) { g.startColumn--; grown = true }
      if (colHas(g.endColumn + 1)) { g.endColumn++; grown = true }
    }
    return g
  }
  const select = (ranges, sheet = ws()) => {
    if (!ranges.length) return false
    const [first] = ranges
    commands.executeCommand('sheet.operation.set-selections', {
      unitId: wb().getId(), subUnitId: sheet.getSheetId(), reveal: true,
      selections: ranges.map((range, i) => ({ range: { ...range, rangeType: 0 }, style: null,
        primary: i === 0 ? { startRow: first.startRow, startColumn: first.startColumn, endRow: first.startRow, endColumn: first.startColumn,
          actualRow: first.startRow, actualColumn: first.startColumn, isMerged: false, isMergedMainCell: false } : null })),
    })
    return true
  }
  // Rectangular ranges from a set of cells, merged row-wise.
  const toRanges = (cells) => {
    const byRow = new Map()
    for (const [r, c] of cells) { if (!byRow.has(r)) byRow.set(r, []); byRow.get(r).push(c) }
    const rows = []
    for (const [r, cols] of [...byRow].sort((a, b) => a[0] - b[0])) {
      cols.sort((a, b) => a - b)
      let s = cols[0], p = cols[0]
      for (const c of [...cols.slice(1), Infinity]) { if (c === p + 1) { p = c; continue } rows.push({ startRow: r, endRow: r, startColumn: s, endColumn: p }); s = p = c }
    }
    // Merge identical column spans on consecutive rows into rectangles.
    const out = []
    for (const g of rows) {
      const prev = out.find((o) => o.endRow === g.startRow - 1 && o.startColumn === g.startColumn && o.endColumn === g.endColumn)
      if (prev) prev.endRow = g.endRow; else out.push({ ...g })
    }
    return out
  }
  const cellKind = (r, c) => {
    const x = cellRaw(r, c)
    if (!x || !hasContent(r, c)) return 'blank'
    if (x.f != null || x.si != null) return 'formula'
    return 'constant'
  }
  const valueKind = (r, c) => {
    const v = ws().getRange(r, c).getCalculatedValue()
    if (typeof v === 'number') return 'numbers'
    if (typeof v === 'boolean' || cellRaw(r, c)?.t === 3) return 'logicals'
    if (typeof v === 'string' && /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|CONNECT!|BLOCKED!|UNKNOWN!|FIELD!|BUSY!|GETTING_DATA)$/.test(v)) return 'errors'
    return 'text'
  }

  // ---------- Go To (F5 / Ctrl+G)
  const goToHistory = []
  function goTo() {
    const names = (wb().getDefinedNames?.() ?? []).filter((n) => !n.getName().startsWith('SheetlyRef_')).map((n) => n.getName())
    const input = el('input', { type: 'text', 'aria-label': 'Reference' })
    const list = el('div', { class: 'sx-list', role: 'listbox' },
      ...[...names, ...goToHistory].map((n) => el('div', { role: 'option', onClick: (e) => { input.value = n; list.querySelectorAll('.sel').forEach((x) => x.classList.remove('sel')); e.target.classList.add('sel') },
        onDblclick: () => { input.value = n; d.ok() } }, n)))
    const d = openDialog({ title: 'Go To', width: 320, initialFocus: input,
      body: el('div', { class: 'sx-body', style: { padding: 0 } }, field('Go to:', list), field('Reference:', input)),
      extraButtons: [el('button', { class: 'sx-btn', type: 'button', onClick: () => { d.close(); goToSpecial() } }, 'Special...')],
      onOk: () => goToReference(input.value) })
  }
  function goToReference(text) {
    const t = String(text).trim()
    if (!t) return false
    const dn = wb().getDefinedName?.(t)
    const target = dn ? dn.getFormulaOrRefString().replace(/^=/, '') : t
    const p = parseRef(target, ws().getMaxRows(), ws().getMaxColumns())
    if (!p) { alertBox('Reference isn’t valid.'); return false }
    const sheet = p.sheetName ? wb().getSheetByName(p.sheetName) : ws()
    if (!sheet) { alertBox('Reference isn’t valid.'); return false }
    const prev = rangeA1(sel(), true)
    if (!goToHistory.includes(prev)) goToHistory.unshift(prev)
    if (sheet.getSheetId() !== ws().getSheetId()) wb().setActiveSheet(sheet)
    sheet.getRange(p.range.startRow, p.range.startColumn, p.range.endRow - p.range.startRow + 1, p.range.endColumn - p.range.startColumn + 1).activate()
    return true
  }
  function goToSpecial() {
    const kind = radios('gts', [['notes', 'Notes'], ['constants', 'Constants'], ['formulas', 'Formulas'], ['blanks', 'Blanks'],
      ['region', 'Current region'], ['last', 'Last cell'], ['visible', 'Visible cells only'], ['rowdiff', 'Row differences'],
      ['coldiff', 'Column differences'], ['precedents', 'Precedents'], ['dependents', 'Dependents']], 'constants')
    const subs = ['numbers', 'text', 'logicals', 'errors'].map((k) => [k, check(k[0].toUpperCase() + k.slice(1), true)])
    openDialog({ title: 'Go To Special', width: 400,
      body: el('div', {}, group('Select', ...kind.nodes), el('div', { class: 'sx-row' }, ...subs.map(([, c]) => c.node))),
      onOk: () => {
        const chosen = subs.filter(([, c]) => c.input.checked).map(([k]) => k)
        const ok = runGoToSpecial(kind.value, chosen)
        if (!ok) { alertBox('No cells were found.'); return false }
      } })
  }
  function runGoToSpecial(kind, subtypes = ['numbers', 'text', 'logicals', 'errors']) {
    const s = sel()
    const single = s.startRow === s.endRow && s.startColumn === s.endColumn
    const scope = single ? usedRange() : s
    const cells = []
    const each = (fn) => { for (let r = scope.startRow; r <= scope.endRow; r++) for (let c = scope.startColumn; c <= scope.endColumn; c++) if (fn(r, c)) cells.push([r, c]) }
    const sheet = ws().getSheet()
    switch (kind) {
      case 'region': return select([currentRegion(s.startRow, s.startColumn)])
      case 'last': { const u = usedRange(); return select([{ startRow: u.endRow, endRow: u.endRow, startColumn: u.endColumn, endColumn: u.endColumn }]) }
      case 'blanks': each((r, c) => cellKind(r, c) === 'blank'); break
      case 'constants': each((r, c) => cellKind(r, c) === 'constant' && subtypes.includes(valueKind(r, c))); break
      case 'formulas': each((r, c) => cellKind(r, c) === 'formula' && subtypes.includes(valueKind(r, c))); break
      case 'notes': each((r, c) => !!ws().getRange(r, c).getNote?.()); break
      case 'visible': each((r, c) => sheet.getRowVisible(r) && sheet.getColVisible(c)); break
      case 'rowdiff': { const ac = s.startColumn; each((r, c) => c !== ac && String(ws().getRange(r, c).getCalculatedValue()) !== String(ws().getRange(r, ac).getCalculatedValue())); break }
      case 'coldiff': { const ar = s.startRow; each((r, c) => r !== ar && String(ws().getRange(r, c).getCalculatedValue()) !== String(ws().getRange(ar, c).getCalculatedValue())); break }
      case 'precedents': {
        const f = cellRaw(s.startRow, s.startColumn)?.f ?? ''
        const ranges = [...f.matchAll(/(?<![\w.!$])\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![\w(])/g)]
          .map((m) => parseRef(m[0], ws().getMaxRows(), ws().getMaxColumns())?.range).filter(Boolean)
        return select(ranges)
      }
      case 'dependents': {
        const me = { r: s.startRow, c: s.startColumn }
        const u = usedRange()
        for (let r = 0; r <= u.endRow; r++) for (let c = 0; c <= u.endColumn; c++) {
          const f = cellRaw(r, c)?.f
          if (!f) continue
          const hit = [...f.matchAll(/(?<![\w.!$])\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![\w(])/g)].some((m) => {
            const g = parseRef(m[0], ws().getMaxRows(), ws().getMaxColumns())?.range
            return g && me.r >= g.startRow && me.r <= g.endRow && me.c >= g.startColumn && me.c <= g.endColumn
          })
          if (hit) cells.push([r, c])
        }
        break
      }
    }
    return cells.length ? select(toRanges(cells)) : false
  }

  // ---------- Insert / Delete (Ctrl+Shift+= / Ctrl+-)
  const isFullRows = (g) => g.startColumn === 0 && g.endColumn >= ws().getMaxColumns() - 1
  const isFullCols = (g) => g.startRow === 0 && g.endRow >= ws().getMaxRows() - 1
  function insertCells() {
    const g = sel()
    if (isFullRows(g)) return ws().insertRowsBefore(g.startRow, g.endRow - g.startRow + 1), true
    if (isFullCols(g)) return ws().insertColumnsBefore(g.startColumn, g.endColumn - g.startColumn + 1), true
    const r = radios('ins', [['right', 'Shift cells r&ight'], ['down', 'Shift cells &down'], ['row', 'Entire &row'], ['col', 'Entire &column']], 'down')
    openDialog({ title: 'Insert', width: 240, body: el('div', {}, ...r.nodes), onOk: () => {
      if (r.value === 'right') commands.executeCommand('sheet.command.insert-range-move-right')
      else if (r.value === 'down') commands.executeCommand('sheet.command.insert-range-move-down')
      else if (r.value === 'row') ws().insertRowsBefore(g.startRow, g.endRow - g.startRow + 1)
      else ws().insertColumnsBefore(g.startColumn, g.endColumn - g.startColumn + 1)
    } })
    return true
  }
  function deleteCells() {
    const g = sel()
    if (isFullRows(g)) return ws().deleteRows(g.startRow, g.endRow - g.startRow + 1), true
    if (isFullCols(g)) return ws().deleteColumns(g.startColumn, g.endColumn - g.startColumn + 1), true
    const r = radios('del', [['left', 'Shift cells &left'], ['up', 'Shift cells &up'], ['row', 'Entire &row'], ['col', 'Entire &column']], 'up')
    openDialog({ title: 'Delete', width: 240, body: el('div', {}, ...r.nodes), onOk: () => {
      if (r.value === 'left') commands.executeCommand('sheet.command.delete-range-move-left')
      else if (r.value === 'up') commands.executeCommand('sheet.command.delete-range-move-up')
      else if (r.value === 'row') ws().deleteRows(g.startRow, g.endRow - g.startRow + 1)
      else ws().deleteColumns(g.startColumn, g.endColumn - g.startColumn + 1)
    } })
    return true
  }

  // ---------- Create Names from Selection (Ctrl+Shift+F3)
  const cleanName = (s) => {
    let n = String(s).trim().replace(/[^\p{L}\p{N}_.\\]/gu, '_')
    if (!n) return null
    if (/^[\p{N}.]/u.test(n)) n = '_' + n
    if (/^[A-Za-z]{1,3}\d+$/.test(n) || /^[RrCc]$/.test(n)) n += '_'
    return n
  }
  function createNames() {
    const g = sel()
    const isText = (r, c) => typeof ws().getRange(r, c).getCalculatedValue() === 'string' && ws().getRange(r, c).getCalculatedValue() !== ''
    const top = check('Top row', g.endRow > g.startRow && isText(g.startRow, Math.min(g.startColumn + 1, g.endColumn)))
    const left = check('Left column', g.endColumn > g.startColumn && isText(Math.min(g.startRow + 1, g.endRow), g.startColumn))
    const bottom = check('Bottom row'), right = check('Right column')
    openDialog({ title: 'Create Names from Selection', width: 300,
      body: group('Create names from values in the:', top.node, left.node, bottom.node, right.node),
      onOk: () => createNamesFrom(g, { top: top.input.checked, left: left.input.checked, bottom: bottom.input.checked, right: right.input.checked }) })
    return true
  }
  function createNamesFrom(g, o) {
    const sheet = quoteSheet(ws().getSheetName())
    const r0 = g.startRow + (o.top ? 1 : 0), r1 = g.endRow - (o.bottom ? 1 : 0)
    const c0 = g.startColumn + (o.left ? 1 : 0), c1 = g.endColumn - (o.right ? 1 : 0)
    const add = (label, range) => {
      const name = cleanName(label)
      if (!name || r1 < r0 || c1 < c0) return
      // A label that is not a valid name here (e.g. it clashes with a function) is skipped, as Excel skips invalid labels.
      try { wb().insertDefinedName(name, `${sheet}!${rangeA1(range, true)}`) } catch { /* invalid name */ }
    }
    for (let c = c0; c <= c1; c++) {
      if (o.top) add(ws().getRange(g.startRow, c).getCalculatedValue(), { startRow: r0, endRow: r1, startColumn: c, endColumn: c })
      if (o.bottom) add(ws().getRange(g.endRow, c).getCalculatedValue(), { startRow: r0, endRow: r1, startColumn: c, endColumn: c })
    }
    for (let r = r0; r <= r1; r++) {
      if (o.left) add(ws().getRange(r, g.startColumn).getCalculatedValue(), { startRow: r, endRow: r, startColumn: c0, endColumn: c1 })
      if (o.right) add(ws().getRange(r, g.endColumn).getCalculatedValue(), { startRow: r, endRow: r, startColumn: c0, endColumn: c1 })
    }
  }

  // ---------- Create Table (Ctrl+T / Ctrl+L)
  function createTable() {
    const s = sel()
    const g = s.startRow === s.endRow && s.startColumn === s.endColumn ? currentRegion(s.startRow, s.startColumn) : s
    const input = el('input', { type: 'text', value: '=' + rangeA1(g, true), 'aria-label': 'Where is the data for your table?' })
    const firstRowText = [...Array(g.endColumn - g.startColumn + 1)].every((_, i) => typeof ws().getRange(g.startRow, g.startColumn + i).getCalculatedValue() === 'string')
    const headers = check('My table has headers', firstRowText && g.endRow > g.startRow)
    openDialog({ title: 'Create Table', width: 300, initialFocus: input,
      body: el('div', {}, field('Where is the data for your table?', input), headers.node),
      onOk: () => {
        const p = parseRef(input.value, ws().getMaxRows(), ws().getMaxColumns())
        if (!p) { alertBox('The reference isn’t valid.'); return false }
        makeTable(p.range, headers.input.checked)
      } })
    return true
  }
  async function makeTable(g, hasHeaders) {
    let range = { ...g }
    if (!hasHeaders) {
      // Excel inserts a header row above the data and names the columns Column1, Column2, ...
      ws().getRange(g.startRow, g.startColumn, 1, g.endColumn - g.startColumn + 1).insertCells(0)
      for (let c = g.startColumn; c <= g.endColumn; c++) ws().getRange(g.startRow, c).setValue(`Column${c - g.startColumn + 1}`)
      range = { ...g, endRow: g.endRow + 1 }
    }
    let n = 1
    const existing = new Set((wb().getTableList?.() ?? []).map((t) => t.name))
    while (existing.has(`Table${n}`)) n++
    return ws().addTable(`Table${n}`, range)
  }

  // ---------- Paste Special (Ctrl+Alt+V)
  let lastCopy = null
  commands.onCommandExecuted((cmd) => {
    if (cmd.id !== 'univer.command.copy' && cmd.id !== 'univer.command.cut' && cmd.id !== 'sheet.command.copy') return
    const g = sel(), sheet = ws()
    const cells = []
    for (let r = g.startRow; r <= g.endRow; r++) {
      const row = []
      for (let c = g.startColumn; c <= g.endColumn; c++) {
        const range = sheet.getRange(r, c)
        row.push({ raw: sheet.getSheet().getCellRaw(r, c) ?? null, value: range.getCalculatedValue(), style: range.getCellStyleData?.() ?? null, numfmt: range.getNumberFormat?.() })
      }
      cells.push(row)
    }
    lastCopy = { sheetName: sheet.getSheetName(), sheetId: sheet.getSheetId(), range: g, cells }
  })
  function pasteSpecial() {
    if (!lastCopy) { alertBox('Copy something first (Ctrl+C), then use Paste Special.'); return true }
    // Access keys as in Excel's Paste Special dialog (Alt E S V = values).
    const what = radios('ps', [['all', '&All'], ['formulas', '&Formulas'], ['values', '&Values'], ['formats', 'Forma&ts'],
      ['notes', '&Comments and Notes'], ['validation', 'Validatio&n'], ['theme', 'All using Source t&heme'], ['noborders', 'All e&xcept borders'],
      ['widths', 'Column &widths'], ['formulasnum', 'Fo&rmulas and number formats'], ['valuesnum', 'Val&ues and number formats'],
      ['mergecf', 'All mer&ging conditional formats']], 'all')
    const op = radios('op', [['none', 'N&one'], ['add', 'A&dd'], ['subtract', '&Subtract'], ['multiply', '&Multiply'], ['divide', 'D&ivide']], 'none')
    const skip = check('Skip &blanks'), transpose = check('Transpos&e')
    const d = openDialog({ title: 'Paste Special', width: 440,
      body: el('div', {}, group('Paste', ...what.nodes), group('Operation', ...op.nodes), el('div', { class: 'sx-row' }, skip.node, transpose.node)),
      extraButtons: [button('Paste &Link', () => { pasteLink(); d.close() })],
      onOk: () => { runPasteSpecial(what.value, op.value, skip.input.checked, transpose.input.checked) } })
    return true
  }
  const hooks = { formulas: 'special-paste-formula', values: 'special-paste-value', formats: 'special-paste-format',
    noborders: 'special-paste-besides-border', widths: 'special-paste-col-width', formulasnum: 'special-paste-formula', valuesnum: 'special-paste-value' }
  async function runPasteSpecial(what, op, skipBlanks, transpose) {
    const manual = op !== 'none' || skipBlanks || transpose || what === 'notes' || what === 'validation'
    if (!manual) {
      if (what === 'all' || what === 'theme' || what === 'mergecf') await commands.executeCommand('univer.command.paste')
      else await commands.executeCommand('univer.command.paste', { value: hooks[what] })
      if (what === 'formulasnum' || what === 'valuesnum') applyNumberFormats(false)
      return
    }
    const t = sel(), src = lastCopy
    const rows = src.cells.length, cols = src.cells[0].length
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
      const s = src.cells[r][c]
      const tr = t.startRow + (transpose ? c : r), tc = t.startColumn + (transpose ? r : c)
      const blank = !s.raw || (s.raw.v == null && s.raw.f == null && s.raw.si == null && !s.raw.p)
      if (skipBlanks && blank) continue
      const target = ws().getRange(tr, tc)
      const withFormulas = ['all', 'theme', 'mergecf', 'formulas', 'formulasnum', 'noborders'].includes(what)
      const withValues = withFormulas || ['values', 'valuesnum'].includes(what)
      const withFormats = ['all', 'theme', 'mergecf', 'formats', 'noborders'].includes(what)
      if (what === 'notes') { const note = ws().getRange(src.range.startRow + r, src.range.startColumn + c).getNote?.(); if (note) target.createOrUpdateNote?.(note); continue }
      if (withValues) {
        const formula = s.raw?.f
        if (op !== 'none' && typeof s.value === 'number' && !formula) {
          const cur = target.getCalculatedValue()
          const base = typeof cur === 'number' ? cur : (cur == null || cur === '' ? 0 : null)
          if (base != null) {
            const v = op === 'add' ? base + s.value : op === 'subtract' ? base - s.value : op === 'multiply' ? base * s.value : (s.value === 0 ? '#DIV/0!' : base / s.value)
            target.setValue(v)
          }
        } else if (formula && withFormulas) {
          target.setFormula(shiftFormula(formula, tr - (src.range.startRow + r), tc - (src.range.startColumn + c)))
        } else if (!blank) target.setValue(s.value)
        else target.setValue(null)
      }
      if (withFormats && s.style) target.setValues?.([[{ s: what === 'noborders' ? { ...s.style, bd: undefined } : s.style }]])
      if ((what === 'formulasnum' || what === 'valuesnum') && s.numfmt) target.setNumberFormat(s.numfmt)
    }
  }
  function applyNumberFormats(transpose) {
    const t = sel(), src = lastCopy
    src.cells.forEach((row, r) => row.forEach((s, c) => {
      if (s.numfmt) ws().getRange(t.startRow + (transpose ? c : r), t.startColumn + (transpose ? r : c)).setNumberFormat(s.numfmt)
    }))
  }
  function pasteLink() {
    const t = sel(), src = lastCopy
    const prefix = src.sheetId === ws().getSheetId() ? '' : quoteSheet(src.sheetName) + '!'
    src.cells.forEach((row, r) => row.forEach((_, c) => ws().getRange(t.startRow + r, t.startColumn + c)
      .setFormula(`=${prefix}${a1(src.range.startRow + r, src.range.startColumn + c)}`)))
  }

  // ---------- Format Cells (Ctrl+1; Ctrl+Shift+F / Ctrl+Shift+P open the Font tab)
  function formatCells(startTab = 'Number') {
    const range = ws().getActiveRange()
    const style = range.getCellStyleData?.() ?? {}
    const value = range.getCalculatedValue()
    const tabs = ['Number', 'Alignment', 'Font', 'Border', 'Fill', 'Protection']
    const panels = {}, apply = []
    // Number
    {
      const cats = {
        General: () => 'General',
        Number: (o) => withNeg(`${o.sep ? '#,##0' : '0'}${dec(o.dec)}`, o.neg),
        Currency: (o) => withNeg(`$#,##0${dec(o.dec)}`, o.neg),
        Accounting: (o) => `_($* #,##0${dec(o.dec)}_);_($* (#,##0${dec(o.dec)});_($* "-"${'?'.repeat(o.dec)}_);_(@_)`,
        Date: (o) => o.type, Time: (o) => o.type,
        Percentage: (o) => `0${dec(o.dec)}%`,
        Fraction: (o) => o.type,
        Scientific: (o) => `0${dec(o.dec)}E+00`,
        Text: () => '@',
        Special: (o) => o.type,
        Custom: (o) => o.code,
      }
      const typeLists = {
        Date: ['m/d/yyyy', 'dddd, mmmm d, yyyy', 'm/d', 'm/d/yy', 'mm/dd/yy', 'd-mmm', 'd-mmm-yy', 'dd-mmm-yy', 'mmm-yy', 'mmmm-yy', 'mmmm d, yyyy', 'm/d/yy h:mm AM/PM', 'm/d/yy h:mm', 'mmmmm', 'mmmmm-yy', 'd-mmm-yyyy'],
        Time: ['h:mm:ss AM/PM', 'h:mm', 'h:mm AM/PM', 'h:mm:ss', 'mm:ss.0', '[h]:mm:ss', 'm/d/yy h:mm AM/PM', 'm/d/yy h:mm'],
        Fraction: ['# ?/?', '# ??/??', '# ???/???', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100'],
        Special: ['00000', '00000-0000', '[<=9999999]###-####;(###) ###-####', '000-00-0000'],
      }
      const current = range.getNumberFormat?.() || 'General'
      const catList = el('div', { class: 'sx-list', role: 'listbox', 'aria-label': 'Category' })
      const decimals = el('input', { type: 'number', min: 0, max: 30, value: 2, 'aria-label': 'Decimal places' })
      const sep = check('Use 1000 Separator (,)')
      const neg = el('select', { 'aria-label': 'Negative numbers' }, ...['-1234.10', '1234.10 (red)', '(1234.10)', '(1234.10) (red)'].map((t, i) => el('option', { value: i }, t)))
      const typeSel = el('select', { 'aria-label': 'Type', size: 6 })
      const code = el('input', { type: 'text', value: current, 'aria-label': 'Type' })
      const sample = el('div', { class: 'sx-sample', 'aria-label': 'Sample' })
      let cat = 'General'
      for (const name of Object.keys(cats)) {
        catList.append(el('div', { role: 'option', 'data-cat': name, onClick: () => setCat(name) }, name))
      }
      const opts = () => ({ dec: Math.max(0, Math.min(30, +decimals.value || 0)), sep: sep.input.checked, neg: +neg.value, type: typeSel.value, code: code.value })
      const pattern = () => cats[cat](opts())
      const refresh = () => {
        let text = ''
        try { text = numfmt.format(pattern(), typeof value === 'number' ? value : value ?? '') } catch { text = '' }
        sample.textContent = text
        if (cat !== 'Custom') code.value = pattern()
      }
      const setCat = (name) => {
        cat = name
        catList.querySelectorAll('div').forEach((d) => d.classList.toggle('sel', d.dataset.cat === name))
        const list = typeLists[name]
        typeSel.replaceChildren(...(list ?? []).map((t) => el('option', { value: t }, t)))
        if (list) typeSel.value = list.includes(current) ? current : list[0]
        for (const [node, show] of [[decimalsRow, ['Number', 'Currency', 'Accounting', 'Percentage', 'Scientific'].includes(name)],
          [sep.node, name === 'Number'], [negRow, ['Number', 'Currency'].includes(name)], [typeRow, !!list], [codeRow, name === 'Custom']]) node.style.display = show ? '' : 'none'
        refresh()
      }
      const decimalsRow = field('Decimal places:', decimals), negRow = field('Negative numbers:', neg), typeRow = field('Type:', typeSel), codeRow = field('Type:', code)
      for (const n of [decimals, sep.input, neg, typeSel, code]) n.addEventListener('input', refresh)
      panels.Number = el('div', { class: 'sx-row' }, field('Category:', catList),
        el('div', { style: { display: 'flex', flexDirection: 'column', gap: '6px' } }, field('Sample', sample), decimalsRow, sep.node, negRow, typeRow, codeRow))
      setCat(guessCategory(current))
      if (cat === 'Custom') code.value = current
      apply.push(() => { if (pattern() !== (range.getNumberFormat?.() || 'General')) range.setNumberFormat(pattern()) })
    }
    // Alignment
    {
      const hv = { 0: 'general', 1: 'left', 2: 'center', 3: 'right', 4: 'justify', 6: 'distributed' }
      const horiz = el('select', { 'aria-label': 'Horizontal' }, ...[['general', 'General'], ['left', 'Left (Indent)'], ['center', 'Center'], ['right', 'Right'], ['justify', 'Justify'], ['distributed', 'Distributed']].map(([v, t]) => el('option', { value: v }, t)))
      horiz.value = hv[style.ht] ?? 'general'
      const vert = el('select', { 'aria-label': 'Vertical' }, ...[['top', 'Top'], ['middle', 'Center'], ['bottom', 'Bottom']].map(([v, t]) => el('option', { value: v }, t)))
      vert.value = { 1: 'top', 2: 'middle', 3: 'bottom' }[style.vt] ?? 'bottom'
      const wrap = check('Wrap text', style.tb === 3), shrink = check('Shrink to fit', !!range.getShrinkToFit?.()), merge = check('Merge cells', range.isMerged?.())
      // Excel's degrees count counterclockwise; Univer's angle clockwise.
      const deg = el('input', { type: 'number', min: -90, max: 90, value: -(style.tr?.a ?? 0) || 0, 'aria-label': 'Degrees' })
      panels.Alignment = el('div', {}, group('Text alignment', field('Horizontal:', horiz), field('Vertical:', vert)),
        group('Text control', wrap.node, shrink.node, merge.node), field('Orientation (degrees):', deg))
      apply.push(() => {
        range.setHorizontalAlignment(horiz.value === 'general' ? 'normal' : horiz.value)
        range.setVerticalAlignment(vert.value)
        range.setWrap(wrap.input.checked)
        range.setShrinkToFit?.(shrink.input.checked)
        if (merge.input.checked && !range.isMerged()) range.merge()
        else if (!merge.input.checked && range.isMerged()) range.breakApart()
        if (-deg.value !== (style.tr?.a ?? 0)) range.setTextRotation(-deg.value || 0)
      })
    }
    // Font
    {
      const families = ['Calibri', 'Aptos', 'Arial', 'Cambria', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana']
      const fam = el('select', { 'aria-label': 'Font' }, ...families.map((f) => el('option', { value: f }, f)))
      fam.value = style.ff && families.includes(style.ff) ? style.ff : 'Calibri'
      const fstyle = el('select', { 'aria-label': 'Font style' }, ...['Regular', 'Italic', 'Bold', 'Bold Italic'].map((t) => el('option', { value: t }, t)))
      fstyle.value = style.bl && style.it ? 'Bold Italic' : style.bl ? 'Bold' : style.it ? 'Italic' : 'Regular'
      const size = el('input', { type: 'number', min: 1, max: 409, value: style.fs ?? 11, 'aria-label': 'Size' })
      const ulMap = [['none', 'None'], ['single', 'Single'], ['double', 'Double'], ['singleacc', 'Single Accounting'], ['doubleacc', 'Double Accounting']]
      const ul = el('select', { 'aria-label': 'Underline' }, ...ulMap.map(([v, t]) => el('option', { value: v }, t)))
      ul.value = !style.ul?.s ? 'none' : ({ [TextDecoration.DOUBLE]: 'double', [TextDecoration.SINGLE_ACCOUNTING]: 'singleacc', [TextDecoration.DOUBLE_ACCOUNTING]: 'doubleacc' })[style.ul.t] ?? 'single'
      const color = el('input', { type: 'color', value: style.cl?.rgb && /^#[0-9a-f]{6}$/i.test(style.cl.rgb) ? style.cl.rgb : '#000000', 'aria-label': 'Color' })
      const strike = check('Strikethrough', !!style.st?.s), sup = check('Superscript', style.va === 3), sub = check('Subscript', style.va === 2)
      panels.Font = el('div', {}, el('div', { class: 'sx-row' }, field('Font:', fam), field('Font style:', fstyle), field('Size:', size)),
        el('div', { class: 'sx-row' }, field('Underline:', ul), field('Color:', color)), group('Effects', strike.node, sup.node, sub.node))
      apply.push(() => {
        range.setFontFamily(fam.value)
        range.setFontWeight(fstyle.value.includes('Bold') ? 'bold' : 'normal')
        range.setFontStyle(fstyle.value.includes('Italic') ? 'italic' : 'normal')
        range.setFontSize(+size.value || 11)
        range.setFontColor(color.value)
        const t = { single: TextDecoration.SINGLE, double: TextDecoration.DOUBLE, singleacc: TextDecoration.SINGLE_ACCOUNTING, doubleacc: TextDecoration.DOUBLE_ACCOUNTING }[ul.value]
        setStyle(range, 'ul', ul.value === 'none' ? { s: 0 } : { s: 1, t })
        setStyle(range, 'st', { s: strike.input.checked ? 1 : 0 })
        setStyle(range, 'va', sup.input.checked ? 3 : sub.input.checked ? 2 : 1)
      })
    }
    // Border
    {
      let preset = null
      const styleSel = el('select', { 'aria-label': 'Line style' }, ...[['THIN', 'Thin'], ['HAIR', 'Hair'], ['DOTTED', 'Dotted'], ['DASHED', 'Dashed'], ['DASH_DOT', 'Dash dot'], ['DOUBLE', 'Double'], ['MEDIUM', 'Medium'], ['THICK', 'Thick']].map(([v, t]) => el('option', { value: v }, t)))
      const color = el('input', { type: 'color', value: '#000000', 'aria-label': 'Line color' })
      const sides = {}
      const btn = (key, label) => el('button', { class: 'sx-btn', type: 'button', 'aria-pressed': 'false', onClick: (e) => { if (['none', 'outside', 'inside'].includes(key)) preset = key; else sides[key] = !sides[key]; e.target.setAttribute('aria-pressed', String(preset === key || !!sides[key])) } }, label)
      panels.Border = el('div', {}, el('div', { class: 'sx-row' }, field('Style:', styleSel), field('Color:', color)),
        group('Presets', btn('none', 'None'), btn('outside', 'Outline'), btn('inside', 'Inside')),
        group('Border', btn('top', 'Top'), btn('bottom', 'Bottom'), btn('left', 'Left'), btn('right', 'Right'), btn('tlbr', 'Diagonal \\'), btn('bl_tr', 'Diagonal /')))
      apply.push(() => {
        const st = BorderStyleTypes[styleSel.value]
        if (preset === 'none') range.setBorder(BorderType.NONE, BorderStyleTypes.NONE, '#000000')
        else if (preset) range.setBorder(preset === 'outside' ? BorderType.OUTSIDE : BorderType.INSIDE, st, color.value)
        for (const [k, on] of Object.entries(sides)) if (on) range.setBorder({ top: BorderType.TOP, bottom: BorderType.BOTTOM, left: BorderType.LEFT, right: BorderType.RIGHT, tlbr: BorderType.TLBR, bl_tr: BorderType.BLTR }[k], st, color.value)
      })
    }
    // Fill
    {
      const palette = ['', '#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47',
        '#F2F2F2', '#808080', '#D0CECE', '#D6DCE4', '#D9E1F2', '#FCE4D6', '#EDEDED', '#FFF2CC', '#DDEBF7', '#E2EFDA', '#C00000', '#FF0000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0']
      let chosen = style.bg?.rgb ?? ''
      const sw = el('div', { class: 'sx-swatches', role: 'listbox', 'aria-label': 'Background color' }, ...palette.map((c) => el('button', {
        type: 'button', title: c || 'No Color', class: c === chosen ? 'sel' : '', style: { background: c || 'repeating-linear-gradient(45deg,#fff 0 3px,#ddd 3px 6px)' },
        onClick: (e) => { chosen = c; sw.querySelectorAll('button').forEach((b) => b.classList.remove('sel')); e.target.classList.add('sel') } })))
      panels.Fill = el('div', {}, field('Background Color:', sw))
      apply.push(() => { if (chosen !== (style.bg?.rgb ?? '')) chosen ? range.setBackgroundColor(chosen) : setStyle(range, 'bg', null) })
    }
    // Protection (takes effect when the sheet is protected, as in Excel)
    {
      const meta = range.getCustomMetaData?.() ?? {}
      const locked = check('Locked', meta.locked !== false), hidden = check('Hidden', !!meta.hidden)
      panels.Protection = el('div', {}, locked.node, hidden.node,
        el('p', {}, 'Locking cells or hiding formulas has no effect until you protect the worksheet.'))
      apply.push(() => range.setCustomMetaData?.({ ...(range.getCustomMetaData?.() ?? {}), locked: locked.input.checked, hidden: hidden.input.checked }))
    }
    const tabBar = el('div', { class: 'sx-tabs', role: 'tablist' })
    const holder = el('div', {})
    const show = (t) => {
      tabBar.querySelectorAll('button').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === t)))
      holder.replaceChildren(panels[t])
    }
    for (const t of tabs) tabBar.append(el('button', { type: 'button', role: 'tab', 'data-tab': t, onClick: () => show(t) }, t))
    show(startTab)
    openDialog({ title: 'Format Cells', width: 560, body: el('div', {}, tabBar, holder), onOk: () => { for (const f of apply) f() } })
    return true
  }
  const setStyle = (range, type, value) => commands.executeCommand('sheet.command.set-style', {
    unitId: wb().getId(), subUnitId: ws().getSheetId(), range: range.getRange(), style: { type, value } })

  // ---------- Print (Ctrl+P / Ctrl+F2)
  function printPreview() {
    const what = el('select', { 'aria-label': 'Print what' }, el('option', { value: 'sheet' }, 'Print Active Sheets'), el('option', { value: 'book' }, 'Print Entire Workbook'), el('option', { value: 'sel' }, 'Print Selection'))
    const orient = el('select', { 'aria-label': 'Orientation' }, el('option', { value: 'portrait' }, 'Portrait Orientation'), el('option', { value: 'landscape' }, 'Landscape Orientation'))
    const paper = el('select', { 'aria-label': 'Paper size' }, el('option', { value: 'Letter' }, 'Letter'), el('option', { value: 'A4' }, 'A4'), el('option', { value: 'Legal' }, 'Legal'))
    const scale = el('select', { 'aria-label': 'Scaling' }, el('option', { value: 'none' }, 'No Scaling'), el('option', { value: 'fit' }, 'Fit Sheet on One Page'))
    const grid = check('Print gridlines')
    const main = el('main', { 'aria-label': 'Print preview' })
    const render = () => {
      const sheets = what.value === 'book' ? wb().getSheets() : [ws()]
      main.replaceChildren(...sheets.map((s) => {
        const g = what.value === 'sel' ? sel() : usedRangeOf(s)
        const page = el('div', { class: `sx-page ${orient.value}` }, tableFor(s, g, grid.input.checked))
        if (scale.value === 'fit') page.style.zoom = String(Math.min(1, (orient.value === 'landscape' ? 1027 : 698) / Math.max(1, tableWidth(s, g))))
        return page
      }))
    }
    for (const n of [what, orient, paper, scale, grid.input]) n.addEventListener('change', render)
    const close = () => { root.remove(); disable?.dispose?.() }
    const printBtn = el('button', { class: 'sx-btn sx-primary', type: 'button', onClick: async () => {
      const html = `<!doctype html><html><head><meta charset="utf-8"><style>@page{size:${paper.value} ${orient.value};margin:0.7in 0.7in 0.75in 0.7in}
        body{margin:0;font:11pt Calibri,Arial,sans-serif}table{border-collapse:collapse}td{padding:1px 4px;white-space:nowrap}.sx-page{page-break-after:always}</style></head><body>${main.innerHTML}</body></html>`
      await window.sheetly?.print?.(html, { landscape: orient.value === 'landscape', pageSize: paper.value })
    } }, 'Print')
    const disable = shortcutServiceRef?.forceDisable?.()
    const root = el('div', { class: 'sx-print', role: 'dialog', 'aria-label': 'Print' },
      el('aside', {}, el('h2', { style: { margin: 0, fontSize: '20px' } }, 'Print'), printBtn, field('Settings', what), orient, paper, scale, grid.node,
        el('button', { class: 'sx-btn', type: 'button', onClick: close }, 'Back')), main)
    root.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') close() })
    document.body.append(root)
    render()
    printBtn.focus()
    return true
  }
  const usedRangeOf = (s) => {
    let er = 0, ec = 0
    s.getSheet().getCellMatrix().forValue((r, c, x) => { if (x && (x.v != null || x.f != null || x.p)) { er = Math.max(er, r); ec = Math.max(ec, c) } })
    return { startRow: 0, startColumn: 0, endRow: er, endColumn: ec }
  }
  const tableWidth = (s, g) => { let w = 0; for (let c = g.startColumn; c <= g.endColumn; c++) w += s.getColumnWidth(c); return w }
  function tableFor(s, g, gridlines) {
    const table = el('table')
    const colgroup = el('colgroup')
    for (let c = g.startColumn; c <= g.endColumn; c++) colgroup.append(el('col', { style: { width: s.getColumnWidth(c) + 'px' } }))
    table.append(colgroup)
    const sheet = s.getSheet()
    for (let r = g.startRow; r <= g.endRow; r++) {
      if (!sheet.getRowVisible(r)) continue
      const tr = el('tr')
      for (let c = g.startColumn; c <= g.endColumn; c++) {
        if (!sheet.getColVisible(c)) continue
        const range = s.getRange(r, c), st = range.getCellStyleData?.() ?? {}
        const v = range.getValue() // the formatted value, as printed
        tr.append(el('td', { style: {
          fontWeight: st.bl ? 'bold' : '', fontStyle: st.it ? 'italic' : '', textDecoration: [st.ul?.s && 'underline', st.st?.s && 'line-through'].filter(Boolean).join(' '),
          color: st.cl?.rgb ?? '', background: st.bg?.rgb ?? '', fontSize: st.fs ? st.fs + 'pt' : '', fontFamily: st.ff ?? '',
          textAlign: ({ 1: 'left', 2: 'center', 3: 'right' })[st.ht] ?? (typeof v === 'number' ? 'right' : 'left'),
          border: gridlines ? '1px solid #d0d0d0' : '' } }, range.getDisplayValue?.() ?? (v ?? '')))
      }
      table.append(tr)
    }
    return table
  }

  // ---------- Help (F1)
  function help(shortcutDocs) {
    const existing = document.querySelector('.sx-pane[aria-label="Help"]')
    if (existing) { existing.remove(); return true }
    const q = el('input', { type: 'search', placeholder: 'Search shortcuts', 'aria-label': 'Search shortcuts', style: { margin: '8px 12px' } })
    const body = el('tbody')
    const fill = () => body.replaceChildren(...shortcutDocs.filter(([k, d]) => !q.value || (k + d).toLowerCase().includes(q.value.toLowerCase()))
      .map(([k, d]) => el('tr', {}, el('td', {}, k), el('td', {}, d))))
    q.addEventListener('input', fill)
    const close = () => { pane.remove(); disable?.dispose?.() }
    const disable = shortcutServiceRef?.forceDisable?.()
    const pane = el('div', { class: 'sx-pane', role: 'dialog', 'aria-label': 'Help' },
      el('div', { class: 'sx-title' }, el('span', {}, 'Help — Keyboard shortcuts'), el('button', { class: 'sx-close', type: 'button', 'aria-label': 'Close', onClick: close }, '×')),
      q, el('div', { style: { overflow: 'auto', flex: 1 } }, el('table', {}, body)))
    pane.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape' || e.key === 'F1') { e.preventDefault(); close() } })
    fill()
    document.body.append(pane)
    q.focus()
    return true
  }

  let shortcutServiceRef = null
  return {
    setShortcutService: (s) => { shortcutServiceRef = s },
    goTo, goToSpecial, goToReference, runGoToSpecial, insertCells, deleteCells, createNames, createNamesFrom, createTable, makeTable,
    pasteSpecial, runPasteSpecial, pasteLink, formatCells, printPreview, help, currentRegion, usedRange, select, getLastCopy: () => lastCopy,
  }
}

const dec = (n) => (n > 0 ? '.' + '0'.repeat(n) : '')
function withNeg(base, neg) {
  return [base, `${base};[Red]${base}`, `${base}_);(${base})`, `${base}_);[Red](${base})`][neg] ?? base
}
function guessCategory(p) {
  if (!p || p === 'General') return 'General'
  if (p === '@') return 'Text'
  if (/E\+/.test(p)) return 'Scientific'
  if (/%$/.test(p)) return 'Percentage'
  if (/^_\(\$\*/.test(p)) return 'Accounting'
  if (/^\$/.test(p)) return 'Currency'
  if (/^[#0,.]+(;|_|$)/.test(p)) return 'Number'
  return 'Custom'
}
function alertBox(message) {
  openDialog({ title: 'Sheetly', width: 320, body: el('p', { style: { margin: 0 } }, message), cancelLabel: 'Close' })
}
