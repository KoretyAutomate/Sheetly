// PivotTables: Insert > PivotTable (Alt N V T; Office 2003 Alt D P), the Create PivotTable dialog, the PivotTable
// Fields task pane (field list with check boxes, Filters / Columns / Rows / Values areas with drag and drop and
// each field's menu), Value Field Settings, report filters, Refresh (Alt+F5) and Refresh All (Ctrl+Alt+F5).
// The report is written into cells in Excel's Compact Form with the default PivotStyleLight16 look. Like Excel it
// does not follow the source until it is refreshed. Definitions live in the report sheet's custom data
// (custom.sheetlyPivots), so they are saved with the workbook; an .xlsx keeps the report's values.
import { BorderStyleTypes, BorderType, CommandType, ICommandService, IUndoRedoService, IUniverInstanceService } from '@univerjs/core'
import { el, field, openDialog, radios } from './ui/dialog.js'
import { BLANK, FN_LABEL, FN_MENU, FUNCS, defaultFn, filterCells, getPivotData, isNumericField, layout, readSource, valueName } from './pivot-model.js'
import { parseRange, rangeA1 } from './xlsx-objects.js'

// PivotStyleLight16: light blue header band, bold headers and totals, blue rules.
const HEADER_BG = '#DDEBF7', RULE = '#9BC2E6'
const EMPTY_EXTENT = { rows: 18, cols: 3 } // the empty report's placeholder (A3:C20)
// The next frame, or a short timer when the window is in the background (Chromium holds animation frames back there).
const soon = (fn) => { let done = false; const run = () => { if (!done) { done = true; fn() } }; requestAnimationFrame(run); setTimeout(run, 60) }
const AREAS = [['filters', 'Filters'], ['cols', 'Columns'], ['rows', 'Rows'], ['values', 'Values']]

export function createPivots(api, { dialogs, refs }) {
  const commands = api._injector.get(ICommandService)
  const wb = () => api.getActiveWorkbook()
  const ws = () => wb().getActiveSheet()
  const listOf = (sheet) => sheet.getSheet().getCustomMetadata?.()?.sheetlyPivots ?? []
  const storeIn = (w, list) => {
    const custom = { ...(w.getCustomMetadata?.() ?? {}) }
    if (list.length) custom.sheetlyPivots = list; else delete custom.sheetlyPivots
    w.setCustomMetadata(Object.keys(custom).length ? custom : undefined)
  }
  const store = (sheet, list) => storeIn(sheet.getSheet(), list)
  // Definitions change through a mutation with the previous list as its undo; a whole change (cells, formats and
  // definition) is one undo step (atomic, below).
  const SET = 'sheetly.mutation.set-pivots'
  commands.registerCommand({ id: SET, type: CommandType.MUTATION, handler: (_a, p) => {
    const w = api._injector.get(IUniverInstanceService).getUnit(p.unitId)?.getSheetBySheetId(p.subUnitId)
    if (!w) return false
    storeIn(w, p.list)
    return true
  } })
  const undoRedo = api._injector.get(IUndoRedoService)
  function commitList(sheet, list) {
    const unitId = wb().getId(), subUnitId = sheet.getSheetId()
    const before = structuredClone(listOf(sheet)), after = structuredClone(list)
    commands.syncExecuteCommand(SET, { unitId, subUnitId, list: after })
    undoRedo.pushUndoRedo({ unitID: unitId, undoMutations: [{ id: SET, params: { unitId, subUnitId, list: before } }], redoMutations: [{ id: SET, params: { unitId, subUnitId, list: after } }] })
  }
  /** Runs fn as a single undo step (everything it changes is undone together). */
  function atomic(fn) {
    const g = undoRedo.beginUndoRedoGroup(wb().getId(), `pivot-${Date.now()}-${Math.random()}`, 'append')
    try { return fn() } finally { g.dispose() }
  }
  const allPivots = () => wb().getSheets().flatMap((s) => listOf(s).map((p) => ({ sheet: s, p })))
  // The source and the report's place are hidden defined names (refs.js): Univer moves them with inserted and
  // deleted rows and columns, in the same undo step. Stored coordinates are the fallback.
  // A table source is read from the table as it is at the time (rows added to it are included on Refresh).
  const sourceAt = (p) => {
    const t = p.sourceTable ? tableData(p.sourceTable, p.sourceTableId) : null
    if (t) return { sheetId: t.sheet.getSheetId(), range: t.range }
    return p.sourceRef ? refs?.resolve(p.sourceRef) : { sheetId: p.source.sheetId, range: p.source.range }
  }
  const anchorAt = (p) => { const a = p.anchorRef ? refs?.resolve(p.anchorRef) : null; return a ? { row: a.range.startRow, col: a.range.startColumn } : p.anchor }
  // Like Excel's PivotCache: the records a report was built from are kept with it (p.cache). Layout changes and
  // GETPIVOTDATA use them; only creating and refreshing read the source again.
  const readLive = (p) => {
    const at = sourceAt(p)
    const s = at && wb().getSheetBySheetId(at.sheetId)
    return s ? readSource((r, c) => { const v = s.getRange(r, c).getCalculatedValue(); return v === '' ? null : v }, at.range) : null
  }
  const sourceOf = (p) => p.cache ?? readLive(p)
  /** The cells the report wrote (filters included), or null before it has written any. */
  // The written area is a hidden name too, so rows inserted or deleted inside the report grow or shrink it.
  const written = (p) => {
    const e = p.extent
    if (!e?.written) return null
    if (p.writtenRef) {
      const w = refs?.resolve(p.writtenRef)
      if (!w) return null
      return { row: w.range.startRow, col: w.range.startColumn, rows: w.range.endRow - w.range.startRow + 1, cols: w.range.endColumn - w.range.startColumn + 1 }
    }
    const a = anchorAt(p)
    return { row: a.row + e.dRow - e.filterRows, col: a.col, rows: e.rows + e.filterRows, cols: e.cols }
  }
  // An empty report shows Excel's placeholder area (A3:C20) for the fields pane.
  const footprint = (p) => written(p) ?? { ...anchorAt(p), rows: EMPTY_EXTENT.rows, cols: EMPTY_EXTENT.cols }
  const inside = (p, r, c) => { const f = footprint(p); return r >= f.row && r < f.row + f.rows && c >= f.col && c < f.col + f.cols }
  /** The PivotTable the active cell is in (on the active sheet). */
  function active() {
    const g = ws().getActiveRange()?.getRange()
    if (!g) return null
    return listOf(ws()).find((p) => inside(p, g.startRow, g.startColumn)) ?? null
  }
  const message = (text) => { openDialog({ title: 'Sheetly', width: 380, body: el('p', { style: { margin: 0 } }, text), cancelLabel: 'Close' }); return true }

  // ---- writing a report: only the cells it wrote last time are cleared; other data in the way needs consent.
  function render(sheet, p, { refresh = false, force = false } = {}) {
    if (refresh || !p.cache) p.cache = readLive(p)
    const src = p.cache
    const old = written(p)
    const { row, col } = anchorAt(p)
    const empty = !src || (!p.rows.length && !p.cols.length && !p.values.length && !p.filters.length)
    const L = empty ? null : layout(p, src)
    const filters = empty ? [] : filterCells(p)
    // Report filters sit above the report with one blank row between; the report moves down to make room.
    const top = Math.max(row, filters.length ? filters.length + 1 : 0)
    const hasGrid = L && (p.rows.length || p.cols.length || p.values.length)
    // Report filters take two columns (name, item), so the report owns at least two when it has filters.
    const width = Math.max(hasGrid ? L.grid.reduce((m, r) => Math.max(m, r.length), 0) : 0, filters.length ? 2 : 0)
    const rows = hasGrid ? L.grid.length : 0
    const filterRows = filters.length ? filters.length + 1 : 0
    const next = width ? { row: top - filterRows, col, rows: rows + filterRows, cols: width } : null
    const ownedBefore = (r, c) => old && r >= old.row && r < old.row + old.rows && c >= old.col && c < old.col + old.cols
    if (next && !force) {
      for (let r = next.row; r < next.row + next.rows; r++) for (let c = next.col; c < next.col + next.cols; c++) {
        // a formula counts as content even when its result is empty (="")
        const cell = sheet.getRange(r, c), v = cell.getCalculatedValue()
        if (((v != null && v !== '') || cell.getFormula()) && !ownedBefore(r, c)) return { conflict: true }
      }
    }
    if (old) sheet.getRange(old.row, old.col, old.rows, old.cols).clear()
    if (!next) { p.extent = null; return p }
    filters.forEach(([name, value], i) => {
      const r = top - filters.length - 1 + i
      sheet.getRange(r, col, 1, 2).setValues([[typed(name), typed(value)]])
      sheet.getRange(r, col).setFontWeight('bold')
    })
    p.extent = { written: true, dRow: top - row, rows, cols: width, filterRows }
    const area = { startRow: next.row, endRow: next.row + next.rows - 1, startColumn: next.col, endColumn: next.col + next.cols - 1 }
    // A new name each time: undo brings back the earlier definition with its own name, still on the earlier area.
    p.writtenRef = refs?.track(sheet.getSheetId(), area) ?? undefined
    if (!hasGrid) return p
    const values = L.grid.map((r) => [...r.map((c) => typed(c.v ?? null)), ...Array(width - r.length).fill(null)])
    sheet.getRange(top, col, values.length, width).setValues(values)
    // Excel's default look: header rows on a light blue band in bold, subtotal and Grand Total rows in bold.
    const headerRows = L.grid.findIndex((r) => r.some((c) => ['rowLabel', 'subtotalLabel', 'value', 'grandLabel', 'grand'].includes(c.kind) && c.kind !== 'colHeader'))
    if (headerRows > 0) { const h = sheet.getRange(top, col, headerRows, width); h.setBackground(HEADER_BG); h.setFontWeight('bold') }
    L.grid.forEach((r, i) => { if (['grandLabel', 'subtotalLabel'].includes(r[0]?.kind) || (i === L.grid.length - 1 && r[0]?.kind === 'grand')) sheet.getRange(top + i, col, 1, width).setFontWeight('bold') })
    if (headerRows > 0) sheet.getRange(top + headerRows - 1, col, 1, width).setBorder(BorderType.BOTTOM, BorderStyleTypes.THIN, RULE)
    const last = L.grid.length - 1
    if (L.grid[last]?.[0]?.kind === 'grandLabel') sheet.getRange(top + last, col, 1, width).setBorder(BorderType.TOP, BorderStyleTypes.THIN, RULE)
    // Compact form indents inner row items.
    L.grid.forEach((r, i) => { if (r[0]?.indent) indent(sheet, top + i, col, r[0].indent) })
    return p
  }
  // Report cells are written with their type: text stays text ("001", "TRUE", "=1+2" are labels, not a number,
  // a logical or a formula), logicals stay logicals.
  function typed(v) {
    if (typeof v === 'string') return { v, t: 1 }
    if (typeof v === 'boolean') return { v: v ? 1 : 0, t: 3 }
    return v
  }
  // Indentation as left padding (Excel indents compact-form items by one level each).
  function indent(sheet, r, c, level) {
    const unitId = wb().getId(), subUnitId = sheet.getSheetId()
    commands.syncExecuteCommand('sheet.command.set-style', { unitId, subUnitId, range: { startRow: r, endRow: r, startColumn: c, endColumn: c },
      style: { type: 'pd', value: { l: 4 + 12 * level, r: 2, t: 0, b: 0 } } })
  }
  function save(sheet, p) {
    const list = listOf(sheet).filter((x) => x.id !== p.id)
    commitList(sheet, [...list, structuredClone(p)])
  }
  // Excel asks before a report grows over other data.
  const askReplace = (sheet, then) => openDialog({ title: 'Sheetly', width: 400, okLabel: 'OK',
    body: el('p', { style: { margin: 0 } }, `There's already data in [${sheet.getSheetName()}]. Do you want to replace it?`), onOk: then })
  function update(sheet, p, opts = {}) {
    const go = (o) => atomic(() => { const r = render(sheet, p, o); if (!r?.conflict) save(sheet, p); return r })
    if (go(opts)?.conflict) { askReplace(sheet, () => { go({ ...opts, force: true }); showPane() }); return true }
    showPane(); return true
  }

  // ---- Insert > PivotTable
  function createDialog() {
    const g0 = ws().getActiveRange().getRange()
    const g = g0.startRow === g0.endRow && g0.startColumn === g0.endColumn ? dialogs.currentRegion(g0.startRow, g0.startColumn) : g0
    const sheetName = ws().getSheetName()
    const q = (n) => (/^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`)
    const abs = (a1) => a1.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')
    const rangeInput = el('input', { type: 'text', value: `${q(sheetName)}!${abs(rangeA1(g))}`, style: { width: '260px' } })
    const where = radios('pivot-where', [['new', '&New Worksheet'], ['existing', '&Existing Worksheet']], 'new')
    const loc = el('input', { type: 'text', value: '', style: { width: '200px' } })
    openDialog({ title: 'Create PivotTable', width: 420, initialFocus: rangeInput,
      body: el('div', {}, el('div', { style: { fontWeight: 600 } }, 'Choose the data that you want to analyze'), field('Table/Range:', rangeInput),
        el('div', { style: { fontWeight: 600, marginTop: '8px' } }, 'Choose where you want the PivotTable report to be placed'), ...where.nodes, field('Location:', loc)),
      onOk: () => create(rangeInput.value, where.value, loc.value) })
    return true
  }
  const sheetByName = (name) => wb().getSheets().find((s) => s.getSheetName().toLowerCase() === name.toLowerCase())
  /** A table's data as it is now: its header row and records, without a totals row (read from its definition). */
  // By id first, so a renamed table is still found; by name for a definition saved without one.
  function tableData(name, id = null) {
    const t = (id ? wb().getTableInfo?.(id) : null) ?? wb().getTableInfoByName?.(name)
    if (!t?.range) return null
    const sheet = wb().getSheetBySheetId(t.subUnitId)
    let footer = false
    try { footer = !!JSON.parse(wb().save().resources?.find((r) => r.name === 'SHEET_TABLE_PLUGIN')?.data ?? '{}')?.[t.subUnitId]?.tables?.find((x) => x.id === t.id)?.options?.showFooter } catch { footer = false }
    return sheet ? { sheet, range: { ...t.range, endRow: t.range.endRow - (footer ? 1 : 0) }, table: t.name, tableId: t.id } : null
  }
  function refOf(text) {
    // A table's name stands for its data with the header row (as Excel takes it), without a totals row.
    const t = /^\s*[\p{L}_][\p{L}\p{N}_.]*\s*$/u.test(text) && !parseRange(text) ? tableData(text.trim()) : null
    if (t) return t
    const m = /^\s*(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(text)
    const sheet = m ? sheetByName((m[1] ?? m[2]).replace(/''/g, "'")) : ws()
    const g = parseRange(m ? m[3] : text)
    return sheet && g ? { sheet, range: g } : null
  }
  // Creating a PivotTable (new sheet, report, definition) is one undo step.
  const create = (...a) => atomic(() => createNow(...a))
  function createNow(rangeText, where, locText) {
    const src = refOf(rangeText)
    if (!src || src.range.endRow <= src.range.startRow) return message('The reference is not valid.') && false
    const data = readSource((r, c) => { const v = src.sheet.getRange(r, c).getCalculatedValue(); return v === '' ? null : v }, src.range)
    if (!data) return message('The PivotTable field name is not valid. To create a PivotTable report, you must use data that is organized as a list with labeled columns.') && false
    let target, anchor
    if (where === 'existing') {
      const t = refOf(locText)
      if (!t) return message('The destination reference is not valid.') && false
      target = t.sheet; anchor = { row: t.range.startRow, col: t.range.startColumn }
    } else {
      // A new sheet, to the left of the source sheet; the report starts at A3.
      const names = new Set(wb().getSheets().map((s) => s.getSheetName()))
      let n = 1
      while (names.has(`Sheet${n}`)) n++
      const index = wb().getSheets().findIndex((s) => s.getSheetId() === src.sheet.getSheetId())
      target = wb().insertSheet(`Sheet${n}`, { index })
      anchor = { row: 2, col: 0 }
    }
    const used = new Set(allPivots().map(({ p }) => p.name))
    let k = 1
    while (used.has(`PivotTable${k}`)) k++
    const p = { id: `pt${Date.now().toString(36)}${k}`, name: `PivotTable${k}`, source: { sheetId: src.sheet.getSheetId(), range: src.range },
      anchor, rows: [], cols: [], values: [], filters: [],
      sourceRef: refs?.track(src.sheet.getSheetId(), src.range) ?? undefined, ...(src.table ? { sourceTable: src.table, sourceTableId: src.tableId } : {}),
      anchorRef: refs?.track(target.getSheetId(), { startRow: anchor.row, endRow: anchor.row, startColumn: anchor.col, endColumn: anchor.col }) ?? undefined }
    wb().setActiveSheet(target)
    render(target, p, { refresh: true, force: true })
    save(target, p)
    target.getRange(anchor.row, anchor.col).activate()
    setTimeout(showPane, 0)
    return true
  }

  // ---- changing the layout
  const current = () => { const p = active(); return p ? { sheet: ws(), p: structuredClone(p) } : null }
  function place(fieldName, area, at) {
    const c = current(); if (!c) return false
    const { sheet, p } = c, src = sourceOf(p)
    for (const [a] of AREAS) p[a] = p[a].filter((x) => (a === 'values' ? true : (a === 'filters' ? x.field : x) !== fieldName))
    if (area === 'values') p.values.splice(at ?? p.values.length, 0, { field: fieldName, fn: defaultFn(src, fieldName) })
    else if (area === 'filters') p.filters.splice(at ?? p.filters.length, 0, { field: fieldName, items: null })
    else p[area].splice(at ?? p[area].length, 0, fieldName)
    return update(sheet, p)
  }
  function removeAt(area, i) { const c = current(); if (!c) return false; c.p[area].splice(i, 1); return update(c.sheet, c.p) }
  function moveWithin(area, i, d) {
    const c = current(); if (!c) return false
    const list = c.p[area], j = i + d
    if (j < 0 || j >= list.length) return true
    ;[list[i], list[j]] = [list[j], list[i]]
    return update(c.sheet, c.p)
  }
  function moveTo(area, i, to) {
    const c = current(); if (!c) return false
    const item = c.p[area][i], name = area === 'values' ? item.field : area === 'filters' ? item.field : item
    c.p[area].splice(i, 1)
    if (to === 'values') c.p.values.push({ field: name, fn: defaultFn(sourceOf(c.p), name) })
    else {
      for (const a of ['rows', 'cols']) c.p[a] = c.p[a].filter((x) => x !== name)
      c.p.filters = c.p.filters.filter((x) => x.field !== name)
      if (to === 'filters') c.p.filters.push({ field: name, items: null }); else c.p[to].push(name)
    }
    return update(c.sheet, c.p)
  }
  /** Check a field in the list: numbers go to Values, other fields to Rows (Excel's choice); uncheck removes it everywhere. */
  function toggleField(name, on) {
    const c = current(); if (!c) return false
    if (!on) {
      for (const a of ['rows', 'cols']) c.p[a] = c.p[a].filter((x) => x !== name)
      c.p.filters = c.p.filters.filter((x) => x.field !== name)
      c.p.values = c.p.values.filter((x) => x.field !== name)
      return update(c.sheet, c.p)
    }
    return place(name, isNumericField(sourceOf(c.p), name) ? 'values' : 'rows')
  }
  function setFilter(i, items) { const c = current(); if (!c) return false; c.p.filters[i].items = items; return update(c.sheet, c.p) }
  function setValueField(i, fn, name) {
    const c = current(); if (!c) return false
    const v = c.p.values[i]
    v.fn = fn
    v.name = name && name !== `${FN_LABEL[fn]} of ${v.field}` ? name : undefined
    return update(c.sheet, c.p)
  }
  function valueFieldSettings(i) {
    const c = current(); if (!c) return message('Select a cell in the PivotTable first.')
    const v = c.p.values[i ?? 0]
    if (!v) return message('The PivotTable has no value field.')
    const name = el('input', { type: 'text', value: valueName(v), style: { width: '220px' } })
    const list = el('select', { size: 8, style: { width: '220px' }, 'aria-label': 'Summarize value field by' }, ...FUNCS.map((f) => el('option', { value: f, selected: f === v.fn }, FN_MENU[f])))
    list.addEventListener('change', () => { if (!v.name) name.value = `${FN_LABEL[list.value]} of ${v.field}` })
    openDialog({ title: 'Value Field Settings', width: 320, initialFocus: list,
      body: el('div', {}, el('div', {}, `Source Name: ${v.field}`), field('Custom Name:', name), field('Summarize value field by', list)),
      onOk: () => { setValueField(i ?? 0, list.value, name.value) } })
    return true
  }
  function refresh(all) {
    const list = all ? allPivots() : (() => { const p = active(); return p ? [{ sheet: ws(), p }] : [] })()
    for (const { sheet, p } of list) update(sheet, structuredClone(p), { refresh: true })
    showPane()
    return true
  }

  // ---- PivotTable Fields task pane (shown while the active cell is in a PivotTable)
  const pane = el('aside', { class: 'sx-task-pane sx-pivot-pane', role: 'complementary', 'aria-label': 'PivotTable Fields', hidden: true })
  document.body.append(pane)
  let drag = null
  function showPane() {
    const p = active()
    if (!p) { pane.hidden = true; return }
    const src = sourceOf(p)
    if (!src) { pane.hidden = true; return }
    pane.hidden = false
    const used = new Set([...p.rows, ...p.cols, ...p.filters.map((f) => f.field), ...p.values.map((v) => v.field)])
    const fields = el('div', { class: 'sx-pivot-fields', role: 'list' }, ...src.fields.map((f) => {
      const box = el('input', { type: 'checkbox', checked: used.has(f), onChange: (e) => toggleField(f, e.target.checked) })
      const item = el('label', { class: 'sx-pivot-field', role: 'listitem', draggable: 'true' }, box, ' ', f)
      item.addEventListener('dragstart', () => { drag = { field: f } })
      return item
    }))
    const area = (key, title) => {
      const items = p[key].map((x, i) => {
        const text = key === 'values' ? valueName(x) : key === 'filters' ? x.field : x
        const menu = el('select', { class: 'sx-pivot-menu', 'aria-label': `${text} options`, onChange: (e) => {
          const a = e.target.value; e.target.value = ''
          if (a === 'up') moveWithin(key, i, -1); else if (a === 'down') moveWithin(key, i, 1); else if (a === 'remove') removeAt(key, i)
          else if (a === 'settings') valueFieldSettings(i); else if (a) moveTo(key, i, a)
        } }, el('option', { value: '' }, '▾'), el('option', { value: 'up' }, 'Move Up'), el('option', { value: 'down' }, 'Move Down'),
        ...AREAS.filter(([a]) => a !== key).map(([a, t]) => el('option', { value: a }, `Move to ${{ filters: 'Report Filter', cols: 'Column Labels', rows: 'Row Labels', values: 'Values' }[a] ?? t}`)),
        el('option', { value: 'remove' }, 'Remove Field'), ...(key === 'values' ? [el('option', { value: 'settings' }, 'Value Field Settings…')] : []))
        const row = el('div', { class: 'sx-pivot-item', draggable: 'true', 'data-area': key }, el('span', {}, text), menu)
        row.addEventListener('dragstart', (e) => { e.stopPropagation(); drag = { from: key, index: i, field: key === 'values' || key === 'filters' ? x.field : x } })
        if (key === 'filters') {
          // Items keep their type (the number 1 and the text "1" are different items); options refer to them by index.
          const items2 = [...new Map(src.records.map((r) => r[src.fields.indexOf(x.field)] ?? BLANK).map((v) => (v === '' ? BLANK : v)).map((v) => [`${typeof v}:${v}`, v])).values()]
          const pick = el('select', { 'aria-label': `${x.field} filter`, onChange: (e) => setFilter(i, e.target.value === '' ? null : [items2[+e.target.value]]) },
            el('option', { value: '' }, '(All)'), ...items2.map((it, k) => el('option', { value: String(k), 'data-item': String(it), selected: x.items?.length === 1 && x.items[0] === it }, String(it))))
          row.append(pick)
        }
        return row
      })
      const box = el('div', { class: 'sx-pivot-area', 'data-area': key }, el('div', { class: 'sx-pivot-area-title' }, title), ...items)
      box.addEventListener('dragover', (e) => e.preventDefault())
      box.addEventListener('drop', (e) => {
        e.preventDefault()
        const d = drag; drag = null
        if (!d) return
        if (d.from) { if (d.from === key) return; moveTo(d.from, d.index, key) } else place(d.field, key)
      })
      return box
    }
    pane.replaceChildren(el('div', { class: 'sx-pivot-title' }, 'PivotTable Fields', el('button', { type: 'button', class: 'sx-close', 'aria-label': 'Close', onClick: () => { pane.hidden = true } }, '×')),
      el('div', { class: 'sx-pivot-hint' }, 'Choose fields to add to report:'), fields,
      el('div', { class: 'sx-pivot-hint' }, 'Drag fields between areas below:'), el('div', { class: 'sx-pivot-areas' }, ...AREAS.map(([k, t]) => area(k, t))))
  }
  commands.onCommandExecuted((cmd) => { if (cmd.id === 'sheet.operation.set-selections' || cmd.id.startsWith('sheet.operation.set-worksheet-active')) soon(showPane) })

  /** A PivotTable's place is on its own sheet. A sheet holding PivotTables placed by references to another sheet
   *  (a copied sheet, also on redo) gets them pointed at itself, with the source too when it was on that sheet. */
  function fixLocal() {
    for (const sheet of wb()?.getSheets?.() ?? []) {
      const list = listOf(sheet)
      let changed = false
      const next = list.map((p) => {
        const at = p.anchorRef ? refs?.resolve(p.anchorRef) : null
        if (!at || at.sheetId === sheet.getSheetId()) return p
        changed = true
        const here = (name) => { const r = name ? refs?.resolve(name) : null; return r && r.sheetId === at.sheetId ? refs?.track(sheet.getSheetId(), r.range) ?? name : name }
        return { ...p, id: `${p.id}c`, anchorRef: here(p.anchorRef), writtenRef: here(p.writtenRef), sourceRef: here(p.sourceRef) }
      })
      if (changed) store(sheet, next)
    }
  }

  /** GETPIVOTDATA: the PivotTable at (sheetId, row, col) answers for data_field and field/item pairs. */
  function lookup(sheetId, row, col, dataField, pairs) {
    const sheet = wb()?.getSheetBySheetId(sheetId)
    const p = sheet && listOf(sheet).find((x) => inside(x, row, col))
    if (!p) return { error: '#REF!' }
    const src = sourceOf(p)
    return src ? getPivotData(p, src, dataField, pairs) : { error: '#REF!' }
  }

  return { fixLocal, createDialog, create, active, place, toggleField, removeAt, moveTo, moveWithin, setFilter, setValueField, valueFieldSettings, refresh, lookup, showPane,
    list: () => allPivots().map(({ sheet, p }) => ({ sheet: sheet.getSheetName(), ...p })) }
}
