// Excel outline (row/column grouping): Group / Ungroup (Alt+Shift+Right / Left), collapse and expand with the
// +/- buttons, level buttons, Show / Hide Detail, Clear Outline, Ctrl+8 to show or hide the outline symbols.
//
// Levels live in Univer's own row/column data (custom.outlineLevel, 1..7), so they move with inserted and
// deleted rows and are saved with the workbook. The summary row/column (below / right of the detail by default,
// above / left when the sheet's outline settings say so) carries custom.outlineCollapsed: a bit per level whose group directly before it was collapsed by
// the user, so that expanding an outer group leaves the inner collapsed groups collapsed, as in Excel.
// All changes go through Univer mutations and form one undo step per user action.
import { CommandType, ICommandService, IUndoRedoService, IUniverInstanceService, sequenceExecute } from '@univerjs/core'
import { IRenderManagerService, SheetExtension, Vector2 } from '@univerjs/engine-render'
import { SheetInterceptorService } from '@univerjs/sheets'
import { HEADER_UNHIDE_RANGE_VISIBLE_CHECK, HeaderUnhideRenderController, SheetSkeletonManagerService } from '@univerjs/sheets-ui'
import { button, check, el, group as fieldset, openDialog, radios } from './ui/dialog.js'
import { MAX_LEVEL, outlineGroups, outlineSettings, summaryIndex } from './outline-model.js'

export { outlineGroups }

const STEP = 16 // px per gutter slot
const BOX = 9
const LINE = '#5f5f5f', BG = '#f3f3f3', TEXT = '#333333'
const FONT = '13px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif'
const TRANSPARENT = 'rgba(0,0,0,0)'
// Univer's default header sizes; the gutter is added on top of these.
const BASE_ROW_HEADER = 46, BASE_COL_HEADER = 20

const columnLetters = (c) => { let s = ''; for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s }

export function createOutline(api) {
  const injector = api._injector
  const commands = injector.get(ICommandService)
  const undoRedo = injector.get(IUndoRedoService)
  const instances = injector.get(IUniverInstanceService)
  const renders = injector.get(IRenderManagerService)
  // Per-sheet settings (worksheet custom.sheetlyOutline): summary direction and Ctrl+8 symbols; saved with the file.
  const settingsOf = (w) => outlineSettings(w.getCustomMetadata?.()?.sheetlyOutline)
  const writeSettings = (w, v) => {
    const custom = { ...(w.getCustomMetadata?.() ?? {}) }, next = outlineSettings(v)
    if (next.isDefault) delete custom.sheetlyOutline; else custom.sheetlyOutline = next.value
    w.setCustomMetadata(Object.keys(custom).length ? custom : undefined)
  }
  const after = (w, axis) => (axis === 'row' ? settingsOf(w).summaryBelow : settingsOf(w).summaryRight)
  const sumOf = (w, axis, g) => summaryIndex(g, after(w, axis))

  const wb = () => api.getActiveWorkbook()
  const ws = () => wb().getActiveSheet()
  const AX = {
    row: { manager: (w) => w.getRowManager(), data: (w) => w.getRowManager().getRowData(), mutation: 'sheet.mutation.set-row-data', key: 'rowData',
      hide: 'sheet.mutation.set-row-hidden', show: 'sheet.mutation.set-row-visible', range: (s, e) => ({ startRow: s, endRow: e, startColumn: 0, endColumn: 0 }) },
    col: { manager: (w) => w.getColumnManager(), data: (w) => w.getColumnManager().getColumnData(), mutation: 'sheet.mutation.set-col-data', key: 'columnData',
      hide: 'sheet.mutation.set-col-hidden', show: 'sheet.mutation.set-col-visible', range: (s, e) => ({ startRow: 0, endRow: 0, startColumn: s, endColumn: e }) },
  }
  const item = (w, axis, i) => (axis === 'row' ? w.getRowManager().getRow(i) : w.getColumnManager().getColumn(i))
  const levelOf = (w, axis, i) => item(w, axis, i)?.custom?.outlineLevel ?? 0
  const flagsOf = (w, axis, i) => item(w, axis, i)?.custom?.outlineCollapsed ?? 0
  const isHidden = (w, axis, i) => item(w, axis, i)?.hd === 1
  const levelMap = (w, axis) => {
    const m = new Map()
    for (const [k, v] of Object.entries(AX[axis].data(w) ?? {})) { const L = v?.custom?.outlineLevel; if (L > 0) m.set(Number(k), L) }
    return m
  }
  const worksheetOf = (unitId, subUnitId) => instances.getUnit(unitId)?.getSheetBySheetId(subUnitId)

  /** Outline state of a worksheet (Univer Worksheet model). */
  function state(w) {
    const rows = outlineGroups(levelMap(w, 'row')), cols = outlineGroups(levelMap(w, 'col'))
    const max = (g) => g.reduce((m, x) => Math.max(m, x.level), 0)
    return { rows, cols, rowLevels: max(rows), colLevels: max(cols) }
  }
  const collapsed = (w, axis, g) => { for (let i = g.start; i <= g.end; i++) if (!isHidden(w, axis, i)) return false; return true }

  // ---- mutations: one undo step per action
  function apply(w, axis, changes, hide, show) {
    const unitId = wb().getId(), subUnitId = w.getSheetId(), A = AX[axis]
    const redo = [], undo = []
    if (changes.size) {
      const after = {}, before = {}
      for (const [i, { level, flags }] of changes) {
        const old = item(w, axis, i)?.custom
        const next = { ...(old ?? {}) }
        delete next.outlineLevel; delete next.outlineCollapsed
        if (level > 0) next.outlineLevel = level
        if (flags) next.outlineCollapsed = flags
        after[i] = { custom: Object.keys(next).length ? next : undefined }
        before[i] = { custom: old ? { ...old } : undefined }
      }
      redo.push({ id: A.mutation, params: { unitId, subUnitId, [A.key]: after } })
      undo.push({ id: A.mutation, params: { unitId, subUnitId, [A.key]: before } })
    }
    const runs = (list) => {
      const r = [], s = [...new Set(list)].sort((a, b) => a - b)
      for (const i of s) { const last = r[r.length - 1]; if (last && last[1] === i - 1) last[1] = i; else r.push([i, i]) }
      return r.map(([a, b]) => A.range(a, b))
    }
    const toHide = hide.filter((i) => !isHidden(w, axis, i)), toShow = show.filter((i) => isHidden(w, axis, i))
    if (toHide.length) { redo.push({ id: A.hide, params: { unitId, subUnitId, ranges: runs(toHide) } }); undo.push({ id: A.show, params: { unitId, subUnitId, ranges: runs(toHide) } }) }
    if (toShow.length) { redo.push({ id: A.show, params: { unitId, subUnitId, ranges: runs(toShow) } }); undo.push({ id: A.hide, params: { unitId, subUnitId, ranges: runs(toShow) } }) }
    if (!redo.length) return true
    const ok = sequenceExecute(redo, commands).result
    if (ok) undoRedo.pushUndoRedo({ unitID: unitId, undoMutations: undo.reverse(), redoMutations: redo })
    return ok
  }

  // Keep a summary-row collapse bit only while a group of that level still ends right before it.
  function cleanFlags(w, axis, levels, changes) {
    const ends = new Set(outlineGroups(levels).map((g) => `${sumOf(w, axis, g)}:${g.level}`))
    const candidates = new Set([...Object.keys(AX[axis].data(w) ?? {}).map(Number).filter((i) => flagsOf(w, axis, i))])
    for (const i of candidates) {
      const f = changes.get(i)?.flags ?? flagsOf(w, axis, i)
      let keep = 0
      for (let k = 1; k <= MAX_LEVEL; k++) if (f & (1 << k) && ends.has(`${i}:${k}`)) keep |= 1 << k
      if (keep !== f) changes.set(i, { level: changes.get(i)?.level ?? levels.get(i) ?? 0, flags: keep })
    }
  }

  function changeLevels(axis, start, end, delta) {
    const w = ws().getSheet(), levels = levelMap(w, axis), changes = new Map()
    for (let i = start; i <= end; i++) {
      const L = levels.get(i) ?? 0, n = Math.max(0, Math.min(MAX_LEVEL, L + delta))
      if (n === L) continue
      if (n) levels.set(i, n); else levels.delete(i)
      changes.set(i, { level: n, flags: flagsOf(w, axis, i) })
    }
    if (!changes.size) return false
    cleanFlags(w, axis, levels, changes)
    return apply(w, axis, changes, [], [])
  }

  // ---- collapse / expand
  function setCollapsed(w, axis, g, collapse) {
    const levels = levelMap(w, axis), changes = new Map(), s = sumOf(w, axis, g)
    const flags = flagsOf(w, axis, s), bit = 1 << g.level
    const nf = collapse ? flags | bit : flags & ~bit
    if (nf !== flags && s >= 0) changes.set(s, { level: levels.get(s) ?? 0, flags: nf })
    const range = []
    for (let i = g.start; i <= g.end; i++) range.push(i)
    if (collapse) {
      const ok = apply(w, axis, changes, range, [])
      // Excel never leaves the active cell on a hidden row/column: it moves to the summary row/column.
      const g0 = sel(), cur = axis === 'row' ? g0.startRow : g0.startColumn
      if (cur >= g.start && cur <= g.end && s >= 0 && s < (axis === 'row' ? w.getRowCount() : w.getColumnCount())) {
        if (axis === 'row') ws().getRange(s, g0.startColumn).activate(); else ws().getRange(g0.startRow, s).activate()
      }
      return ok
    }
    // Expanding leaves nested groups that are themselves collapsed (bit set on their summary) hidden.
    const keepHidden = new Set()
    for (const n of outlineGroups(levels)) {
      if (n.level <= g.level || n.start < g.start || n.end > g.end) continue
      if (flagsOf(w, axis, sumOf(w, axis, n)) & (1 << n.level)) for (let i = n.start; i <= n.end; i++) keepHidden.add(i)
    }
    return apply(w, axis, changes, [], range.filter((i) => !keepHidden.has(i)))
  }
  function toggle(axis, g) { const w = ws().getSheet(); return setCollapsed(w, axis, g, !collapsed(w, axis, g)) }

  /** Level button k: groups of level >= k collapsed, those below k expanded (rows with level >= k hidden). */
  function showLevel(axis, k) {
    const w = ws().getSheet(), levels = levelMap(w, axis), changes = new Map(), hide = [], show = []
    for (const [i, L] of levels) (L >= k ? hide : show).push(i)
    const summaries = new Map()
    for (const g of outlineGroups(levels)) {
      const s = sumOf(w, axis, g), f = summaries.get(s) ?? flagsOf(w, axis, s)
      summaries.set(s, g.level >= k ? f | (1 << g.level) : f & ~(1 << g.level))
    }
    for (const [s, f] of summaries) if (f !== flagsOf(w, axis, s)) changes.set(s, { level: levels.get(s) ?? 0, flags: f })
    return apply(w, axis, changes, hide, show)
  }

  // ---- user commands
  const sel = () => ws().getActiveRange().getRange()
  const isFullRows = (g) => g.startColumn === 0 && g.endColumn >= ws().getMaxColumns() - 1
  const isFullCols = (g) => g.startRow === 0 && g.endRow >= ws().getMaxRows() - 1
  const span = (axis, g) => (axis === 'row' ? [g.startRow, g.endRow] : [g.startColumn, g.endColumn])
  function askAxis(title, run) {
    const r = radios('outline-axis', [['row', '&Rows'], ['col', '&Columns']], 'row')
    openDialog({ title, width: 200, body: el('div', {}, ...r.nodes), onOk: () => { run(r.value) } })
    return true
  }
  const message = (text) => { openDialog({ title: 'Sheetly', width: 320, body: el('p', { style: { margin: 0 } }, text), cancelLabel: 'Close' }); return true }
  function byAxis(title, fn) {
    if ((ws().getSelection?.()?.getActiveRangeList?.() ?? []).length > 1) return message('This command cannot be used on multiple selections.')
    const g = sel()
    if (isFullRows(g)) return fn('row', ...span('row', g)) || true
    if (isFullCols(g)) return fn('col', ...span('col', g)) || true
    return askAxis(title, (axis) => fn(axis, ...span(axis, g)))
  }
  const group = () => byAxis('Group', (axis, s, e) => {
    const w = ws().getSheet()
    for (let i = s; i <= e; i++) if (levelOf(w, axis, i) >= MAX_LEVEL) return message(`Outlines can have at most ${MAX_LEVEL + 1} levels.`)
    return changeLevels(axis, s, e, +1)
  })
  const ungroup = () => byAxis('Ungroup', (axis, s, e) => changeLevels(axis, s, e, -1))

  // Hide / Show Detail act on the innermost group holding the active cell (or summarised by it).
  function detail(collapse) {
    const g = sel(), w = ws().getSheet(), st = state(w)
    const find = (axis, list, i) => list.filter((x) => (x.start <= i && i <= x.end) || sumOf(w, axis, x) === i).sort((a, b) => b.level - a.level)
    const pick = (axis) => {
      const i = axis === 'row' ? g.startRow : g.startColumn
      const c = find(axis, axis === 'row' ? st.rows : st.cols, i)
      // Show Detail on a summary row expands the group it summarises; otherwise the innermost group.
      return c.find((x) => (collapse ? !collapsed(w, axis, x) : collapsed(w, axis, x))) ?? null
    }
    const order = isFullCols(g) ? ['col', 'row'] : ['row', 'col']
    for (const axis of order) { const x = pick(axis); if (x) return setCollapsed(w, axis, x, collapse) }
    return true
  }
  function clearOutline() {
    const w = ws().getSheet()
    for (const axis of ['row', 'col']) {
      const changes = new Map()
      for (const i of levelMap(w, axis).keys()) changes.set(i, { level: 0, flags: 0 })
      for (const i of Object.keys(AX[axis].data(w) ?? {}).map(Number)) if (flagsOf(w, axis, i)) changes.set(i, { level: 0, flags: 0 })
      if (changes.size) apply(w, axis, changes, [], [])
    }
    return true
  }
  function toggleSymbols() {
    const w = ws().getSheet()
    writeSettings(w, { ...settingsOf(w).value, hideSymbols: !settingsOf(w).hideSymbols })
    schedule()
    return true
  }
  /** Outline Settings: summary rows below / columns right. Collapse bits move to the new summary rows/columns. */
  function setDirection({ summaryBelow, summaryRight }) {
    const w = ws().getSheet(), old = settingsOf(w)
    const moves = [['row', old.summaryBelow, summaryBelow], ['col', old.summaryRight, summaryRight]]
    const bits = moves.map(([axis, from, to]) => {
      if (from === to) return null
      const list = []
      for (const g of outlineGroups(levelMap(w, axis))) {
        const a = summaryIndex(g, from), b = summaryIndex(g, to), bit = 1 << g.level
        if (a >= 0 && flagsOf(w, axis, a) & bit) list.push([a, b, bit])
      }
      return list
    })
    writeSettings(w, { ...old.value, summaryBelow, summaryRight })
    moves.forEach(([axis], k) => {
      if (!bits[k]?.length) return
      const levels = levelMap(w, axis), flags = new Map()
      const get = (i) => (flags.has(i) ? flags.get(i) : flagsOf(w, axis, i))
      for (const [a, , bit] of bits[k]) flags.set(a, get(a) & ~bit)
      for (const [, b, bit] of bits[k]) if (b >= 0) flags.set(b, get(b) | bit)
      const changes = new Map([...flags].map(([i, f]) => [i, { level: levels.get(i) ?? 0, flags: f }]))
      apply(w, axis, changes, [], [])
    })
    schedule()
    return true
  }

  // Auto Outline: a row whose formulas refer to cells of their own column ending right next to it (above when
  // summaries are below) summarises the rows from the first referenced one; a referenced row that is itself a
  // summary pulls in its detail, so a grand total over subtotals (=B5+B9 or SUM(B2:B9)) nests over them. A single
  // adjacent cell (=A3*2) is not a summary. Columns likewise with references along the row.
  function autoOutline({ confirm } = {}) {
    const w = ws().getSheet(), st = settingsOf(w)
    const found = { row: [], col: [] }
    const ref = /(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?/g
    const colNum = (L) => [...L].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
    const formulas = []
    // With a range selected only that range is outlined (Excel); with one cell, the whole sheet's data.
    const g = sel(), scoped = g.startRow !== g.endRow || g.startColumn !== g.endColumn
    const inScope = (r, c) => !scoped || (r >= g.startRow && r <= g.endRow && c >= g.startColumn && c <= g.endColumn)
    w.getCellMatrix().forValue((r, c, cell) => { if (typeof cell?.f === 'string' && cell.f.startsWith('=') && inScope(r, c)) formulas.push([r, c, cell.f]) })
    const refsOf = (f) => {
      // String literals and references into other sheets do not outline this one.
      const text = f.replace(/"(?:[^"]|"")*"/g, '""').replace(/'[^']*'![$A-Z0-9:]+|[A-Za-z_][\w.]*![$A-Z0-9:]+/g, '')
      return [...text.matchAll(ref)].map((m) => { const r0 = +m[4] - 1, c0 = colNum(m[2]); return m[6] ? { r0, c0, r1: +m[8] - 1, c1: colNum(m[6]), range: true } : { r0, c0, r1: r0, c1: c0, range: false } })
    }
    for (const [axis, forward] of [['row', st.summaryBelow], ['col', st.summaryRight]]) {
      const groups = new Map() // summary index -> [start, end]
      const list = formulas.map(([r, c, f]) => [axis === 'row' ? r : c, axis === 'row' ? c : r, f]).sort((x, y) => (forward ? x[0] - y[0] : y[0] - x[0]))
      for (const [i, j, f] of list) {
        // references lying along this summary's line (same column for rows, same row for columns)
        const along = refsOf(f).filter((x) => (axis === 'row' ? x.c0 === j && x.c1 === j : x.r0 === j && x.r1 === j))
          .map((x) => (axis === 'row' ? [x.r0, x.r1, x.range] : [x.c0, x.c1, x.range]))
        if (!along.length || (along.length === 1 && !along[0][2])) continue
        const lo = Math.min(...along.map((x) => x[0])), hi = Math.max(...along.map((x) => x[1]))
        if (forward ? hi !== i - 1 || lo >= i : lo !== i + 1 || hi <= i) continue
        let [a, b] = forward ? [lo, i - 1] : [i + 1, hi]
        for (const [k, [ga, gb]] of groups) if (k >= a && k <= b) { a = Math.min(a, ga); b = Math.max(b, gb) }
        const prev = groups.get(i)
        groups.set(i, prev ? [Math.min(prev[0], a), Math.max(prev[1], b)] : [a, b])
      }
      found[axis] = [...groups.values()]
    }
    if (!found.row.length && !found.col.length) return message('Cannot create an outline.')
    const run = () => {
      clearOutline()
      for (const axis of ['row', 'col']) {
        const levels = new Map()
        const spans = [...new Set(found[axis].map((x) => x.join(':')))].map((x) => x.split(':').map(Number))
        for (const [a, b] of spans) for (let i = a; i <= b; i++) levels.set(i, Math.min(MAX_LEVEL, (levels.get(i) ?? 0) + 1))
        const changes = new Map([...levels].map(([i, L]) => [i, { level: L, flags: 0 }]))
        if (changes.size) apply(w, axis, changes, [], [])
      }
      if (settingsOf(w).autoStyles) applyStyles()
      schedule()
      return true
    }
    const st0 = state(w)
    if ((st0.rowLevels || st0.colLevels) && confirm !== true) {
      openDialog({ title: 'Sheetly', width: 300, body: el('p', { style: { margin: 0 } }, 'Modify existing outline?'), okLabel: 'OK', onOk: run })
      return true
    }
    return run()
  }

  function settingsDialog() {
    const w = ws().getSheet(), st = settingsOf(w)
    const below = check('Summary rows &below detail', st.summaryBelow), right = check('Summary columns to &right of detail', st.summaryRight)
    const auto = check('Automatic st&yles', st.autoStyles)
    const dir = () => { setDirection({ summaryBelow: below.input.checked, summaryRight: right.input.checked }); writeSettings(w, { ...settingsOf(w).value, autoStyles: auto.input.checked }) }
    let d = null
    const create = button('&Create', () => { dir(); d.close(); autoOutline() })
    const styles = button('Apply &Styles', () => { dir(); d.close(); applyStyles() })
    d = openDialog({ title: 'Settings', width: 340, body: el('div', {}, fieldset('Direction', below.node, right.node), auto.node), extraButtons: [create, styles], onOk: () => { dir() } })
    return true
  }

  // Apply Styles: Excel's RowLevel_n / ColLevel_n cell styles on the summary rows / columns — level 1 bold, level 2
  // italic (levels 3-7 carry no formatting) — across the used part of the sheet.
  function applyStyles() {
    const w = ws().getSheet(), s = ws()
    const used = { rows: Math.max(0, w.getLastRowWithContent?.() ?? 0), cols: Math.max(0, w.getLastColumnWithContent?.() ?? 0) }
    for (const axis of ['row', 'col']) {
      for (const g of outlineGroups(levelMap(w, axis))) {
        if (g.level > 2) continue
        const i = sumOf(w, axis, g)
        if (i < 0) continue
        const r = axis === 'row' ? s.getRange(i, 0, 1, used.cols + 1) : s.getRange(0, i, used.rows + 1, 1)
        if (g.level === 1) r.setFontWeight('bold'); else r.setFontStyle('italic')
      }
    }
    return true
  }


  // ---- rendering: gutter inside a widened row / column header
  const gutter = (w, st) => {
    if (settingsOf(w).hideSymbols) return { row: 0, col: 0 }
    return { row: st.rowLevels ? (st.rowLevels + 1) * STEP + 2 : 0, col: st.colLevels ? (st.colLevels + 1) * STEP + 2 : 0 }
  }
  const box = (ctx, x, y, plus) => {
    const h = BOX / 2
    ctx.fillStyle = '#ffffff'; ctx.fillRect(x - h, y - h, BOX, BOX)
    ctx.strokeStyle = LINE; ctx.lineWidth = 1; ctx.strokeRect(x - h + 0.5, y - h + 0.5, BOX - 1, BOX - 1)
    ctx.beginPath(); ctx.moveTo(x - 2, y + 0.5); ctx.lineTo(x + 3, y + 0.5)
    if (plus) { ctx.moveTo(x + 0.5, y - 2); ctx.lineTo(x + 0.5, y + 3) }
    ctx.stroke()
  }

  class RowOutline extends SheetExtension {
    constructor() { super(); this.uKey = 'SheetlyRowOutline'; this.Z_INDEX = 30 }
    draw(ctx, _scale, sk) {
      const w = sk.worksheet, st = state(w), G = gutter(w, st).row
      if (!G) return
      const { rowHeightAccumulation: acc, rowHeaderWidth: W, rowColumnSegment: seg } = sk
      const top = (r) => (r > 0 ? acc[r - 1] : 0), bottom = (r) => acc[r]
      const y0 = top(Math.max(0, seg.startRow)), y1 = bottom(Math.min(acc.length - 1, seg.endRow))
      ctx.save()
      ctx.fillStyle = BG; ctx.fillRect(0, y0, G, y1 - y0)
      ctx.strokeStyle = '#d4d4d4'; ctx.beginPath(); ctx.moveTo(G - 0.5, y0); ctx.lineTo(G - 0.5, y1); ctx.stroke()
      // Row numbers, centred in the part of the header right of the gutter.
      ctx.fillStyle = TEXT; ctx.font = FONT; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'
      for (let r = Math.max(0, seg.startRow); r <= Math.min(acc.length - 1, seg.endRow); r++) {
        const h = bottom(r) - top(r)
        if (h >= 4) ctx.fillText(String(r + 1), G + (W - G) / 2, top(r) + h / 2 + 1)
      }
      for (const g of st.rows) {
        if (g.end < seg.startRow - 1 || g.start > seg.endRow + 1) continue
        const x = (g.level - 1) * STEP + STEP / 2 + 1
        const shut = collapsed(w, 'row', g), below = after(w, 'row')
        if (!shut && bottom(g.end) > top(g.start)) {
          // The bracket runs along the detail; its tick sits at the end away from the summary row.
          const a = below ? top(g.start) + 2 : top(g.start), b = below ? bottom(g.end) : bottom(g.end) - 2, tick = below ? a + 0.5 : b - 0.5
          ctx.strokeStyle = LINE; ctx.lineWidth = 1; ctx.beginPath()
          ctx.moveTo(x + 0.5, a); ctx.lineTo(x + 0.5, b); ctx.stroke()
          ctx.beginPath(); ctx.moveTo(x - 3, tick); ctx.lineTo(x + 1, tick); ctx.stroke()
        }
        const s = sumOf(w, 'row', g)
        if (s >= 0 && s < acc.length && bottom(s) - top(s) >= BOX) box(ctx, x + 0.5, top(s) + (bottom(s) - top(s)) / 2, shut)
      }
      ctx.restore()
    }
  }
  class ColOutline extends SheetExtension {
    constructor() { super(); this.uKey = 'SheetlyColumnOutline'; this.Z_INDEX = 30 }
    draw(ctx, _scale, sk) {
      const w = sk.worksheet, st = state(w), G = gutter(w, st).col
      if (!G) return
      const { columnWidthAccumulation: acc, columnHeaderHeight: H, rowColumnSegment: seg } = sk
      const left = (c) => (c > 0 ? acc[c - 1] : 0), right = (c) => acc[c]
      const x0 = left(Math.max(0, seg.startColumn)), x1 = right(Math.min(acc.length - 1, seg.endColumn))
      ctx.save()
      ctx.fillStyle = BG; ctx.fillRect(x0, 0, x1 - x0, G)
      ctx.strokeStyle = '#d4d4d4'; ctx.beginPath(); ctx.moveTo(x0, G - 0.5); ctx.lineTo(x1, G - 0.5); ctx.stroke()
      ctx.fillStyle = TEXT; ctx.font = FONT; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'
      for (let c = Math.max(0, seg.startColumn); c <= Math.min(acc.length - 1, seg.endColumn); c++) {
        const wd = right(c) - left(c)
        if (wd >= 4) ctx.fillText(columnLetters(c), left(c) + wd / 2, G + (H - G) / 2 + 1)
      }
      for (const g of st.cols) {
        if (g.end < seg.startColumn - 1 || g.start > seg.endColumn + 1) continue
        const y = (g.level - 1) * STEP + STEP / 2 + 1
        const shut = collapsed(w, 'col', g), toRight = after(w, 'col')
        if (!shut && right(g.end) > left(g.start)) {
          const a = toRight ? left(g.start) + 2 : left(g.start), b = toRight ? right(g.end) : right(g.end) - 2, tick = toRight ? a + 0.5 : b - 0.5
          ctx.strokeStyle = LINE; ctx.lineWidth = 1; ctx.beginPath()
          ctx.moveTo(a, y + 0.5); ctx.lineTo(b, y + 0.5); ctx.stroke()
          ctx.beginPath(); ctx.moveTo(tick, y - 3); ctx.lineTo(tick, y + 1); ctx.stroke()
        }
        const s = sumOf(w, 'col', g)
        if (s >= 0 && s < acc.length && right(s) - left(s) >= BOX) box(ctx, left(s) + (right(s) - left(s)) / 2, y + 0.5, shut)
      }
      ctx.restore()
    }
  }

  // Per workbook render unit: extensions, header sizes, corner level buttons, gutter clicks.
  const installed = new Map() // unitId -> { render, corner, base }
  // Univer's services and render controllers are created after this module and after the first render, so
  // installation is retried on each refresh until they exist (the lookups that can fail run first).
  function install(unitId) {
    if (installed.has(unitId)) return installed.get(unitId)
    const render = renders.getRenderUnitById(unitId)
    const rowHeader = render?.components.get('__SpreadsheetRowHeader__'), colHeader = render?.components.get('__SpreadsheetColumnHeader__')
    if (!rowHeader || !colHeader) return null
    let unhide, sk
    try { unhide = render.with(HeaderUnhideRenderController); sk = render.with(SheetSkeletonManagerService); hookInsert() } catch { return null }
    if (!sk.getCurrentSkeleton()) return null
    rowHeader.register(new RowOutline())
    colHeader.register(new ColOutline())
    const corner = el('div', { class: 'sx-outline-corner', style: { position: 'fixed', zIndex: 10, display: 'none', pointerEvents: 'none' } })
    document.body.append(corner)
    const canvas = render.engine.getCanvasElement()
    canvas.addEventListener('pointerdown', (e) => onGutterPointer(unitId, e), true)
    // Univer draws its own "unhide" arrows next to hidden rows/columns; collapsed outline groups use +/- instead.
    unhide.interceptor.intercept(HEADER_UNHIDE_RANGE_VISIBLE_CHECK, {
      handler: (value, payload, next) => {
        const { axis, range, worksheet } = payload, a = axis === 'row' ? 'row' : 'col'
        const [s, e] = a === 'row' ? [range.startRow, range.endRow] : [range.startColumn, range.endColumn]
        let grouped = true
        for (let i = s; i <= e && grouped; i++) grouped = levelOf(worksheet, a, i) > 0
        return grouped ? false : next(value)
      } })
    const entry = { render, corner, canvas }
    installed.set(unitId, entry)
    return entry
  }

  function onGutterPointer(unitId, e) {
    const entry = installed.get(unitId)
    if (!entry || e.button !== 0) return
    const w = ws().getSheet(), st = state(w), G = gutter(w, st)
    if (!G.row && !G.col) return
    const { scene } = entry.render
    const sk = entry.render.with(SheetSkeletonManagerService).getCurrentSkeleton()
    const { scaleX, scaleY } = scene.getAncestorScale()
    const x = e.offsetX / scaleX, y = e.offsetY / scaleY
    const inRowGutter = x < G.row && y > sk.columnHeaderHeight, inColGutter = y < G.col && x > sk.rowHeaderWidth
    if (!inRowGutter && !inColGutter) return
    e.preventDefault(); e.stopImmediatePropagation()
    const rel = scene.getCoordRelativeToViewport(Vector2.FromArray([e.offsetX, e.offsetY]))
    const scroll = scene.getScrollXYInfoByViewport(rel)
    const axis = inRowGutter ? 'row' : 'col'
    const i = axis === 'row' ? sk.getRowIndexByOffsetY(rel.y, scaleY, scroll) : sk.getColumnIndexByOffsetX(rel.x, scaleX, scroll)
    const level = Math.floor((axis === 'row' ? x : y) / STEP) + 1
    const g = (axis === 'row' ? st.rows : st.cols).find((x2) => x2.level === level && (sumOf(w, axis, x2) === i || (x2.start <= i && i <= x2.end)))
    if (g) commands.executeCommand('sheetly.outline.toggle', { axis, start: g.start, end: g.end, level: g.level })
  }

  // Corner: level buttons (1..n+1) for rows along the bottom, for columns down the right.
  function drawCorner(entry, w, st, G) {
    const { corner, canvas } = entry
    if (!G.row && !G.col) { corner.style.display = 'none'; return }
    const r = canvas.getBoundingClientRect(), { scaleX, scaleY } = entry.render.scene.getAncestorScale()
    const sk = entry.render.with(SheetSkeletonManagerService).getCurrentSkeleton()
    Object.assign(corner.style, { display: 'block', left: r.left + 'px', top: r.top + 'px', width: sk.rowHeaderWidth * scaleX + 'px', height: sk.columnHeaderHeight * scaleY + 'px' })
    corner.replaceChildren()
    const btn = (axis, k, left, top) => el('button', { type: 'button', class: 'sx-outline-level', title: `Show ${axis === 'row' ? 'row' : 'column'} level ${k}`,
      'data-axis': axis, 'data-level': k,
      style: { position: 'absolute', left: left * scaleX + 'px', top: top * scaleY + 'px', width: (STEP - 3) * scaleX + 'px', height: (STEP - 3) * scaleY + 'px',
        font: `${Math.round(10 * scaleY)}px system-ui, sans-serif`, padding: 0, lineHeight: 1, border: `1px solid ${LINE}`, background: '#fff', color: TEXT, pointerEvents: 'auto', cursor: 'pointer' },
      onPointerdown: (e) => { e.preventDefault(); e.stopPropagation() },
      onClick: () => commands.executeCommand('sheetly.outline.show-level', { axis, level: k }) }, String(k))
    const midY = G.col + (sk.columnHeaderHeight - G.col - STEP + 3) / 2
    for (let k = 1; k <= st.rowLevels + 1 && G.row; k++) corner.append(btn('row', k, (k - 1) * STEP + 2, midY))
    const midX = G.row + (sk.rowHeaderWidth - G.row - STEP + 3) / 2
    for (let k = 1; k <= st.colLevels + 1 && G.col; k++) corner.append(btn('col', k, midX, (k - 1) * STEP + 2))
  }

  // Bring header sizes, default header text and the corner in line with the active sheet's outline. The gutter
  // is added to the worksheet's own header config (what Univer lays the viewports out from) on top of Univer's
  // default sizes, so a saved, widened size never compounds.
  function refresh() {
    const book = wb()
    if (!book) return
    const unitId = book.getId(), entry = install(unitId)
    if (!entry) return
    const w = ws().getSheet(), id = w.getSheetId(), st = state(w), G = gutter(w, st)
    const cfg = w.getConfig()
    const b = { width: BASE_ROW_HEADER, height: BASE_COL_HEADER }
    // Univer widens the row header to fit 7-digit row numbers; keep that room for the numbers next to the gutter.
    const numbers = Math.max(b.width, String(w.getRowCount()).length * 8)
    const width = G.row ? numbers + G.row : b.width, height = b.height + G.col
    const rowHeader = entry.render.components.get('__SpreadsheetRowHeader__'), colHeader = entry.render.components.get('__SpreadsheetColumnHeader__')
    // While a gutter is shown the outline extension draws the numbers / letters itself.
    rowHeader.rowHeaderLayoutExtension.configHeaderRow({ headerStyle: G.row ? { fontColor: TRANSPARENT } : {} }, id)
    colHeader.columnHeaderLayoutExtension.configHeaderColumn({ headerStyle: G.col ? { fontColor: TRANSPARENT } : {} }, id)
    if (cfg.rowHeader.width !== width || cfg.columnHeader.height !== height) {
      cfg.rowHeader.width = width
      cfg.columnHeader.height = height
      const sk = entry.render.with(SheetSkeletonManagerService).getCurrentSkeleton()
      sk?.makeDirty(true); sk?.calculate()
      // Re-applying the current freeze makes Univer lay the viewports out again from the new header size.
      const { startRow = -1, startColumn = -1, ySplit = 0, xSplit = 0 } = cfg.freeze ?? {}
      commands.syncExecuteCommand('sheet.mutation.set-frozen', { unitId, subUnitId: id, startRow, startColumn, ySplit, xSplit, resetScroll: false })
    }
    rowHeader.makeDirty(true); colHeader.makeDirty(true)
    entry.render.scene.makeDirty?.(true)
    drawCorner(entry, w, st, G)
  }

  const commandDefs = {
    'sheetly.outline.group': () => group(),
    'sheetly.outline.ungroup': () => ungroup(),
    'sheetly.outline.toggle': (p) => toggle(p.axis, p),
    'sheetly.outline.show-level': (p) => showLevel(p.axis, p.level),
    'sheetly.outline.hide-detail': () => detail(true),
    'sheetly.outline.show-detail': () => detail(false),
    'sheetly.outline.clear': () => clearOutline(),
    'sheetly.outline.toggle-symbols': () => toggleSymbols(),
    'sheetly.outline.auto': (p) => autoOutline(p),
    'sheetly.outline.settings': () => settingsDialog(),
    'sheetly.outline.set-direction': (p) => setDirection(p),
    'sheetly.outline.apply-styles': () => applyStyles(),
  }
  for (const [id, fn] of Object.entries(commandDefs)) commands.registerCommand({ id, type: CommandType.COMMAND, handler: (_a, p) => fn(p ?? {}) !== false })

  // A row/column inserted inside a group joins it (Excel): it gets the lower level of its two neighbours.
  // Registered on first use: the service is not available yet while the app is being created.
  let insertHooked = false
  const hookInsert = () => { if (insertHooked) return
  injector.get(SheetInterceptorService).interceptCommand({ getMutations: (info) => {
    const axis = info.id === 'sheet.command.insert-row' ? 'row' : info.id === 'sheet.command.insert-col' ? 'col' : null
    const p = info.params
    if (!axis || !p?.range) return { redos: [], undos: [] }
    const w = worksheetOf(p.unitId, p.subUnitId)
    if (!w) return { redos: [], undos: [] }
    const [s, e] = axis === 'row' ? [p.range.startRow, p.range.endRow] : [p.range.startColumn, p.range.endColumn]
    const level = Math.min(levelOf(w, axis, s - 1), levelOf(w, axis, s))
    if (!level) return { redos: [], undos: [] }
    const data = {}
    for (let i = s; i <= e; i++) data[i] = { custom: { outlineLevel: level } }
    // The inserted rows/columns disappear on undo, so nothing to undo here.
    return { redos: [{ id: AX[axis].mutation, params: { unitId: p.unitId, subUnitId: p.subUnitId, [AX[axis].key]: data } }], undos: [] }
  } })
  insertHooked = true
  }

  // Re-sync after anything that can change levels, hidden state, sizes or the visible sheet (incl. undo/redo).
  let pending = false
  const schedule = () => { if (pending) return; pending = true; requestAnimationFrame(() => { pending = false; refresh() }) }
  commands.onCommandExecuted((cmd) => {
    if (cmd.type === CommandType.MUTATION || cmd.id.startsWith('sheet.operation.set-worksheet-active') || cmd.id === 'sheet.command.set-zoom-ratio' || cmd.id.includes('zoom')) schedule()
  })
  window.addEventListener('resize', schedule)
  setTimeout(refresh, 0)

  return {
    group, ungroup, toggle, showLevel, hideDetail: () => detail(true), showDetail: () => detail(false), clearOutline, toggleSymbols, refresh,
    autoOutline, settingsDialog, setDirection, applyStyles,
    settings: () => settingsOf(ws().getSheet()).value,
    showSymbols: () => { const w = ws().getSheet(); if (settingsOf(w).hideSymbols) writeSettings(w, { ...settingsOf(w).value, hideSymbols: false }); schedule() },
    /** For tests: outline of the active sheet. */
    info: () => { const w = ws().getSheet(), st = state(w); return { ...st, gutter: gutter(w, st),
      rowsCollapsed: st.rows.map((g) => collapsed(w, 'row', g)), colsCollapsed: st.cols.map((g) => collapsed(w, 'col', g)),
      rowHeaderWidth: installed.get(wb().getId())?.render.with(SheetSkeletonManagerService).getCurrentSkeleton()?.rowHeaderWidth } },
    levelOf: (axis, i) => levelOf(ws().getSheet(), axis, i),
  }
}
