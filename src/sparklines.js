// Sparklines: Insert > Sparklines > Line / Column / Win/Loss (Alt N S L / O / W), the Create Sparklines dialog
// (Data Range and Location Range), drawn inside their cells on the grid and redrawn as the data change. Each
// location cell plots one row (or column) of the data range, as in Excel. Groups are kept in the sheet's custom
// data (custom.sheetlySparklines) and written to .xlsx as x14:sparklineGroups (see xlsx-zip.js).
import { CommandType, ICommandService, IUniverInstanceService } from '@univerjs/core'
import { IRenderManagerService, SheetExtension } from '@univerjs/engine-render'
import { el, field, openDialog } from './ui/dialog.js'
import { parseRange, rangeA1 } from './xlsx-objects.js'

// Excel's default sparkline colours (Style 1 of the 2010 set): dark blue series, red negative points and markers.
export const SPARK_DEFAULTS = { color: '#376092', negative: '#D00000', markers: '#D00000' }
const PAD = 2
// The next frame, or a short timer when the window is in the background (Chromium holds animation frames back there).
const soon = (fn) => { let done = false; const run = () => { if (!done) { done = true; fn() } }; requestAnimationFrame(run); setTimeout(run, 60) }


/** Values of one sparkline: numbers in order; text and blanks are gaps (null). */
export const sparkValues = (get, g) => {
  const out = []
  for (let r = g.startRow; r <= g.endRow; r++) for (let c = g.startColumn; c <= g.endColumn; c++) { const v = get(r, c); out.push(typeof v === 'number' ? v : null) }
  return out
}

/** Pairs each location cell with its slice of the data range; null when the shapes do not match (Excel refuses). */
export function pairUp(data, loc) {
  const cells = []
  if (loc.startRow === loc.endRow) for (let c = loc.startColumn; c <= loc.endColumn; c++) cells.push({ row: loc.startRow, col: c })
  else if (loc.startColumn === loc.endColumn) for (let r = loc.startRow; r <= loc.endRow; r++) cells.push({ row: r, col: loc.startColumn })
  else return { error: 'Location reference is not valid because the cells are not all in the same column or same row. Select cells that are all in a single column or row.' }
  const rows = data.endRow - data.startRow + 1, cols = data.endColumn - data.startColumn + 1
  if (cells.length === 1) return { items: [{ ...cells[0], source: data }] }
  // n location cells: n data rows (one sparkline per row), else n data columns.
  if (rows === cells.length) return { items: cells.map((x, i) => ({ ...x, source: { startRow: data.startRow + i, endRow: data.startRow + i, startColumn: data.startColumn, endColumn: data.endColumn } })) }
  if (cols === cells.length) return { items: cells.map((x, i) => ({ ...x, source: { startRow: data.startRow, endRow: data.endRow, startColumn: data.startColumn + i, endColumn: data.startColumn + i } })) }
  return { error: 'Data Range reference is not valid: it does not match the number of cells in the location range.' }
}

/** Draws one sparkline into a box. */
export function drawSparkline(ctx, type, values, box, colors = SPARK_DEFAULTS) {
  const nums = values.filter((v) => v != null)
  if (!nums.length) return
  const x0 = box.x + PAD, y0 = box.y + PAD, w = Math.max(1, box.w - 2 * PAD), h = Math.max(1, box.h - 2 * PAD)
  const n = values.length
  ctx.save()
  if (type === 'column' || type === 'winloss') {
    const lo = Math.min(0, ...nums), hi = Math.max(0, ...nums)
    const slot = w / n, bw = Math.max(1, slot * 0.7)
    const y = (v) => (type === 'winloss' ? (v > 0 ? y0 : v < 0 ? y0 + h / 2 : null) : y0 + ((hi - v) / (hi - lo || 1)) * h)
    const zero = type === 'winloss' ? y0 + h / 2 : y0 + ((hi - 0) / (hi - lo || 1)) * h
    values.forEach((v, i) => {
      if (v == null || (type === 'winloss' && v === 0)) return
      ctx.fillStyle = v < 0 ? colors.negative : colors.color
      const top = type === 'winloss' ? y(v) : Math.min(y(v), zero), height = type === 'winloss' ? h / 2 - 1 : Math.max(1, Math.abs(zero - y(v)))
      ctx.fillRect(x0 + i * slot + (slot - bw) / 2, top, bw, height)
    })
  } else {
    const lo = Math.min(...nums), hi = Math.max(...nums)
    const px = (i) => x0 + (n === 1 ? w / 2 : (i / (n - 1)) * w), py = (v) => y0 + (hi === lo ? h / 2 : ((hi - v) / (hi - lo)) * h)
    ctx.strokeStyle = colors.color; ctx.lineWidth = 1.5; ctx.lineJoin = 'round'
    ctx.beginPath()
    let pen = false
    values.forEach((v, i) => { if (v == null) { pen = false; return } if (pen) ctx.lineTo(px(i), py(v)); else { ctx.moveTo(px(i), py(v)); pen = true } })
    ctx.stroke()
  }
  ctx.restore()
}

export function createSparklines(api) {
  const injector = api._injector
  const commands = injector.get(ICommandService)
  const wb = () => api.getActiveWorkbook()
  const ws = () => wb().getActiveSheet()
  const groupsOf = (w) => w.getCustomMetadata?.()?.sheetlySparklines ?? []
  const store = (w, list) => {
    const custom = { ...(w.getCustomMetadata?.() ?? {}) }
    if (list.length) custom.sheetlySparklines = list; else delete custom.sheetlySparklines
    w.setCustomMetadata(Object.keys(custom).length ? custom : undefined)
  }

  class SparklineExt extends SheetExtension {
    constructor() { super(); this.uKey = 'SheetlySparklines'; this.Z_INDEX = 45 }
    draw(ctx, _scale, sk) {
      const w = sk.worksheet
      const groups = groupsOf(w)
      if (!groups.length) return
      const book = injector.get(IUniverInstanceService).getUnit(w.getUnitId?.() ?? wb().getId())
      const seg = sk.rowColumnSegment
      for (const g of groups) for (const it of g.items) {
        if (it.row < seg.startRow - 1 || it.row > seg.endRow + 1 || it.col < seg.startColumn - 1 || it.col > seg.endColumn + 1) continue
        if (!w.getRowVisible(it.row) || !w.getColVisible(it.col)) continue
        const srcSheet = (book?.getSheetBySheetId?.(it.sheetId ?? w.getSheetId())) ?? w
        const values = sparkValues((r, c) => srcSheet.getCellRaw(r, c)?.v ?? null, it.source)
        const cell = sk.getCellWithCoordByIndex(it.row, it.col, false)
        drawSparkline(ctx, g.type, values, { x: cell.startX, y: cell.startY, w: cell.endX - cell.startX, h: cell.endY - cell.startY }, { ...SPARK_DEFAULTS, ...g.colors })
      }
    }
  }
  const installed = new Set()
  function install() {
    const book = wb()
    if (!book || installed.has(book.getId())) return
    const render = injector.get(IRenderManagerService).getRenderUnitById(book.getId())
    const main = render?.components.get('__SpreadsheetRender__')
    if (!main) return
    main.register(new SparklineExt())
    installed.add(book.getId())
  }
  const redraw = () => {
    install()
    const render = injector.get(IRenderManagerService).getRenderUnitById(wb()?.getId())
    render?.components.get('__SpreadsheetRender__')?.makeDirty(true)
  }
  let pending = false
  commands.onCommandExecuted((cmd) => {
    if (cmd.type !== CommandType.MUTATION || pending) return
    pending = true
    soon(() => { pending = false; redraw() })
  })
  setTimeout(install, 0)

  // Inserted / deleted rows and columns move sparklines and their data, as cell references move in Excel; a
  // sparkline whose cell is deleted goes with it. Undo replays the opposite mutation, which moves them back.
  const SHIFT = { 'sheet.mutation.insert-row': ['row', 1], 'sheet.mutation.remove-rows': ['row', -1], 'sheet.mutation.insert-col': ['col', 1], 'sheet.mutation.remove-col': ['col', -1] }
  const moveIndex = (i, sign, s, e) => (sign > 0 ? (i >= s ? i + (e - s + 1) : i) : i < s ? i : i > e ? i - (e - s + 1) : null)
  const moveRange = (g, axis, sign, s, e) => {
    const [a, b] = axis === 'row' ? ['startRow', 'endRow'] : ['startColumn', 'endColumn']
    const n = e - s + 1, out = { ...g }
    if (sign > 0) { if (s <= g[a]) { out[a] += n; out[b] += n } else if (s <= g[b]) out[b] += n; return out }
    const gone = (x) => Math.max(0, Math.min(e, x) - s + 1)
    out[a] = g[a] - gone(g[a] - 1); out[b] = g[b] - gone(g[b])
    return out.endRow >= out.startRow && out.endColumn >= out.startColumn ? out : null
  }
  commands.onCommandExecuted((cmd) => {
    const how = SHIFT[cmd.id], p = cmd.params
    if (!how || !p?.range) return
    const [axis, sign] = how
    const [s, e] = axis === 'row' ? [p.range.startRow, p.range.endRow] : [p.range.startColumn, p.range.endColumn]
    const book = injector.get(IUniverInstanceService).getUnit(p.unitId)
    for (const w of book?.getSheets?.() ?? []) {
      const groups = groupsOf(w)
      if (!groups.length) continue
      const here = w.getSheetId() === p.subUnitId
      const next = groups.map((g) => ({ ...g, items: g.items.map((it) => {
        let { row, col, source } = it
        if (here) { if (axis === 'row') row = moveIndex(row, sign, s, e); else col = moveIndex(col, sign, s, e) }
        if (row == null || col == null) return null
        if ((it.sheetId ?? w.getSheetId()) === p.subUnitId) source = moveRange(source, axis, sign, s, e)
        return source ? { ...it, row, col, source } : null
      }).filter(Boolean) })).filter((g) => g.items.length)
      if (JSON.stringify(next) !== JSON.stringify(groups)) store(w, next)
    }
  })

  function add(type, dataText, locText) {
    const ref = (t) => { const m = /^\s*(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(t); const sheet = m ? wb().getSheets().find((s) => s.getSheetName() === (m[1] ?? m[2]).replace(/''/g, "'")) : ws(); const g = parseRange(m ? m[3] : t); return sheet && g ? { sheet, g } : null }
    const data = ref(dataText), loc = ref(locText)
    if (!data) return { error: 'Data Range reference is not valid.' }
    if (!loc) return { error: 'Location reference is not valid.' }
    if (loc.sheet.getSheetId() !== ws().getSheetId()) return { error: 'Location reference is not valid.' }
    const p = pairUp(data.g, loc.g)
    if (p.error) return p
    const w = ws().getSheet()
    const taken = new Set(p.items.map((x) => `${x.row}:${x.col}`))
    // A cell holds one sparkline: the new group replaces older ones in the same cells.
    const groups = groupsOf(w).map((g) => ({ ...g, items: g.items.filter((x) => !taken.has(`${x.row}:${x.col}`)) })).filter((g) => g.items.length)
    groups.push({ id: `spk${Date.now().toString(36)}`, type, colors: {}, items: p.items.map((x) => ({ ...x, sheetId: data.sheet.getSheetId() })) })
    store(w, groups)
    redraw()
    return { ok: true }
  }
  function dialog(type) {
    const g = ws().getActiveRange().getRange()
    const dataIn = el('input', { type: 'text', value: rangeA1(g), style: { width: '220px' } })
    const locIn = el('input', { type: 'text', value: '', style: { width: '220px' } })
    openDialog({ title: 'Create Sparklines', width: 360, initialFocus: locIn,
      body: el('div', {}, el('div', { style: { fontWeight: 600 } }, 'Choose the data that you want'), field('Data Range:', dataIn),
        el('div', { style: { fontWeight: 600, marginTop: '8px' } }, 'Choose where you want the sparklines to be placed'), field('Location Range:', locIn)),
      onOk: () => {
        const r = add(type, dataIn.value, locIn.value.replace(/\$/g, ''))
        if (r.error) { openDialog({ title: 'Sheetly', width: 380, body: el('p', { style: { margin: 0 } }, r.error), cancelLabel: 'Close' }); return false }
        return true
      } })
    return true
  }
  /** Clear Selected Sparklines (Sparkline tab > Clear). */
  function clear() {
    const g = ws().getActiveRange().getRange(), w = ws().getSheet()
    const inSel = (x) => x.row >= g.startRow && x.row <= g.endRow && x.col >= g.startColumn && x.col <= g.endColumn
    store(w, groupsOf(w).map((gr) => ({ ...gr, items: gr.items.filter((x) => !inSel(x)) })).filter((gr) => gr.items.length))
    redraw()
    return true
  }
  return { dialog, add, clear, groups: () => groupsOf(ws().getSheet()), redraw }
}
