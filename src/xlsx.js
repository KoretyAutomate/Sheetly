// .xlsx <-> Univer snapshot via ExcelJS: values (dates as Excel serials), formulas (shared ones translated, the
// _xlfn. prefixes of newer functions removed on load and added on save) and their cached results, cell / row /
// column styles (font, fill, borders, alignment, number format; theme and indexed colours resolved), column
// widths, row heights (custom ones only), merges, frozen panes, tab colours, hidden sheets, defined names (incl.
// sheet-scoped names and constants / formulas), notes, and the outline: row/column levels, hidden flags, Excel's
// "collapsed" flag on summary rows/columns, summary direction (outlinePr) and the Ctrl+8 outline-symbols setting
// (sheetView showOutlineSymbols). ExcelJS is patched for collapsed, showOutlineSymbols, customHeight and raw
// defined names (patches/exceljs+4.4.0.patch).
// On load a summary flagged collapsed, or a group whose detail is all hidden, sets custom.outlineCollapsed.
import ExcelJS from 'exceljs'
import { outlineGroups, outlineSettings } from './outline-model.js'

const NOTE = 'SHEET_NOTE_PLUGIN', NAMES = 'SHEET_DEFINED_NAME_PLUGIN', ALL_SHEETS = 'AllDefaultWorkbook'
const EPOCH = Date.UTC(1899, 11, 30)
const PX_PER_PT = 4 / 3
const NOTE_SIZE = { width: 160, height: 72 }

// ---- colours
/** '#rgb', '#rrggbb', '#rrggbbaa', 'rgb(r,g,b)', 'rgba(r,g,b,a)' -> '#rrggbb' (alpha dropped). */
export function normColor(c) {
  if (typeof c !== 'string') return undefined
  const s = c.trim().toLowerCase()
  let m = s.match(/^#([0-9a-f]{3})$/)
  if (m) return '#' + [...m[1]].map((x) => x + x).join('')
  m = s.match(/^#([0-9a-f]{6})([0-9a-f]{2})?$/)
  if (m) return '#' + m[1]
  m = s.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/)
  if (m) return '#' + m.slice(1, 4).map((x) => Math.min(255, +x).toString(16).padStart(2, '0')).join('')
  return undefined
}
const argb = (c) => { const n = normColor(c); return n ? 'FF' + n.slice(1).toUpperCase() : undefined }

// Excel's legacy indexed palette (0..63); 64 = system foreground, 65 = system background.
const INDEXED = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333 000000 FFFFFF').split(' ')
// Office theme (2013-2022) used when the file has no theme part: dk1 lt1 dk2 lt2 accent1..6 hlink folHlink.
const OFFICE_THEME = ['000000', 'FFFFFF', '44546A', 'E7E6E6', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']
const SCHEME = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

function themePalette(wb) {
  const xml = wb._themes?.theme1
  if (typeof xml !== 'string') return OFFICE_THEME
  const out = SCHEME.map((name, i) => {
    const m = xml.match(new RegExp(`<a:${name}>([\\s\\S]*?)</a:${name}>`))
    const v = m?.[1].match(/srgbClr val="([0-9A-Fa-f]{6})"/)?.[1] ?? m?.[1].match(/lastClr="([0-9A-Fa-f]{6})"/)?.[1]
    return (v ?? OFFICE_THEME[i]).toUpperCase()
  })
  return out
}
// Theme colour index -> scheme slot: 0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2 (swapped), then accents, links.
const THEME_SLOT = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11]
function applyTint(hex, tint) {
  if (!tint) return hex
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
  const max = Math.max(r, g, b), min = Math.min(r, g, b)
  let h = 0, s = 0, l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
    h /= 6
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  const f = (p, q, t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p }
  if (s === 0) r = g = b = l
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3) }
  return [r, g, b].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase()
}
/** ExcelJS colour ({argb} | {theme, tint} | {indexed}) -> '#rrggbb'. */
function colorIn(c, theme) {
  if (!c) return undefined
  if (typeof c.argb === 'string' && /^[0-9A-F]{8}$/i.test(c.argb)) return '#' + c.argb.slice(2).toLowerCase()
  if (typeof c.theme === 'number') { const hex = theme[THEME_SLOT[c.theme] ?? c.theme]; return hex ? '#' + applyTint(hex, c.tint).toLowerCase() : undefined }
  if (typeof c.indexed === 'number' && INDEXED[c.indexed]) return '#' + INDEXED[c.indexed].toLowerCase()
  return undefined
}

// ---- styles
// Univer BorderStyleTypes <-> SpreadsheetML border styles.
const BORDERS = ['', 'thin', 'hair', 'dotted', 'dashed', 'dashDot', 'dashDotDot', 'double', 'medium', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot', 'thick']
const SIDES = { t: 'top', b: 'bottom', l: 'left', r: 'right' }
// Univer HorizontalAlign 1..6, VerticalAlign 1..3, WrapStrategy 3 = wrap.
const H_ALIGN = { 1: 'left', 2: 'center', 3: 'right', 4: 'justify', 5: 'justify', 6: 'distributed' }
const H_ALIGN_IN = { left: 1, center: 2, right: 3, justify: 4, distributed: 6, centerContinuous: 2, fill: 1 }
const V_ALIGN = { 1: 'top', 2: 'middle', 3: 'bottom' }
const V_ALIGN_IN = { top: 1, middle: 2, center: 2, bottom: 3, justify: 2, distributed: 2 }
const WRAP = 3
// ExcelJS reports built-in format 14 (Excel's short date, m/d/yyyy in en-US) as 'mm-dd-yy'.
const NUMFMT_IN = { 'mm-dd-yy': 'm/d/yyyy' }

/** ExcelJS style (cell, row or column) -> Univer IStyleData (only what is set). */
function styleIn(x, theme) {
  const s = {}, f = x.font, a = x.alignment, fill = x.fill, b = x.border
  if (f) {
    if (f.bold) s.bl = 1
    if (f.italic) s.it = 1
    if (f.underline && f.underline !== 'none') s.ul = { s: 1 }
    if (f.strike) s.st = { s: 1 }
    if (f.size && f.size !== 11) s.fs = f.size
    if (f.name && f.name !== 'Calibri') s.ff = f.name
    const c = colorIn(f.color, theme)
    if (c && c !== '#000000') s.cl = { rgb: c }
  }
  if (fill?.type === 'pattern' && fill.pattern === 'solid') { const c = colorIn(fill.fgColor, theme); if (c) s.bg = { rgb: c } }
  if (b) {
    const bd = {}
    for (const [k, side] of Object.entries(SIDES)) {
      const y = b[side], i = BORDERS.indexOf(y?.style)
      if (i > 0) bd[k] = { s: i, cl: { rgb: colorIn(y.color, theme) ?? '#000000' } }
    }
    if (Object.keys(bd).length) s.bd = bd
  }
  if (a) {
    if (H_ALIGN_IN[a.horizontal]) s.ht = H_ALIGN_IN[a.horizontal]
    if (V_ALIGN_IN[a.vertical]) s.vt = V_ALIGN_IN[a.vertical]
    if (a.wrapText) s.tb = WRAP
    // Excel counts degrees counterclockwise, Univer clockwise ("angle up" is -45 there).
    if (a.textRotation === 'vertical') s.tr = { a: 0, v: 1 }
    else if (a.textRotation) s.tr = { a: -a.textRotation, v: 0 }
  }
  if (x.numFmt && x.numFmt !== 'General') s.n = { pattern: NUMFMT_IN[x.numFmt] ?? x.numFmt }
  return s
}

/** Univer IStyleData -> ExcelJS style object. */
function styleOut(st) {
  const out = {}, font = {}
  if (st.bl) font.bold = true
  if (st.it) font.italic = true
  if (st.ul?.s) font.underline = true
  if (st.st?.s) font.strike = true
  if (st.fs) font.size = st.fs
  if (st.ff) font.name = st.ff
  if (argb(st.cl?.rgb)) font.color = { argb: argb(st.cl.rgb) }
  if (Object.keys(font).length) out.font = { name: 'Calibri', size: 11, ...font }
  if (argb(st.bg?.rgb)) out.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(st.bg.rgb) } }
  if (st.bd) {
    const border = {}
    for (const [k, side] of Object.entries(SIDES)) {
      const d = st.bd[k]
      if (d?.s > 0 && BORDERS[d.s]) border[side] = { style: BORDERS[d.s], color: { argb: argb(d.cl?.rgb) ?? 'FF000000' } }
    }
    if (Object.keys(border).length) out.border = border
  }
  const al = {}
  if (H_ALIGN[st.ht]) al.horizontal = H_ALIGN[st.ht]
  if (V_ALIGN[st.vt]) al.vertical = V_ALIGN[st.vt]
  if (st.tb === WRAP) al.wrapText = true
  if (st.tr?.v) al.textRotation = 'vertical'
  else if (st.tr?.a) al.textRotation = Math.max(-90, Math.min(90, -st.tr.a))
  if (Object.keys(al).length) out.alignment = al
  if (st.n?.pattern) out.numFmt = st.n.pattern
  return out
}
const isEmptyStyle = (s) => !s || !Object.keys(s).length
// Univer resolves a cell's look as column style < row style < cell style; Excel stores the result on the cell.
const mergeStyles = (...list) => {
  const out = {}
  for (const s of list) if (s) for (const [k, v] of Object.entries(s)) out[k] = k === 'bd' && out.bd ? { ...out.bd, ...v } : v
  return out
}

// ---- values and formulas
export const ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#GETTING_DATA',
  '#CONNECT!', '#BLOCKED!', '#UNKNOWN!', '#FIELD!', '#BUSY!'])
// Functions Excel stores with the _xlfn. prefix (added after Excel 2007); FILTER and SORT also carry _xlws.
const XLFN = new Set(('ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT ' +
  'BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS ' +
  'CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DETECTLANGUAGE DROP ' +
  'ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FIELDVALUE FILTERXML FLOOR.MATH ' +
  'FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA ' +
  'GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS GROUPBY HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH ' +
  'IMTAN ISFORMULA ISO.CEILING ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL ' +
  'MUNIT NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC ' +
  'PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI PIVOTBY POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ ' +
  'REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P SORTBY STDEV.P STDEV.S STOCKHISTORY SWITCH ' +
  'T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW TRANSLATE TRIMRANGE UNICHAR ' +
  'UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST').split(' '))
const XLWS = new Set(['FILTER', 'SORT'])
// Apply fn to the parts of a formula outside string literals.
const outsideStrings = (f, fn) => f.split(/("(?:[^"]|"")*")/).map((p, i) => (i % 2 ? p : fn(p))).join('')
export const formulaIn = (f) => outsideStrings(f, (p) => p.replace(/_xlfn\.(_xlws\.)?/gi, '').replace(/_xlpm\./gi, ''))
export const formulaOut = (f) => outsideStrings(f, (p) => p.replace(/(^|[^A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)(?=\()/g, (m, pre, name) => {
  const u = name.toUpperCase()
  return XLWS.has(u) ? `${pre}_xlfn._xlws.${name}` : XLFN.has(u) ? `${pre}_xlfn.${name}` : m
}))

// Dates come from ExcelJS as Date objects with millisecond precision; round off the float noise.
const serial = (d) => Math.round(((d.getTime() - EPOCH) / 864e5) * 1e10) / 1e10

function cellValueIn(cell) {
  const v = cell.value, o = {}
  if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
    const f = cell.formula // translated for shared-formula clones
    if (f) o.f = '=' + formulaIn(f)
    const res = v.result
    if (res instanceof Date) o.v = serial(res)
    else if (res && typeof res === 'object' && 'error' in res) o.v = res.error
    else if (res !== undefined) o.v = res
  } else if (v && typeof v === 'object' && 'richText' in v) o.v = v.richText.map((t) => t.text).join('')
  else if (v && typeof v === 'object' && 'error' in v) { o.v = v.error; o.t = 1 }
  else if (v && typeof v === 'object' && 'hyperlink' in v) o.v = v.text
  else if (v instanceof Date) o.v = serial(v)
  else o.v = v
  if (typeof o.v === 'number') o.t = 2
  else if (typeof o.v === 'boolean') { o.t = 3; o.v = o.v ? 1 : 0 }
  else if (typeof o.v === 'string' && o.t == null) o.t = 1
  return o
}

const resource = (snap, name) => {
  const r = snap.resources?.find((x) => x.name === name)
  try { return r?.data ? JSON.parse(r.data) : {} } catch { return {} }
}
const noteText = (n) => (typeof n === 'string' ? n : n?.texts ? n.texts.map((t) => t.text).join('') : '')
const a1 = (r, c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s + (r + 1) }
const parseA1 = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref ?? ''); if (!m) return null; return { r: +m[2] - 1, c: [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 } }

export async function xlsxToSnapshot(buf) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf)
  const theme = themePalette(wb)
  const sheets = {}, order = [], notes = {}, idByIndex = {}
  wb.eachSheet((ws, i) => {
    const id = `sheet${i}`, cellData = {}, columnData = {}, rowData = {}, mergeData = []
    idByIndex[order.length] = id
    const sheetNotes = {}
    ws.eachRow({ includeEmpty: true }, (row, r) => {
      const rowCells = {}
      // includeEmpty: cells with only a style (borders, fills on blank cells) are kept too.
      row.eachCell({ includeEmpty: true }, (cell, c) => {
        const s = styleIn(cell, theme)
        // Merged-away cells repeat the master's value in ExcelJS; Univer keeps them empty.
        if (cell.isMerged && cell.master !== cell) { if (!isEmptyStyle(s)) rowCells[c - 1] = { s }; return }
        const o = cellValueIn(cell)
        if (o.v == null) { delete o.v; delete o.t }
        if (!isEmptyStyle(s)) o.s = s
        if (Object.keys(o).length) rowCells[c - 1] = o
        if (cell.note) {
          const text = typeof cell.note === 'string' ? cell.note : (cell.note.texts ?? []).map((t) => t.text).join('')
          ;(sheetNotes[r - 1] ??= {})[c - 1] = { id: `n${i}_${r}_${c}`, row: r - 1, col: c - 1, note: text, ...NOTE_SIZE }
        }
      })
      if (Object.keys(rowCells).length) cellData[r - 1] = rowCells
      const d = {}
      if (row.hidden) d.hd = 1
      // Only heights set by hand are fixed; auto-fit rows (ht without customHeight) stay automatic in Univer.
      if (row.height && row.customHeight !== false) { d.h = Math.round(row.height * PX_PER_PT); d.ia = 0 }
      if (row.outlineLevel) d.custom = { outlineLevel: row.outlineLevel }
      if (row.collapsed) d.collapsed = true
      const rs = row.style && styleIn(row, theme)
      if (!isEmptyStyle(rs)) d.s = rs
      if (Object.keys(d).length) rowData[r - 1] = d
    })
    if (Object.keys(sheetNotes).length) notes[id] = sheetNotes
    ws.columns?.forEach((col, i2) => {
      const d = {}
      if (col.width) d.w = Math.round(col.width * 7)
      if (col.hidden) d.hd = 1
      if (col.outlineLevel) d.custom = { outlineLevel: col.outlineLevel }
      if (col.collapsed) d.collapsed = true
      const cs = col.style && styleIn(col, theme)
      if (!isEmptyStyle(cs)) d.s = cs
      if (Object.keys(d).length) columnData[i2] = d
    })
    const op = ws.properties?.outlineProperties ?? {}, view = ws.views?.[0] ?? {}
    const settings = outlineSettings({ summaryBelow: op.summaryBelow !== false, summaryRight: op.summaryRight !== false, hideSymbols: view.showOutlineSymbols === false })
    markCollapsed(rowData, settings.summaryBelow)
    markCollapsed(columnData, settings.summaryRight)
    Object.keys(ws._merges || {}).forEach((k) => {
      const m = ws._merges[k].model
      mergeData.push({ startRow: m.top - 1, endRow: m.bottom - 1, startColumn: m.left - 1, endColumn: m.right - 1 })
    })
    const sheet = { id, name: ws.name, cellData, columnData, rowData, mergeData, rowCount: 1048576, columnCount: 16384 }
    if (view.state === 'frozen' && (view.xSplit || view.ySplit)) {
      // startRow / startColumn: first scrolling row / column (the pane's top-left cell, later than the split
      // when the sheet was scrolled before freezing).
      const xSplit = view.xSplit || 0, ySplit = view.ySplit || 0, tl = parseA1(view.topLeftCell)
      sheet.freeze = { xSplit, ySplit, startRow: ySplit ? Math.max(ySplit, tl?.r ?? ySplit) : -1, startColumn: xSplit ? Math.max(xSplit, tl?.c ?? xSplit) : -1 }
    }
    if (view.showGridLines === false) sheet.showGridlines = 0
    const tab = colorIn(ws.properties?.tabColor, theme)
    if (tab) sheet.tabColor = tab
    if (ws.state === 'hidden') sheet.hidden = 1
    if (ws.state === 'veryHidden') sheet.hidden = 2
    if (!settings.isDefault) sheet.custom = { sheetlyOutline: settings.value }
    sheets[id] = sheet
    order.push(id)
  })
  const names = {}
  let n = 0
  for (const dn of wb.rawDefinedNames ?? []) {
    const text = (dn.text ?? (dn.ranges ?? []).join(',')).trim()
    if (!dn.name || dn.name.startsWith('_xlnm.') || !text) continue
    const key = `dn${n++}`
    names[key] = { id: key, name: dn.name, formulaOrRefString: formulaIn(text.replace(/^=/, '')),
      localSheetId: dn.localSheetId != null ? (idByIndex[dn.localSheetId] ?? ALL_SHEETS) : ALL_SHEETS }
  }
  const resources = []
  if (Object.keys(notes).length) resources.push({ name: NOTE, data: JSON.stringify(notes) })
  if (Object.keys(names).length) resources.push({ name: NAMES, data: JSON.stringify(names) })
  return { id: 'wb', sheetOrder: order, sheets, resources }
}

/**
 * snap: a Univer workbook snapshot (FWorkbook.save()). Cells that Univer stores as a shared formula id only
 * ({si} without f, e.g. after fill-down) need their formula text: pass resolveFormulas(snap, api) first, or a
 * formulaOf(sheetId, row, col) callback here.
 */
export async function snapshotToXlsx(snap, { formulaOf } = {}) {
  const wb = new ExcelJS.Workbook()
  const styleOf = (s) => (typeof s === 'string' ? snap.styles?.[s] : s) ?? null
  const notes = resource(snap, NOTE)
  for (const id of snap.sheetOrder) {
    const s = snap.sheets[id], settings = outlineSettings(s.custom?.sheetlyOutline)
    const view = {}
    if (s.freeze && (s.freeze.xSplit > 0 || s.freeze.ySplit > 0)) {
      const xSplit = s.freeze.xSplit || 0, ySplit = s.freeze.ySplit || 0
      Object.assign(view, { state: 'frozen', xSplit, ySplit,
        topLeftCell: a1(ySplit ? Math.max(ySplit, s.freeze.startRow) : 0, xSplit ? Math.max(xSplit, s.freeze.startColumn) : 0) })
    }
    if (s.showGridlines === 0) view.showGridLines = false
    if (settings.hideSymbols) view.showOutlineSymbols = false
    const props = {}
    if (argb(s.tabColor)) props.tabColor = { argb: argb(s.tabColor) }
    if (!settings.summaryBelow || !settings.summaryRight) props.outlineProperties = { summaryBelow: settings.summaryBelow, summaryRight: settings.summaryRight }
    const ws = wb.addWorksheet(s.name, { views: [view], properties: props })
    if (s.hidden === 2) ws.state = 'veryHidden'
    else if (s.hidden) ws.state = 'hidden'
    const rowStyle = (r) => styleOf(s.rowData?.[r]?.s), colStyle = (c) => styleOf(s.columnData?.[c]?.s)
    // Row and column styles first: ExcelJS gives new cells in them that style as their default.
    for (const [c, col] of Object.entries(s.columnData || {})) { const st = styleOf(col?.s); if (!isEmptyStyle(st)) ws.getColumn(+c + 1).style = styleOut(st) }
    for (const [r, row] of Object.entries(s.rowData || {})) { const st = styleOf(row?.s); if (!isEmptyStyle(st)) ws.getRow(+r + 1).style = styleOut(mergeStyles(st)) }
    for (const [r, row] of Object.entries(s.cellData || {})) for (const [c, cell] of Object.entries(row ?? {})) {
      if (!cell) continue
      const x = ws.getCell(+r + 1, +c + 1)
      const f = cell.f || (cell.si ? formulaOf?.(id, +r, +c) : null)
      if (f) x.value = { formula: formulaOut(f.replace(/^=/, '')), result: typeof cell.v === 'string' && ERRORS.has(cell.v) ? { error: cell.v } : cell.v ?? undefined }
      else if (typeof cell.v === 'string' && ERRORS.has(cell.v) && cell.t !== 1) x.value = { error: cell.v }
      else if (cell.v !== undefined && cell.v !== null && cell.v !== '') x.value = cell.t === 3 ? !!cell.v : cell.v
      else if (cell.p?.body?.dataStream) x.value = cell.p.body.dataStream.replace(/\r?\n$/, '').replace(/\r/g, '\n')
      const st = styleOf(cell.s)
      if (st || rowStyle(r) || colStyle(c)) {
        const style = styleOut(mergeStyles(colStyle(c), rowStyle(r), st))
        for (const k of ['font', 'fill', 'border', 'alignment', 'numFmt']) if (style[k]) x[k] = style[k]
      }
      const note = notes[id]?.[r]?.[c]
      if (note) x.note = noteText(note.note)
    }
    for (const [c, col] of Object.entries(s.columnData || {})) if (col?.w) ws.getColumn(+c + 1).width = col.w / 7
    let maxRow = 0, maxCol = 0
    for (const [r, row] of Object.entries(s.rowData || {})) {
      if (!row) continue
      const L = row.custom?.outlineLevel ?? 0
      if (L) { ws.getRow(+r + 1).outlineLevel = L; maxRow = Math.max(maxRow, L) }
      if (L || row.custom?.outlineCollapsed) ws.getRow(+r + 1).collapsed = !!row.custom?.outlineCollapsed
      if (row.hd) ws.getRow(+r + 1).hidden = true
      if (row.h && row.ia === 0) ws.getRow(+r + 1).height = Math.round((row.h / PX_PER_PT) * 100) / 100
    }
    for (const [c, col] of Object.entries(s.columnData || {})) {
      if (!col) continue
      const L = col.custom?.outlineLevel ?? 0
      if (L) { ws.getColumn(+c + 1).outlineLevel = L; maxCol = Math.max(maxCol, L) }
      if (L || col.custom?.outlineCollapsed) ws.getColumn(+c + 1).collapsed = !!col.custom?.outlineCollapsed
      if (col.hd) ws.getColumn(+c + 1).hidden = true
    }
    if (maxRow) ws.properties.outlineLevelRow = maxRow
    if (maxCol) ws.properties.outlineLevelCol = maxCol
    for (const m of s.mergeData || []) ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1)
  }
  // Defined names go out as written (ranges, constants, formulas; sheet-scoped ones with localSheetId).
  const sheetIndex = Object.fromEntries(snap.sheetOrder.map((id, i) => [id, i]))
  wb.extraDefinedNames = Object.values(resource(snap, NAMES)).filter((dn) => dn?.name && dn.formulaOrRefString).map((dn) => {
    const local = dn.localSheetId && dn.localSheetId !== ALL_SHEETS ? sheetIndex[dn.localSheetId] : undefined
    return { name: dn.name, ranges: [formulaOut(dn.formulaOrRefString.replace(/^=/, ''))], ...(local != null ? { localSheetId: local } : {}) }
  })
  return new Uint8Array(await wb.xlsx.writeBuffer())
}

/** Fills in the formula text of shared-formula cells ({si} only) from the live workbook (FWorkbook). */
export function formulaResolver(fworkbook) {
  return (sheetId, r, c) => fworkbook.getSheetBySheetId(sheetId)?.getRange(r, c).getFormula() || null
}

// A group is collapsed when its summary row/column carries Excel's collapsed flag (the group one level deeper than
// the summary itself) or when its detail is all hidden: set that level's bit on the summary (after the group when
// summaries are below / right, before it when above / left).
function markCollapsed(data, after = true) {
  const levels = new Map()
  for (const [i, d] of Object.entries(data)) if (d.custom?.outlineLevel) levels.set(+i, d.custom.outlineLevel)
  for (const g of outlineGroups(levels)) {
    const at = after ? g.end + 1 : g.start - 1
    if (at < 0) continue
    const summary = data[at]
    let hidden = true
    for (let i = g.start; i <= g.end && hidden; i++) hidden = data[i]?.hd === 1
    const flagged = summary?.collapsed && g.level === (summary.custom?.outlineLevel ?? 0) + 1
    if (!hidden && !flagged) continue
    const s = (data[at] ??= {})
    s.custom = { ...(s.custom ?? {}), outlineCollapsed: (s.custom?.outlineCollapsed ?? 0) | (1 << g.level) }
  }
  for (const d of Object.values(data)) delete d.collapsed
}
