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
import { CF, CHART_COMPONENT, DRAWINGS, DV, REF_PREFIX, TABLES, rangeA1, cfIn, cfOut, dvIn, dvOut, endFrom, imagesIn, imagesOut, parseRange, placed, tablesIn, tablesOut } from './xlsx-objects.js'
import { finishPackage, readCharts, readDynamicArrays, readSparklines } from './xlsx-zip.js'
import { resolveSource } from './chart-model.js'
import { EXCEL_FUNCTIONS } from './excel-function-names.js'

const NOTE = 'SHEET_NOTE_PLUGIN', NAMES = 'SHEET_DEFINED_NAME_PLUGIN', ALL_SHEETS = 'AllDefaultWorkbook'
const EPOCH = Date.UTC(1899, 11, 30), EPOCH_1904 = Date.UTC(1904, 0, 1)
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
// LET / LAMBDA parameter names are stored as _xlpm.name inside their call (Excel reads a bare name as a defined name).
// The formula is scanned with its string literals, quoted sheet names and structured references ([...]) masked, so
// text such as "LET(x,1,x)" in quotes, 'Sales x Plan'!A1 or Table1[x] is never taken for a call or a parameter.
function mask(f) {
  let out = '', inStr = false, inSheet = false, depth = 0
  for (let i = 0; i < f.length; i++) {
    const ch = f[i]
    if (inStr) { out += '\u0000'; if (ch === '"') { if (f[i + 1] === '"') { out += '\u0000'; i++ } else inStr = false } continue }
    // 'Quoted sheet names' ('' is an apostrophe inside one)
    if (inSheet) { out += '\u0000'; if (ch === "'") { if (f[i + 1] === "'") { out += '\u0000'; i++ } else inSheet = false } continue }
    if (ch === "'") { inSheet = true; out += '\u0000'; continue }
    if (depth) { out += '\u0000'; if (ch === '[') depth++; else if (ch === ']') depth--; continue }
    if (ch === '"') { inStr = true; out += '\u0000'; continue }
    if (ch === '[') { depth = 1; out += '\u0000'; continue }
    out += ch
  }
  return out
}
function xlpm(f) {
  const edits = [] // [position, name] where '_xlpm.' goes
  const m0 = mask(f)
  const call = /(^|[^A-Za-z0-9_.])(?:_xlfn\.)?(LET|LAMBDA)\(/gi
  let m
  while ((m = call.exec(m0))) {
    const open = m.index + m[0].length - 1
    const commas = []
    let depth = 0, close = -1
    for (let i = open; i < m0.length; i++) {
      const ch = m0[i]
      if (ch === '(' || ch === '{') depth++
      else if (ch === ')' || ch === '}') { depth--; if (depth === 0) { close = i; break } }
      else if (ch === ',' && depth === 1) commas.push(i)
    }
    if (close < 0) break
    const bounds = [open, ...commas, close]
    const args = bounds.slice(0, -1).map((b, k) => m0.slice(b + 1, bounds[k + 1]).trim())
    const isLet = m[2].toUpperCase() === 'LET'
    const valid = (x) => /^[A-Za-z_][A-Za-z0-9_.]*$/.test(x) && !/^_xlpm\./i.test(x)
    // Declarations with their argument positions (LET: names at even positions; LAMBDA: all but the body).
    const decls = args.map((x, k) => [x, k]).filter(([x, k]) => k < args.length - 1 && (!isLet || k % 2 === 0) && valid(x))
    // Names that can be called: LET names bound to a LAMBDA, and LAMBDA parameters (a function passed in) unless
    // they are spelled like a built-in function, which a call then means.
    // A call to a bound name means the binding unless the name is spelled like a built-in function and the binding
    // cannot hold a function (LET(sum,5,SUM(sum,1)) calls the function SUM). A LET value holds a function when it
    // contains a LAMBDA or names an earlier binding that does (an alias).
    const builtin = (x) => EXCEL_FUNCTIONS.has(x.toUpperCase()) || XLFN.has(x.toUpperCase()) || XLWS.has(x.toUpperCase())
    const holdsFunction = new Set()
    if (isLet) for (let k = 0; k < args.length - 1; k += 2) {
      const value = (args[k + 1] ?? '').trim()
      if (/LAMBDA\(/i.test(value) || holdsFunction.has(value.toLowerCase())) holdsFunction.add(args[k].toLowerCase())
    }
    const callable = (x) => !builtin(x) || holdsFunction.has(x.toLowerCase())
    // Whole names only: letters of any script count as part of a name (税x is not x).
    const re = (name, calls) => new RegExp(`(^|[^\\p{L}\\p{N}_.!$'\u0000])(${name.replace(/\./g, '\\.')})${calls ? '(?=\\()' : "(?![\\p{L}\\p{N}_.(!\u0000])"}`, 'giu')
    const scan = (name, startAt, endAt, calls) => { const span = m0.slice(startAt, endAt); let t; const r = re(name, calls); while ((t = r.exec(span))) edits.push(startAt + t.index + t[1].length) }
    // Where each name is bound: a LAMBDA parameter in the whole call; a LET name from its declaration on, skipping
    // its own value (LET(a,y,y,1,a+y): the first y is the workbook's y, not the parameter).
    for (const [name, k] of decls) {
      if (isLet) {
        const declAt = bounds[k] + 1, decl = m0.slice(declAt, bounds[k + 1])
        edits.push(declAt + (decl.length - decl.trimStart().length))
        const from = bounds[k + 2] ?? close
        scan(name, from, close + 1, false)
        if (callable(name, k)) scan(name, from, close + 1, true)
      } else {
        scan(name, open, close + 1, false)
        if (callable(name, k)) scan(name, open, close + 1, true)
      }
    }
  }
  // Apply from the end so earlier positions stay valid; a position is prefixed once.
  for (const pos of [...new Set(edits)].sort((x, y) => y - x)) f = f.slice(0, pos) + '_xlpm.' + f.slice(pos)
  return f
}
export const formulaOut = (f) => xlpmThen(f)
const xlpmThen = (f) => xlpm(fnPrefix(f))
const fnPrefix = (f) => outsideStrings(f, (p) => p.replace(/(^|[^A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)(?=\()/g, (m, pre, name) => {
  const u = name.toUpperCase()
  return XLWS.has(u) ? `${pre}_xlfn._xlws.${name}` : XLFN.has(u) ? `${pre}_xlfn.${name}` : m
}))

// Dates come from ExcelJS as Date objects with millisecond precision; round off the float noise.
// epoch: the workbook's own (1904 date system workbooks count days from 1904-01-01).
const serial = (d, epoch = EPOCH) => Math.round(((d.getTime() - epoch) / 864e5) * 1e10) / 1e10

function cellValueIn(cell, epoch = EPOCH) {
  const v = cell.value, o = {}
  if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
    const f = cell.formula // translated for shared-formula clones
    if (f) o.f = '=' + formulaIn(f)
    const res = v.result
    if (res instanceof Date) o.v = serial(res, epoch)
    else if (res && typeof res === 'object' && 'error' in res) o.v = res.error
    else if (res !== undefined) o.v = res
  } else if (v && typeof v === 'object' && 'richText' in v) o.v = v.richText.map((t) => t.text).join('')
  else if (v && typeof v === 'object' && 'error' in v) { o.v = v.error; o.t = 1 }
  else if (v && typeof v === 'object' && 'hyperlink' in v) { o.v = typeof v.text === 'string' ? v.text : v.text?.richText?.map((t) => t.text).join('') ?? ''; o.link = v.hyperlink }
  else if (v instanceof Date) o.v = serial(v, epoch)
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
const docText = (p) => p.body.dataStream.replace(/\r?\n$/, '').replace(/\r/g, '\n')
const linkOf = (p) => p?.body?.customRanges?.find((r) => r.rangeType === 0 && r.properties?.url)?.properties.url
let linkSeq = 0
/** A one-paragraph Univer document whose whole text links to url. */
function linkDoc(text, url) {
  const id = `link${++linkSeq}`
  return { id: `d_${id}`, documentStyle: {}, body: { dataStream: `${text}\r\n`, textRuns: [],
    paragraphs: [{ startIndex: text.length }], sectionBreaks: [{ startIndex: text.length + 1 }],
    customRanges: [{ startIndex: 0, endIndex: Math.max(0, text.length - 1), rangeType: 0, rangeId: id, properties: { url } }] } }
}
const noteText = (n) => (typeof n === 'string' ? n : n?.texts ? n.texts.map((t) => t.text).join('') : '')
const a1 = (r, c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s + (r + 1) }
const parseA1 = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref ?? ''); if (!m) return null; return { r: +m[2] - 1, c: [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 } }

export async function xlsxToSnapshot(buf) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf)
  const theme = themePalette(wb)
  // 1904 workbooks keep their serials; Univer is told the date system.
  const date1904 = !!wb.properties?.date1904
  const epoch = date1904 ? EPOCH_1904 : EPOCH
  const sheets = {}, order = [], notes = {}, idByIndex = {}, links = [], idByName = {}
  const cfs = {}, dvs = {}, drawings = {}, tables = []
  const dynamic = await readDynamicArrays(buf), legacyArrays = []
  wb.eachSheet((ws, i) => {
    idByName[ws.name] = `sheet${i}`
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
        const o = cellValueIn(cell, epoch)
        if (o.link) { links.push([o, o.link]); delete o.link }
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
    // Array formulas. A dynamic array (cm= in the file) holds its spilled values as plain cells inside the anchor's
    // ref; Univer computes them itself (kept as constants they would make the formula #SPILL!). A legacy array
    // formula (Ctrl+Shift+Enter) is read the same way and listed in sheetlyLegacyArrays: once the workbook has been
    // calculated, one whose result is a single value gets the formula in every cell of its range (main.js).
    const sheetIndex = order.length
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value
      if (!v || typeof v !== 'object' || v.shareType !== 'array' || !v.ref) return
      const g = parseRange(v.ref)
      if (!g) return
      const r0 = cell.row - 1, c0 = cell.col - 1, anchor = cellData[r0]?.[c0]
      const others = []
      for (let r = g.startRow; r <= g.endRow; r++) for (let c = g.startColumn; c <= g.endColumn; c++) if (r !== r0 || c !== c0) others.push([r, c])
      if (!dynamic.has(`${sheetIndex}!${cell.address}`) && anchor?.f) legacyArrays.push({ sheetId: id, row: r0, col: c0, range: g, formula: anchor.f })
      for (const [r, c] of others) {
        const x = cellData[r]?.[c]
        if (!x || x.f) continue
        delete x.v; delete x.t
        if (!Object.keys(x).length) delete cellData[r][c]
      }
    }))
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
    // Sheet defaults from <sheetFormatPr> (Excel's 15 pt rows and 8.43-character columns when absent), Calibri 11.
    const fp = ws.properties ?? {}
    const sheet = { id, name: ws.name, cellData, columnData, rowData, mergeData, rowCount: 1048576, columnCount: 16384,
      defaultRowHeight: Math.round((fp.defaultRowHeight || 15) * PX_PER_PT), defaultColumnWidth: fp.defaultColWidth ? Math.round(fp.defaultColWidth * 7) : 64,
      defaultStyle: { ff: 'Calibri', fs: 11 }, custom: { sheetlyDefaults: true } }
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
    if (!settings.isDefault) sheet.custom.sheetlyOutline = settings.value
    sheets[id] = sheet
    order.push(id)
    const rules = cfIn(ws, id, (c) => colorIn(c, theme))
    if (rules.length) cfs[id] = rules
    const vals = dvIn(ws, id)
    if (vals.length) dvs[id] = vals
    // Tables with their header row hidden keep their definition (column names live only there): Univer's tables
    // would name the columns from the first data row. They are written back as they were.
    for (const t of tablesIn(ws)) {
      if (t.showHeader) tables.push({ sheetId: id, ...t })
      else (sheet.custom.sheetlyTablesNoHeader ??= []).push({ name: t.name, range: t.range, showFooter: t.showFooter, columns: t.columns })
    }
    const pics = imagesIn(wb, ws, sheet, 'wb')
    if (pics.length) drawings[id] = { data: Object.fromEntries(pics.map((d) => [d.drawingId, d])), order: pics.map((d) => d.drawingId) }
  })
  // Object ranges are kept as hidden defined names (see refs.js) so they move with structural edits.
  const refNames = {}
  let refSeq = 0
  const track = (sheetName, range) => {
    if (!range) return undefined
    const name = `${REF_PREFIX}f${(++refSeq).toString(36)}`
    const q = /^[A-Za-z_][\w.]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`
    refNames[name] = { id: name, name, formulaOrRefString: `${q}!${rangeA1(range).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`, hidden: true, localSheetId: ALL_SHEETS }
    return name
  }
  // Sparklines (x14 extension of the sheet).
  for (const { sheetIndex, groups } of await readSparklines(buf)) {
    const id = idByIndex[sheetIndex]
    if (!id) continue
    sheets[id].custom.sheetlySparklines = groups.map((g, k) => ({ id: `spk${k}`, type: g.type, colors: g.colors,
      items: g.items.filter((x) => idByName[x.sheetName]).map((x) => ({ row: x.row, col: x.col, sheetId: idByName[x.sheetName], source: x.source,
        cellRef: track(sheets[id].name, { startRow: x.row, endRow: x.row, startColumn: x.col, endColumn: x.col }), sourceRef: track(x.sheetName, x.source) })) })).filter((g) => g.items.length)
  }
  // Charts: from the chart parts, as Sheetly charts linked to the same range.
  let chartSeq = 0
  for (const c of await readCharts(buf)) {
    const id = idByIndex[c.sheetIndex], sourceId = idByName[c.spec.source.sheetName]
    if (!id || !sourceId) continue
    const { sheetName, ...source } = c.spec.source
    source.ref = track(sheetName, source.range)
    if (source.series) {
      source.series = source.series.map((x) => ({ ...x, valuesRef: track(sheetName, x.values), xValuesRef: track(sheetName, x.xValues),
        nameRef: x.name ? track(sheetName, { startRow: x.name[0], endRow: x.name[0], startColumn: x.name[1], endColumn: x.name[1] }) : undefined }))
      source.categoriesRef = track(sheetName, source.categories)
    }
    const spec = { ...c.spec, source: { ...source, sheetId: sourceId, local: sourceId === id } }
    const drawingId = `chart${++chartSeq}_${Date.now().toString(36)}`
    const d = placed(sheets[id], c.from, c.to ?? endFrom(sheets[id], c.from, c.size.width, c.size.height), { unitId: 'wb', subUnitId: id, drawingId, drawingType: 8, componentKey: CHART_COMPONENT, data: spec, allowTransform: true })
    const entry = (drawings[id] ??= { data: {}, order: [] })
    entry.data[drawingId] = d; entry.order.push(drawingId)
  }
  // Hyperlinks become Univer link text: a rich-text cell whose text carries the URL; links into the workbook
  // ("#Sheet2!A1") use Univer's "#gid=<sheet id>&range=A1" form.
  for (const [o, url] of links) {
    // '#Sheet'!A1 or #Sheet!A1; a quoted name doubles its apostrophes ('O''Brien').
    const m = /^#(?:'((?:[^']|'')+)'|([^!']+))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/.exec(url)
    const sheetName = m ? (m[1] ?? m[2]).replace(/''/g, "'") : null
    const target = m && idByName[sheetName] ? `#gid=${idByName[sheetName]}&range=${m[3].replace(/\$/g, '')}` : url
    o.p = linkDoc(String(o.v ?? url), target)
  }
  const names = {}
  let n = 0
  for (const dn of wb.rawDefinedNames ?? []) {
    const text = (dn.text ?? (dn.ranges ?? []).join(',')).trim()
    if (!dn.name || dn.name.startsWith('_xlnm.') || !text) continue
    const key = `dn${n++}`
    names[key] = { id: key, name: dn.name, formulaOrRefString: formulaIn(text.replace(/^=/, '')),
      localSheetId: dn.localSheetId != null ? (idByIndex[dn.localSheetId] ?? ALL_SHEETS) : ALL_SHEETS }
  }
  Object.assign(names, refNames)
  const resources = []
  if (Object.keys(notes).length) resources.push({ name: NOTE, data: JSON.stringify(notes) })
  if (Object.keys(names).length) resources.push({ name: NAMES, data: JSON.stringify(names) })
  if (Object.keys(cfs).length) resources.push({ name: CF, data: JSON.stringify(cfs) })
  if (Object.keys(dvs).length) resources.push({ name: DV, data: JSON.stringify(dvs) })
  if (Object.keys(drawings).length) resources.push({ name: DRAWINGS, data: JSON.stringify(drawings) })
  // Tables are created through Univer's table API once the workbook exists (sheetlyTables, see main.js).
  return { id: 'wb', sheetOrder: order, sheets, resources, ...(date1904 ? { dateSystem: 'date1904' } : {}), ...(tables.length ? { sheetlyTables: tables } : {}),
    ...(legacyArrays.length ? { sheetlyLegacyArrays: legacyArrays } : {}) }
}

/**
 * snap: a Univer workbook snapshot (FWorkbook.save()). Cells that Univer stores as a shared formula id only
 * ({si} without f, e.g. after fill-down) need their formula text: pass resolveFormulas(snap, api) first, or a
 * formulaOf(sheetId, row, col) callback here.
 */
export async function snapshotToXlsx(snap, { formulaOf, arrayRangeOf } = {}) {
  const wb = new ExcelJS.Workbook()
  const styleOf = (s) => (typeof s === 'string' ? snap.styles?.[s] : s) ?? null
  const notes = resource(snap, NOTE), cfs = resource(snap, CF), dvs = resource(snap, DV), allDrawings = resource(snap, DRAWINGS), allTables = resource(snap, TABLES)
  if (snap.dateSystem === 'date1904') wb.properties.date1904 = true
  // Excel recalculates on open: spilled ranges and anything Sheetly could not cache are filled in.
  wb.calcProperties = { fullCalcOnLoad: true }
  const charts = [], spills = new Set(), usedTableNames = new Set(), sparklines = []
  // Hidden names that keep object ranges current (refs.js): where they point now.
  const nameMap = Object.fromEntries(Object.values(resource(snap, NAMES)).filter((d) => d?.name).map((d) => [d.name, d.formulaOrRefString]))
  const sheetIdByName = Object.fromEntries(snap.sheetOrder.map((id) => [snap.sheets[id].name, id]))
  const resolveName = (name) => {
    const text = nameMap[name]
    if (!text || /#REF!/i.test(text)) return null
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(text)
    const sheetId = m && sheetIdByName[(m[1] ?? m[2]).replace(/''/g, "'")], range = m && parseRange(m[3])
    return sheetId && range ? { sheetId, range } : null
  }
  for (const [sheetIndex, id] of snap.sheetOrder.entries()) {
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
    if (s.defaultRowHeight && s.defaultRowHeight !== 20) props.defaultRowHeight = Math.round((s.defaultRowHeight / PX_PER_PT) * 100) / 100
    if (s.defaultColumnWidth && s.defaultColumnWidth !== 64) props.defaultColWidth = Math.round((s.defaultColumnWidth / 7) * 100) / 100
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
      const spill = f ? arrayRangeOf?.(id, +r, +c) : null
      const result = typeof cell.v === 'string' && ERRORS.has(cell.v) ? { error: cell.v } : cell.v ?? undefined
      if (f && spill && (spill.endRow > spill.startRow || spill.endColumn > spill.startColumn)) {
        // A formula that spills: Excel's dynamic array (t="array" over the spill range, cm="1" added by finishPackage).
        x.value = { formula: formulaOut(f.replace(/^=/, '')), result, shareType: 'array', ref: `${a1(spill.startRow, spill.startColumn)}:${a1(spill.endRow, spill.endColumn)}` }
        spills.add(`${sheetIndex}!${a1(+r, +c)}`)
      } else if (f) x.value = { formula: formulaOut(f.replace(/^=/, '')), result }
      else if (typeof cell.v === 'string' && ERRORS.has(cell.v) && cell.t !== 1) x.value = { error: cell.v }
      else if (cell.v !== undefined && cell.v !== null && cell.v !== '') x.value = cell.t === 3 ? !!cell.v : cell.v
      else if (cell.p?.body?.dataStream) x.value = docText(cell.p)
      const url = linkOf(cell.p)
      if (url && !f) {
        const m = /^#gid=([^&]+)&range=(.+)$/.exec(url), target = m && snap.sheets[m[1]] ? `#'${snap.sheets[m[1]].name.replace(/'/g, "''")}'!${m[2]}` : url
        x.value = { text: cell.p?.body?.dataStream ? docText(cell.p) : String(cell.v ?? url), hyperlink: target }
      }
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
    const sg = (s.custom?.sheetlySparklines ?? []).map((g) => ({ ...g, items: g.items.map((x) => {
      const at = x.cellRef ? resolveName(x.cellRef) : { sheetId: id, range: { startRow: x.row, startColumn: x.col } }
      const src = x.sourceRef ? resolveName(x.sourceRef) : { sheetId: x.sheetId ?? id, range: x.source }
      return at && src ? { row: at.range.startRow, col: at.range.startColumn, source: src.range, sheetName: snap.sheets[src.sheetId]?.name ?? s.name } : null
    }).filter(Boolean) })).filter((g) => g.items.length)
    if (sg.length) sparklines.push({ sheetIndex, groups: sg })
    cfOut(ws, cfs[id])
    dvOut(ws, dvs[id])
    tablesOut(ws, [...(allTables[id]?.tables ?? []), ...(s.custom?.sheetlyTablesNoHeader ?? []).map((t) => ({ name: t.name, range: t.range,
      options: { showHeader: false, showFooter: t.showFooter }, columns: (t.columns ?? []).map((n) => ({ displayName: n })) }))], usedTableNames)
    const list = Object.values(allDrawings[id]?.data ?? {})
    imagesOut(wb, ws, list)
    for (const d of list) {
      if (d.componentKey !== CHART_COMPONENT || !d.data?.source || !d.sheetTransform) continue
      const source = resolveSource(d.data.source, resolveName)
      const src = source && snap.sheets[source.sheetId]
      if (!src) continue
      const get = (r, c) => { const v = src.cellData?.[r]?.[c]; return v?.v ?? (v?.p?.body?.dataStream ? docText(v.p) : null) }
      charts.push({ sheetIndex, spec: { ...d.data, source }, from: d.sheetTransform.from, to: d.sheetTransform.to, get, sourceName: src.name })
    }
  }
  // Defined names go out as written (ranges, constants, formulas; sheet-scoped ones with localSheetId).
  const sheetIndex = Object.fromEntries(snap.sheetOrder.map((id, i) => [id, i]))
  wb.extraDefinedNames = Object.values(resource(snap, NAMES)).filter((dn) => dn?.name && dn.formulaOrRefString && !dn.name.startsWith(REF_PREFIX)).map((dn) => {
    const local = dn.localSheetId && dn.localSheetId !== ALL_SHEETS ? sheetIndex[dn.localSheetId] : undefined
    return { name: dn.name, ranges: [formulaOut(dn.formulaOrRefString.replace(/^=/, ''))], ...(local != null ? { localSheetId: local } : {}) }
  })
  const buf = new Uint8Array(await wb.xlsx.writeBuffer())
  // Charts are written into the sheet named by their own source; chartXml uses that sheet's name in the refs.
  return finishPackage(buf, { charts: charts.map((c) => ({ ...c, sheetName: c.sourceName })), spills, sparklines })
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
