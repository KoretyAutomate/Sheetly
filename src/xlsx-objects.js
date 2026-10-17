// .xlsx <-> Univer for sheet objects that live outside the cells: conditional formats, data validation, tables,
// pictures, and the anchors of charts (the chart parts themselves are in xlsx-zip.js). Used by xlsx.js.
//
// Univer keeps these as workbook resources:
//   SHEET_CONDITIONAL_FORMATTING_PLUGIN {sheetId: [{cfId, ranges, stopIfTrue, rule}]}
//   SHEET_DATA_VALIDATION_PLUGIN       {sheetId: [{uid, ranges, type, operator, formula1, formula2, ...}]}
//   SHEET_TABLE_PLUGIN                 {sheetId: {tables: [{id, name, range, options, columns}]}}
//   SHEET_DRAWING_PLUGIN               {sheetId: {data: {id: drawing}, order: [id]}}
export const CF = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', DV = 'SHEET_DATA_VALIDATION_PLUGIN', TABLES = 'SHEET_TABLE_PLUGIN', DRAWINGS = 'SHEET_DRAWING_PLUGIN'
export const CHART_COMPONENT = 'sheetly-chart'
// Internal hidden defined names that keep object ranges up to date (see refs.js); never written to .xlsx.
export const REF_PREFIX = 'SheetlyRef_'
const EMU_PER_PX = 9525

// ---- A1 helpers
export const colName = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s }
const colNum = (L) => [...L.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
export const rangeA1 = (g) => (g.startRow === g.endRow && g.startColumn === g.endColumn ? `${colName(g.startColumn)}${g.startRow + 1}`
  : `${colName(g.startColumn)}${g.startRow + 1}:${colName(g.endColumn)}${g.endRow + 1}`)
export function parseRange(a1) {
  // (references are case-insensitive: a1:b3 is A1:B3)
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(String(a1).trim().replace(/^'?[^!]*'?!/, ''))
  if (!m) return null
  const r0 = +m[2] - 1, c0 = colNum(m[1]), r1 = m[4] ? +m[4] - 1 : r0, c1 = m[3] ? colNum(m[3]) : c0
  return { startRow: Math.min(r0, r1), endRow: Math.max(r0, r1), startColumn: Math.min(c0, c1), endColumn: Math.max(c0, c1) }
}
const absA1 = (g) => rangeA1(g).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')
const sqrefIn = (sqref) => String(sqref).split(/\s+/).map(parseRange).filter(Boolean)
const sqrefOut = (ranges) => ranges.map(rangeA1).join(' ')
const plainRange = (g) => ({ startRow: g.startRow, endRow: g.endRow, startColumn: g.startColumn, endColumn: g.endColumn })

// ---- colours / dxf styles (Univer IStyleBase <-> ExcelJS dxf)
const hex = (c) => { if (typeof c !== 'string') return undefined; const s = c.trim().toLowerCase(); let m = s.match(/^#([0-9a-f]{3})$/); if (m) return [...m[1]].map((x) => x + x).join('')
  m = s.match(/^#([0-9a-f]{6})/); if (m) return m[1]; m = s.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/); return m ? m.slice(1, 4).map((x) => (+x).toString(16).padStart(2, '0')).join('') : undefined }
const argb = (c) => { const h = hex(c); return h ? 'FF' + h.toUpperCase() : undefined }
const fromArgb = (c) => (c?.argb && /^[0-9A-F]{8}$/i.test(c.argb) ? '#' + c.argb.slice(2).toLowerCase() : undefined)
function dxfOut(st = {}) {
  const font = {}
  if (argb(st.cl?.rgb)) font.color = { argb: argb(st.cl.rgb) }
  if (st.bl) font.bold = true
  if (st.it) font.italic = true
  if (st.st?.s) font.strike = true
  if (st.ul?.s) font.underline = true
  const out = {}
  if (Object.keys(font).length) out.font = font
  if (argb(st.bg?.rgb)) out.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(st.bg.rgb) }, fgColor: { argb: argb(st.bg.rgb) } }
  return out
}
function dxfIn(style = {}, color = fromArgb) {
  const s = {}, f = style.font ?? {}, fill = style.fill
  const cl = color(f.color); if (cl) s.cl = { rgb: cl }
  if (f.bold) s.bl = 1
  if (f.italic) s.it = 1
  if (f.strike) s.st = { s: 1 }
  if (f.underline) s.ul = { s: 1 }
  // A dxf's solid fill colour is bgColor (fgColor when a writer put it there).
  const bg = color(fill?.bgColor) ?? color(fill?.fgColor); if (bg) s.bg = { rgb: bg }
  return s
}

// ---- conditional formatting
const cfvoOut = (v) => ({ type: v?.type ?? 'min', ...(v?.value !== undefined && v?.value !== '' ? { value: typeof v.value === 'string' ? v.value.replace(/^=/, '') : v.value } : {}) })
const cfvoIn = (v) => ({ type: ['num', 'min', 'max', 'percent', 'percentile', 'formula'].includes(v?.type) ? v.type : 'num',
  ...(v?.value !== undefined ? { value: v.type === 'formula' ? `=${v.value}` : v.value } : {}) })

/** Univer rules of one sheet -> ExcelJS ws.addConditionalFormatting calls. */
export function cfOut(ws, rules = []) {
  let priority = 1
  for (const r of rules) {
    const ref = sqrefOut(r.ranges ?? []), x = r.rule
    if (!ref || !x) continue
    const tl = r.ranges[0], cell = `${colName(tl.startColumn)}${tl.startRow + 1}`, abs = `$${colName(tl.startColumn)}$${tl.startRow + 1}:$${colName(r.ranges[0].endColumn)}$${r.ranges[0].endRow + 1}`
    let rule = null
    if (x.type === 'highlightCell') {
      const style = dxfOut(x.style)
      const lit = (v) => (typeof v === 'number' ? String(v) : `"${String(v ?? '').replace(/"/g, '""')}"`)
      switch (x.subType) {
        case 'number': rule = { type: 'cellIs', operator: x.operator, formulae: Array.isArray(x.value) ? x.value.map(String) : [String(x.value ?? 0)], style }; break
        case 'text':
          rule = ['equal', 'notEqual'].includes(x.operator) ? { type: 'cellIs', operator: x.operator, formulae: [lit(x.value ?? '')], style }
            : { type: 'containsText', operator: x.operator, text: x.value ?? '', style }
          break
        case 'timePeriod': rule = { type: 'timePeriod', timePeriod: x.operator, style }; break
        case 'rank': rule = { type: 'top10', rank: x.value ?? 10, percent: !!x.isPercent, bottom: !!x.isBottom, style }; break
        case 'average':
          rule = ['equal', 'notEqual'].includes(x.operator) ? { type: 'expression', formulae: [`${cell}${x.operator === 'equal' ? '=' : '<>'}AVERAGE(${abs})`], style }
            : { type: 'aboveAverage', aboveAverage: x.operator.startsWith('greater'), equalAverage: x.operator.endsWith('OrEqual'), style }
          break
        case 'formula': rule = x.sheetlyStdDev ? { type: 'aboveAverage', aboveAverage: x.sheetlyStdDev.above, equalAverage: x.sheetlyStdDev.equal, stdDev: x.sheetlyStdDev.n, style }
          : { type: 'expression', formulae: [String(x.value ?? '').replace(/^=/, '')], style }; break
        case 'duplicateValues': case 'uniqueValues': rule = { type: x.subType, style }; break
        default: break
      }
    } else if (x.type === 'dataBar') {
      // gradient bars are Excel's default (x14 extension); solid bars say gradient="0".
      rule = { type: 'dataBar', cfvo: [cfvoOut(x.config?.min), cfvoOut(x.config?.max)], color: { argb: argb(x.config?.positiveColor) ?? 'FF638EC6' },
        ...(x.config?.isGradient === false ? { gradient: false } : {}), ...(x.isShowValue === false ? { showValue: false } : {}) }
    } else if (x.type === 'colorScale') {
      const list = [...(x.config ?? [])].sort((a, b) => a.index - b.index)
      rule = { type: 'colorScale', cfvo: list.map((c) => cfvoOut(c.value)), color: list.map((c) => ({ argb: argb(c.color) ?? 'FF000000' })) }
    } else if (x.type === 'iconSet') {
      // Univer lists the bands from the highest down (iconId '0' = best icon); Excel from the lowest up.
      const list = x.config ?? [], n = list.length
      const cfvo = [{ type: 'percent', value: 0 }]
      for (let k = n - 2; k >= 0; k--) cfvo.push({ ...cfvoOut(list[k].value), ...(list[k].operator === 'greaterThan' ? { gte: false } : {}) })
      const reverse = n > 0 && String(list[0].iconId) !== '0'
      rule = { type: 'iconSet', iconSet: list[0]?.iconType ?? '3TrafficLights1', showValue: x.isShowValue !== false, reverse, cfvo }
    }
    if (!rule) continue
    ws.addConditionalFormatting({ ref, rules: [{ ...rule, priority: priority++, ...(r.stopIfTrue ? { stopIfTrue: true } : {}) }] })
  }
}

let cfSeq = 0
/** ExcelJS ws.conditionalFormattings -> Univer rules for one sheet. */
/** color: the workbook's colour resolver (theme, tint, indexed and ARGB colours). */
export function cfIn(ws, sheetId, color = fromArgb) {
  const out = []
  const list = (ws.conditionalFormattings ?? []).flatMap((cf) => (cf.rules ?? []).map((rule) => ({ ref: cf.ref, rule })))
  list.sort((a, b) => (a.rule.priority ?? 0) - (b.rule.priority ?? 0))
  for (const { ref, rule: x } of list) {
    const ranges = sqrefIn(ref).map((g) => ({ ...g, rangeType: 0, sheetId }))
    if (!ranges.length) continue
    const style = dxfIn(x.style, color)
    let rule = null
    switch (x.type) {
      case 'cellIs': {
        const f = (x.formulae ?? []).map((v) => String(v))
        const text = f.length === 1 && /^".*"$/.test(f[0])
        if (text && ['equal', 'notEqual'].includes(x.operator)) rule = { type: 'highlightCell', subType: 'text', operator: x.operator, value: f[0].slice(1, -1).replace(/""/g, '"'), style }
        else if (['between', 'notBetween'].includes(x.operator) && f.length === 2 && f.every((v) => !isNaN(Number(v)))) rule = { type: 'highlightCell', subType: 'number', operator: x.operator, value: [Number(f[0]), Number(f[1])], style }
        else if (['between', 'notBetween'].includes(x.operator)) {
          // Bounds that are references or formulas: the same test as a formula rule relative to the first cell.
          const g = ranges[0], cell = `${colName(g.startColumn)}${g.startRow + 1}`
          const test = `AND(${cell}>=MIN(${f[0]},${f[1]}),${cell}<=MAX(${f[0]},${f[1]}))`
          rule = { type: 'highlightCell', subType: 'formula', value: `=${x.operator === 'between' ? test : `NOT(${test})`}`, style }
        }
        else if (f.every((v) => !isNaN(Number(v)))) rule = { type: 'highlightCell', subType: 'number', operator: x.operator, value: Number(f[0]), style }
        else {
          // cellIs against a reference or formula: Univer's equivalent is a formula rule relative to the first cell.
          const g = ranges[0], cell = `${colName(g.startColumn)}${g.startRow + 1}`
          const op = { equal: '=', notEqual: '<>', greaterThan: '>', greaterThanOrEqual: '>=', lessThan: '<', lessThanOrEqual: '<=' }[x.operator] ?? '='
          rule = { type: 'highlightCell', subType: 'formula', value: `=${cell}${op}${f[0]}`, style }
        }
        break
      }
      case 'containsText': rule = { type: 'highlightCell', subType: 'text', operator: x.operator ?? 'containsText', value: x.text ?? textFromFormula(x.formulae?.[0]) ?? '', style }; break
      case 'timePeriod': rule = { type: 'highlightCell', subType: 'timePeriod', operator: x.timePeriod, style }; break
      case 'top10': rule = { type: 'highlightCell', subType: 'rank', isBottom: !!x.bottom, isPercent: !!x.percent, value: x.rank ?? 10, style }; break
      case 'aboveAverage': {
        const operator = `${x.aboveAverage === false ? 'lessThan' : 'greaterThan'}${x.equalAverage ? 'OrEqual' : ''}`
        if (!x.stdDev) { rule = { type: 'highlightCell', subType: 'average', operator, style }; break }
        // N standard deviations above / below the average: the same test as a formula (Univer has no such rule),
        // marked so it is written back as Excel's rule.
        const g = ranges[0], cell = `${colName(g.startColumn)}${g.startRow + 1}`, all = ranges.map((r) => absA1(r)).join(',')
        const above = x.aboveAverage !== false, op = above ? (x.equalAverage ? '>=' : '>') : (x.equalAverage ? '<=' : '<')
        rule = { type: 'highlightCell', subType: 'formula', value: `=${cell}${op}AVERAGE(${all})${above ? '+' : '-'}${x.stdDev}*STDEV.P(${all})`, style,
          sheetlyStdDev: { n: x.stdDev, above, equal: !!x.equalAverage } }
        break
      }
      case 'expression': rule = { type: 'highlightCell', subType: 'formula', value: `=${x.formulae?.[0] ?? 'FALSE'}`, style }; break
      case 'duplicateValues': case 'uniqueValues': rule = { type: 'highlightCell', subType: x.type, style }; break
      case 'dataBar': rule = { type: 'dataBar', isShowValue: x.showValue !== false,
        config: { min: cfvoIn(x.cfvo?.[0]), max: cfvoIn(x.cfvo?.[1]), isGradient: x.gradient !== false, positiveColor: color(x.color) ?? '#638ec6', nativeColor: '#ff0000' } }; break
      case 'colorScale': rule = { type: 'colorScale', config: (x.cfvo ?? []).map((v, i) => ({ index: i, color: color(x.color?.[i]) ?? '#000000', value: cfvoIn(v) })) }; break
      case 'iconSet': {
        const cfvo = x.cfvo ?? [], n = cfvo.length
        const config = []
        for (let k = 0; k < n; k++) {
          // band k from the top uses Excel threshold n-1-k; the last band is "below the next one up".
          const t = n - 1 - k
          const id = String(x.reverse ? n - 1 - k : k)
          if (k < n - 1) config.push({ operator: cfvo[t].gte === false ? 'greaterThan' : 'greaterThanOrEqual', value: cfvoIn(cfvo[t]), iconType: x.iconSet ?? '3TrafficLights1', iconId: id })
          else config.push({ operator: cfvo[1]?.gte === false ? 'lessThanOrEqual' : 'lessThan', value: cfvoIn(cfvo[1] ?? cfvo[0]), iconType: x.iconSet ?? '3TrafficLights1', iconId: id })
        }
        rule = { type: 'iconSet', isShowValue: x.showValue !== false, config }
        break
      }
      default: break
    }
    if (rule) out.push({ cfId: `cf${++cfSeq}`, ranges, stopIfTrue: !!x.stopIfTrue, rule })
  }
  return out
}
const textFromFormula = (f) => (typeof f === 'string' ? /"((?:[^"]|"")*)"/.exec(f)?.[1]?.replace(/""/g, '"') : undefined)

// ---- data validation
const ERROR_STYLE_OUT = { 0: 'information', 1: 'stop', 2: 'warning' }, ERROR_STYLE_IN = { information: 0, stop: 1, warning: 2 }
/** Univer list formula1 (JSON array, comma list, or =range) -> Excel formula1. */
function listOut(f) {
  const s = String(f ?? '')
  if (s.startsWith('=')) return s.slice(1)
  let items
  try { const j = JSON.parse(s); items = Array.isArray(j) ? j : null } catch { items = null }
  items ??= s.split(',')
  // a quote inside an item is doubled in the literal (12" -> 12"")
  return `"${items.map((x) => String(x).replace(/"/g, '""')).join(',')}"`
}
export function dvOut(ws, rules = []) {
  for (const r of rules) {
    const type = { list: 'list', listMultiple: 'list', whole: 'whole', decimal: 'decimal', date: 'date', time: 'time', textLength: 'textLength', custom: 'custom' }[r.type]
    if (!type || !r.ranges?.length) continue
    const f = (v) => (v == null || v === '' ? undefined : typeof v === 'number' ? v : String(v).replace(/^=/, ''))
    const formulae = type === 'list' ? [listOut(r.formula1)] : [f(r.formula1), f(r.formula2)].filter((x) => x !== undefined)
      .map((x) => (typeof x === 'string' && !isNaN(Number(x)) && x.trim() !== '' ? Number(x) : x))
    ws.dataValidations.model[sqrefOut(r.ranges)] = {
      type, formulae, allowBlank: r.allowBlank !== false,
      ...(type !== 'list' && type !== 'custom' ? { operator: r.operator ?? 'between' } : {}),
      showErrorMessage: r.showErrorMessage !== false, ...(r.errorStyle != null ? { errorStyle: ERROR_STYLE_OUT[r.errorStyle] ?? 'stop' } : {}),
      ...(r.error ? { error: r.error } : {}), ...(r.errorTitle ? { errorTitle: r.errorTitle } : {}),
      // the prompt text is kept even while showing it is switched off
      ...(r.showInputMessage ? { showInputMessage: true } : {}), ...(r.prompt ? { prompt: r.prompt } : {}), ...(r.promptTitle ? { promptTitle: r.promptTitle } : {}),
    }
  }
}
let dvSeq = 0
export function dvIn(ws, sheetId) {
  const out = []
  for (const [sqref, v] of Object.entries(ws.dataValidations?.model ?? {})) {
    if (!v || v.type === 'any') continue
    const ranges = sqrefIn(sqref).map((g) => ({ ...g, rangeType: 0, sheetId }))
    if (!ranges.length) continue
    const raw = (x) => (x == null ? undefined : x instanceof Date ? String(Math.round((x.getTime() - Date.UTC(1899, 11, 30)) / 864e5)) : typeof x === 'number' ? String(x) : String(x))
    let formula1 = raw(v.formulae?.[0]), formula2 = raw(v.formulae?.[1])
    // a literal list keeps its items exactly (doubled quotes become one); items with a comma or quote go as JSON
    if (v.type === 'list') {
      if (formula1 && /^".*"$/s.test(formula1)) {
        const items = formula1.slice(1, -1).replace(/""/g, '"').split(',')
        formula1 = items.some((x) => x.includes('"')) ? JSON.stringify(items) : items.join(',')
      } else formula1 = formula1 ? `=${formula1}` : ''
    }
    else {
      const ref = (x) => (x === undefined || !isNaN(Number(x)) ? x : `=${x}`)
      formula1 = ref(formula1); formula2 = ref(formula2)
    }
    out.push({ uid: `dv${++dvSeq}`, ranges, type: v.type, ...(v.operator && v.type !== 'list' && v.type !== 'custom' ? { operator: v.operator } : {}),
      // In the file an omitted allowBlank / showErrorMessage means off.
      formula1, ...(formula2 !== undefined ? { formula2 } : {}), allowBlank: v.allowBlank === true, showDropDown: true,
      showErrorMessage: v.showErrorMessage === true, errorStyle: ERROR_STYLE_IN[v.errorStyle ?? 'stop'] ?? 1,
      ...(v.error ? { error: v.error } : {}), ...(v.errorTitle ? { errorTitle: v.errorTitle } : {}),
      ...(v.showInputMessage ? { showInputMessage: true } : {}), ...(v.prompt ? { prompt: v.prompt } : {}), ...(v.promptTitle ? { promptTitle: v.promptTitle } : {}) })
  }
  return out
}

// ---- tables
// The text a cell shows: plain values, rich text, a link's text, a formula's result.
const cellText = (v) => {
  if (v == null) return ''
  if (typeof v !== 'object') return String(v)
  if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('')
  if ('hyperlink' in v) return cellText(v.text)
  if ('formula' in v || 'sharedFormula' in v) return cellText(v.result)
  if ('error' in v) return v.error
  return String(v.text ?? '')
}
const tableName = (s, used) => {
  // Excel allows letters of any script, digits, _ and . (first: a letter or _); only invalid characters change.
  let n = String(s || 'Table').replace(/[^\p{L}\p{N}_.]/gu, '_').replace(/^([^\p{L}_])/u, '_$1')
  const base = n
  for (let i = 2; used.has(n.toLowerCase()); i++) n = `${base}${i}`
  used.add(n.toLowerCase())
  return n
}
/** Univer tables of one sheet -> ExcelJS tables (call after the cells are written; values are kept). */
export function tablesOut(ws, tables = [], used = new Set()) {
  for (const t of tables) {
    const g = t.range
    if (!g) continue
    const header = t.options?.showHeader !== false, footer = !!t.options?.showFooter
    const first = header ? g.startRow + 1 : g.startRow, last = footer ? g.endRow - 1 : g.endRow
    const columns = []
    for (let c = g.startColumn; c <= g.endColumn; c++) {
      const col = t.columns?.[c - g.startColumn]
      const name = header ? (cellText(ws.getCell(g.startRow + 1, c + 1).value) || col?.displayName || `Column${c - g.startColumn + 1}`) : col?.displayName ?? `Column${c - g.startColumn + 1}`
      columns.push({ name, filterButton: t.options?.showAutoFilter !== false })
    }
    const rows = []
    for (let r = first; r <= last; r++) {
      const row = []
      for (let c = g.startColumn; c <= g.endColumn; c++) row.push(ws.getCell(r + 1, c + 1).value ?? null)
      rows.push(row)
    }
    // (a table with only its header keeps its range: no row is added below it)
    const styles = []
    for (let r = g.startRow; r <= g.endRow; r++) for (let c = g.startColumn; c <= g.endColumn; c++) styles.push([r, c, ws.getCell(r + 1, c + 1).style])
    // The totals row as it is (labels, SUBTOTAL formulas and their results): ExcelJS would write its defaults.
    // The header row as it is too (a header may be a link or rich text; ExcelJS would write the plain name).
    const headerCells = header ? Array.from({ length: g.endColumn - g.startColumn + 1 }, (_, i) => [g.startRow, g.startColumn + i, ws.getCell(g.startRow + 1, g.startColumn + i + 1).value]) : []
    const footerCells = footer ? Array.from({ length: g.endColumn - g.startColumn + 1 }, (_, i) => [g.endRow, g.startColumn + i, ws.getCell(g.endRow + 1, g.startColumn + i + 1).value]) : []
    footerCells.forEach(([, , v], i) => {
      if (typeof v === 'string') columns[i].totalsRowLabel = v
      else if (v && typeof v === 'object' && v.formula) { columns[i].totalsRowFunction = 'custom'; columns[i].totalsRowFormula = v.formula; columns[i].totalsRowResult = v.result }
    })
    ws.addTable({ name: tableName(t.name, used), ref: `${colName(g.startColumn)}${g.startRow + 1}`, headerRow: header, totalsRow: footer,
      style: { theme: 'TableStyleMedium2', showRowStripes: true }, columns, rows })
    // addTable rewrites the cells; their own formatting and the totals row are put back.
    for (const [r, c, st] of styles) ws.getCell(r + 1, c + 1).style = st
    for (const [r, c, v] of [...headerCells, ...footerCells]) ws.getCell(r + 1, c + 1).value = v
  }
}
/** ExcelJS tables -> [{name, range, showHeader, showFooter}] for FWorksheet.addTable after the workbook loads. */
export function tablesIn(ws) {
  const out = []
  for (const t of Object.values(ws.tables ?? {})) {
    const m = t.table ?? t.model ?? t
    const g = parseRange(m.tableRef ?? m.ref)
    if (!g) continue
    out.push({ name: m.displayName ?? m.name, range: g, showHeader: m.headerRow !== false, showFooter: !!m.totalsRow, columns: (m.columns ?? []).map((c) => c.name) })
  }
  return out
}

// ---- drawing anchors (pictures and charts)
/** Pixel geometry of a sheet snapshot, for converting cell anchors to positions. */
export function sheetGeometry(sheet) {
  const cw = (c) => (sheet.columnData?.[c]?.hd ? 0 : sheet.columnData?.[c]?.w ?? sheet.defaultColumnWidth ?? 64)
  const rh = (r) => (sheet.rowData?.[r]?.hd ? 0 : sheet.rowData?.[r]?.h ?? sheet.defaultRowHeight ?? 20)
  const left = (c) => { let x = 0; for (let i = 0; i < c; i++) x += cw(i); return x }
  const top = (r) => { let y = 0; for (let i = 0; i < r; i++) y += rh(i); return y }
  return { cw, rh, left, top }
}
const anchorIn = (a) => ({ column: a?.nativeCol ?? Math.floor(a?.col ?? 0), columnOffset: Math.round((a?.nativeColOff ?? 0) / EMU_PER_PX), row: a?.nativeRow ?? Math.floor(a?.row ?? 0), rowOffset: Math.round((a?.nativeRowOff ?? 0) / EMU_PER_PX) })
export const anchorOut = (p) => ({ nativeCol: p.column, nativeColOff: Math.max(0, Math.round((p.columnOffset ?? 0) * EMU_PER_PX)), nativeRow: p.row, nativeRowOff: Math.max(0, Math.round((p.rowOffset ?? 0) * EMU_PER_PX)) })
/** The cell anchor width x height px away from a start anchor. */
export function endFrom(sheet, from, width, height) {
  const geo = sheetGeometry(sheet)
  let c = from.column, x = from.columnOffset + width
  while (x > geo.cw(c) && c < 16383) { x -= geo.cw(c); c++ }
  let r = from.row, y = from.rowOffset + height
  while (y > geo.rh(r) && r < 1048575) { y -= geo.rh(r); r++ }
  return { column: c, columnOffset: Math.round(x), row: r, rowOffset: Math.round(y) }
}
/** A Univer drawing placed by a from/to cell anchor. */
export function placed(sheet, from, to, extra) {
  const geo = sheetGeometry(sheet)
  const x0 = geo.left(from.column) + from.columnOffset, y0 = geo.top(from.row) + from.rowOffset
  const x1 = geo.left(to.column) + to.columnOffset, y1 = geo.top(to.row) + to.rowOffset
  const sheetTransform = { from, to }
  return { ...extra, sheetTransform, axisAlignSheetTransform: structuredClone(sheetTransform),
    transform: { left: x0, top: y0, width: Math.max(8, x1 - x0), height: Math.max(8, y1 - y0) } }
}

/** Pictures of one sheet (Univer drawings of type 0 with a data URL) -> ExcelJS images. */
export function imagesOut(wb, ws, drawings = []) {
  for (const d of drawings) {
    if (d.drawingType !== 0 || typeof d.source !== 'string') continue
    const m = /^data:image\/(png|jpe?g|gif);base64,(.+)$/i.exec(d.source)
    if (!m || !d.sheetTransform) continue
    const id = wb.addImage({ base64: m[2], extension: m[1].toLowerCase() === 'jpg' ? 'jpeg' : m[1].toLowerCase() })
    ws.addImage(id, { tl: anchorOut(d.sheetTransform.from), br: anchorOut(d.sheetTransform.to), editAs: 'oneCell' })
  }
}
let imgSeq = 0
// Base64 in Node (Buffer) and in the app (btoa over chunks, so large pictures do not overflow the call stack).
function toBase64(buf) {
  if (typeof Buffer !== 'undefined') return Buffer.from(buf).toString('base64')
  const bytes = new Uint8Array(buf.buffer ?? buf)
  let s = ''
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  return btoa(s)
}
/** ExcelJS images of one sheet -> Univer picture drawings. */
export function imagesIn(wb, ws, sheet, unitId) {
  const out = []
  for (const im of ws.getImages?.() ?? []) {
    const media = wb.getImage(im.imageId)
    if (!media?.buffer || !im.range?.tl) continue
    const ext = media.extension === 'jpeg' ? 'jpeg' : media.extension
    const b64 = toBase64(media.buffer)
    const from = anchorIn(im.range.tl)
    // One-cell anchors store the picture's size (ext, px) instead of an end cell.
    const to = im.range.br ? anchorIn(im.range.br) : endFrom(sheet, from, im.range.ext?.width ?? 192, im.range.ext?.height ?? 100)
    const drawingId = `img${++imgSeq}_${Date.now().toString(36)}`
    out.push(placed(sheet, from, to, { unitId, subUnitId: sheet.id, drawingId, drawingType: 0, imageSourceType: 'BASE64', source: `data:image/${ext};base64,${b64}`, anchorType: '0' }))
  }
  return out
}
export { plainRange }
