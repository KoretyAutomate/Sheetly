// How Excel reads a chart's source range (shared by the chart view and the .xlsx chart writer; no DOM).
// A header row when a later column's first cell is text; a label column when a later row's first cell is text (or
// the top-left cell is empty above a header row); series by columns when there are at least as many rows as columns.
// Empty cells are gaps (null), not zeros, as with Excel's default "show empty cells as gaps".
const isNum = (x) => typeof x === 'number'
const filled = (x) => x != null && x !== ''

export function seriesFrom(get, g, plotBy, { hiddenRow = () => false, hiddenCol = () => false } = {}) {
  const v = (r, c) => get(r, c)
  const headerRow = [...Array(g.endColumn - g.startColumn + 1)].some((_, i) => i > 0 && filled(v(g.startRow, g.startColumn + i)) && !isNum(v(g.startRow, g.startColumn + i)))
  const firstData = g.startRow + (headerRow ? 1 : 0)
  // Leading text columns are category labels (several make a multi-level axis, e.g. Region then Product).
  const textCol = (c) => { let any = false; for (let r = firstData; r <= g.endRow; r++) { const x = v(r, c); if (isNum(x)) return false; if (filled(x)) any = true } return any }
  let labelCount = 0
  while (g.startColumn + labelCount < g.endColumn && textCol(g.startColumn + labelCount)) labelCount++
  if (!labelCount && [...Array(g.endRow - g.startRow + 1)].some((_, i) => i > 0 && filled(v(g.startRow + i, g.startColumn)) && !isNum(v(g.startRow + i, g.startColumn)))) labelCount = 1
  if (!labelCount && headerRow && !filled(v(g.startRow, g.startColumn))) labelCount = 1
  const labelCol = labelCount > 0
  const r0 = firstData, c0 = g.startColumn + labelCount
  // Hidden rows and columns are not plotted (Excel's default).
  const rows = [], cols = []
  for (let r = r0; r <= g.endRow; r++) if (!hiddenRow(r)) rows.push(r)
  for (let c = c0; c <= g.endColumn; c++) if (!hiddenCol(c)) cols.push(c)
  const byCols = plotBy ? plotBy === 'cols' : g.endRow - r0 + 1 >= g.endColumn - c0 + 1
  const val = (r, c) => { const x = v(r, c); return isNum(x) ? x : filled(x) && !isNaN(Number(x)) ? Number(x) : null }
  const labels = [], series = []
  const levels = []
  if (byCols) {
    for (const r of rows) labels.push(labelCol ? String(v(r, c0 - 1) ?? '') : String(r - r0 + 1))
    // Outer label levels (left of the innermost): shown once per group, as Excel's multi-level axis does.
    for (let k = labelCount - 2; k >= 0; k--) { let prev = null; levels.push(rows.map((r) => { const x = String(v(r, g.startColumn + k) ?? ''); const show = filled(v(r, g.startColumn + k)) && x !== prev; if (filled(v(r, g.startColumn + k))) prev = x; return show ? x : '' })) }
    for (const [i, c] of cols.entries()) series.push({ name: headerRow ? String(v(g.startRow, c) ?? '') : `Series${i + 1}`, values: rows.map((r) => val(r, c)), line: c,
      nameCell: headerRow ? [g.startRow, c] : null })
  } else {
    for (const c of cols) labels.push(headerRow ? String(v(g.startRow, c) ?? '') : String(c - c0 + 1))
    for (const [i, r] of rows.entries()) series.push({ name: labelCol ? String(v(r, c0 - 1) ?? '') : `Series${i + 1}`, values: cols.map((c) => val(r, c)), line: r,
      nameCell: labelCol ? [r, c0 - 1] : null })
  }
  return { labels, levels, series, byCols, headerRow, labelCol, labelCount, r0, c0, rows, cols }
}

/** Automatic title: the series name for one series, else "Chart Title"; '' = no title. */
export const chartTitle = (spec, d) => (spec?.title != null ? spec.title : d.series.length === 1 ? d.series[0].name : 'Chart Title')

/**
 * Data of a chart whose series were picked one by one (not one block): source.series = [{ name: [r, c] | null,
 * nameText, values: IRange }], source.categories: IRange | null. Values run along each range; points in hidden rows
 * or columns are left out (Excel's default), together with their categories. A category range that is several
 * columns wide (vertical series) or rows high (horizontal series) is a multi-level axis: the last column / row is
 * the inner level, the others outer levels shown once per group.
 */
export function explicitSeries(get, src, { hiddenRow = () => false, hiddenCol = () => false } = {}) {
  const val = (x) => (isNum(x) ? x : filled(x) && !isNaN(Number(x)) ? Number(x) : null)
  const first = src.series[0]?.values
  const vertical = !first || first.startColumn === first.endColumn
  // Each range is read along its own direction (a chart may mix a column range with a row range).
  const isVertical = (g) => g.startColumn === g.endColumn && (g.startRow !== g.endRow || vertical)
  const along = (g) => { const out = []; if (isVertical(g)) for (let r = g.startRow; r <= g.endRow; r++) out.push([r, g.startColumn]); else for (let c = g.startColumn; c <= g.endColumn; c++) out.push([g.startRow, c]); return out }
  // A hidden row (vertical series) or column (horizontal) drops that point from every series, with its category; a
  // hidden column (vertical) or row (horizontal) drops just the series in it.
  // Each range's own cells decide which of its points are shown (a hidden row hides the points in that row).
  const keepFor = (g) => along(g).map(([r, c]) => (isVertical(g) ? !hiddenRow(r) : !hiddenCol(c)))
  const keep = first ? keepFor(first) : []
  const seriesShown = (g) => (isVertical(g) ? !hiddenCol(g.startColumn) : !hiddenRow(g.startRow))
  const series = src.series.map((s, i) => ({ s, i })).filter(({ s }) => seriesShown(s.values)).map(({ s, i }) => {
    const mask = keepFor(s.values) // once per series, shared by its values and scatter X values
    return { name: s.name ? String(get(s.name[0], s.name[1]) ?? '') : s.nameText ?? `Series${i + 1}`,
      values: along(s.values).filter((_, k) => mask[k] !== false).map(([r, c]) => val(get(r, c))), nameCell: s.name,
      ...(s.xValues ? { xs: along(s.xValues).filter((_, k) => mask[k] !== false).map(([r, c]) => val(get(r, c))) } : {}) }
  })
  const n = Math.max(0, ...series.map((s) => s.values.length))
  let labels = Array.from({ length: n }, (_, i) => String(i + 1)), levels = []
  const cat = src.categories
  if (cat) {
    // inner level: the last column (vertical) or row (horizontal); outer levels before it
    const lines = vertical ? Array.from({ length: cat.endColumn - cat.startColumn + 1 }, (_, i) => cat.startColumn + i) : Array.from({ length: cat.endRow - cat.startRow + 1 }, (_, i) => cat.startRow + i)
    const cellsOf = (line) => (vertical ? along({ startRow: cat.startRow, endRow: cat.endRow, startColumn: line, endColumn: line }) : along({ startRow: line, endRow: line, startColumn: cat.startColumn, endColumn: cat.endColumn }))
    const pick = (line) => { const cells = cellsOf(line), k2 = cells.map(([r, c]) => (vertical ? !hiddenRow(r) : !hiddenCol(c))); return cells.filter((_, k) => k2[k]) }
    labels = pick(lines[lines.length - 1]).map(([r, c]) => String(get(r, c) ?? ''))
    for (let k = lines.length - 2; k >= 0; k--) {
      let prev = null
      levels.push(pick(lines[k]).map(([r, c]) => { const x = get(r, c); if (!filled(x)) return ''; const t = String(x); const show = t !== prev; prev = t; return show ? t : '' }))
    }
  }
  return { labels, levels, series, byCols: vertical, headerRow: false, labelCol: !!cat, labelCount: cat ? (vertical ? cat.endColumn - cat.startColumn + 1 : cat.endRow - cat.startRow + 1) : 0, explicit: true }
}

/**
 * Where a chart's references point now. Ranges are kept as hidden defined names (ref, valuesRef, nameRef,
 * categoriesRef) that the spreadsheet moves with structural edits; resolve(name) gives { sheetId, range } or null
 * (#REF!). Charts without names use their stored ranges. Returns null when the source itself is gone.
 */
export function resolveSource(src, resolve) {
  if (!src) return null
  const at = (name, sheetId, range) => (name ? resolve(name) ?? null : range ? { sheetId, range } : null)
  const main = at(src.ref, src.sheetId, src.range)
  if (!main) return null
  const out = { ...src, sheetId: main.sheetId, range: main.range }
  if (src.series) {
    out.series = src.series.map((x) => {
      const v = at(x.valuesRef, src.sheetId, x.values), n = x.nameRef ? at(x.nameRef) : null
      const xv = x.xValuesRef || x.xValues ? at(x.xValuesRef, src.sheetId, x.xValues) : null
      return { ...x, values: v?.range ?? null, xValues: xv?.range ?? undefined, name: n ? [n.range.startRow, n.range.startColumn] : x.nameRef ? null : x.name }
    }).filter((x) => x.values)
    out.categories = at(src.categoriesRef, src.sheetId, src.categories)?.range ?? null
  }
  return out
}
