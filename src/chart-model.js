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
