// Outline model shared by the outline UI and .xlsx conversion (no DOM / Univer dependencies).
export const MAX_LEVEL = 7

/** Groups of a sparse level map: maximal runs of consecutive indexes with level >= k, for every k. */
export function outlineGroups(levels) {
  const idx = [...levels.keys()].sort((a, b) => a - b)
  const out = [], open = new Array(MAX_LEVEL + 1).fill(null)
  const closeFrom = (k) => { for (let j = k; j <= MAX_LEVEL; j++) if (open[j]) { out.push(open[j]); open[j] = null } }
  for (const i of idx) {
    const L = levels.get(i)
    for (let k = 1; k <= MAX_LEVEL; k++) {
      if (k > L) { closeFrom(k); break }
      if (open[k] && open[k].end === i - 1) open[k].end = i
      else { if (open[k]) out.push(open[k]); open[k] = { level: k, start: i, end: i } }
    }
  }
  closeFrom(1)
  return out.sort((a, b) => a.level - b.level || a.start - b.start)
}

/** Per-sheet outline settings (worksheet custom.sheetlyOutline): Excel's defaults are summary rows below the
 *  detail, summary columns to the right, outline symbols shown. */
export function outlineSettings(v) {
  const summaryBelow = v?.summaryBelow !== false, summaryRight = v?.summaryRight !== false, hideSymbols = v?.hideSymbols === true
  const autoStyles = v?.autoStyles === true
  return { summaryBelow, summaryRight, hideSymbols, autoStyles, isDefault: summaryBelow && summaryRight && !hideSymbols && !autoStyles,
    value: { summaryBelow, summaryRight, hideSymbols, autoStyles } }
}

/** Index of a group's summary row/column: after the group (below / right) or before it (above / left). */
export const summaryIndex = (g, after) => (after ? g.end + 1 : g.start - 1)
