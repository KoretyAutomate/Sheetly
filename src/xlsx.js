// .xlsx <-> Univer snapshot via ExcelJS (values, formulas, basic styles, widths, merges, hidden rows/columns and
// outline levels with Excel's "collapsed" flag on summary rows/columns; ExcelJS is patched to carry it). On load a
// summary flagged collapsed, or a group whose detail is all hidden, sets custom.outlineCollapsed for that group.
import ExcelJS from 'exceljs'
import { outlineGroups } from './outline-model.js'

const colorToArgb = (c) => (c ? 'FF' + c.replace('#', '').toUpperCase() : undefined)

export async function xlsxToSnapshot(buf) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf)
  const sheets = {}, order = []
  wb.eachSheet((ws, i) => {
    const id = `sheet${i}`, cellData = {}, columnData = {}, rowData = {}, mergeData = []
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      cellData[r - 1] = {}
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        const v = cell.value, o = {}
        if (v && typeof v === 'object' && 'formula' in v) { o.f = '=' + v.formula; if (v.result !== undefined) o.v = v.result }
        else if (v && typeof v === 'object' && 'richText' in v) o.v = v.richText.map((t) => t.text).join('')
        else if (v instanceof Date) o.v = v.toISOString()
        else o.v = v
        if (typeof o.v === 'number') o.t = 2; else if (typeof o.v === 'boolean') { o.t = 3; o.v = o.v ? 1 : 0 }
        const s = {}
        if (cell.font?.bold) s.bl = 1
        if (cell.font?.italic) s.it = 1
        if (cell.numFmt && cell.numFmt !== 'General') s.n = { pattern: cell.numFmt }
        if (Object.keys(s).length) o.s = s
        cellData[r - 1][c - 1] = o
      })
    })
    ws.columns?.forEach((col, i2) => {
      const d = {}
      if (col.width) d.w = Math.round(col.width * 7)
      if (col.hidden) d.hd = 1
      if (col.outlineLevel) d.custom = { outlineLevel: col.outlineLevel }
      if (col.collapsed) d.collapsed = true
      if (Object.keys(d).length) columnData[i2] = d
    })
    ws.eachRow({ includeEmpty: true }, (row, r) => {
      const d = {}
      if (row.hidden) d.hd = 1
      if (row.outlineLevel) d.custom = { outlineLevel: row.outlineLevel }
      if (row.collapsed) d.collapsed = true
      if (Object.keys(d).length) rowData[r - 1] = d
    })
    markCollapsed(rowData)
    markCollapsed(columnData)
    Object.keys(ws._merges || {}).forEach((k) => {
      const m = ws._merges[k].model
      mergeData.push({ startRow: m.top - 1, endRow: m.bottom - 1, startColumn: m.left - 1, endColumn: m.right - 1 })
    })
    sheets[id] = { id, name: ws.name, cellData, columnData, rowData, mergeData,
      rowCount: 1048576, columnCount: 16384 }
    order.push(id)
  })
  return { id: 'wb', sheetOrder: order, sheets }
}

export async function snapshotToXlsx(snap) {
  const wb = new ExcelJS.Workbook()
  for (const id of snap.sheetOrder) {
    const s = snap.sheets[id], ws = wb.addWorksheet(s.name)
    for (const [r, row] of Object.entries(s.cellData || {})) for (const [c, cell] of Object.entries(row)) {
      const x = ws.getCell(+r + 1, +c + 1)
      if (cell.f) x.value = { formula: cell.f.replace(/^=/, ''), result: cell.v }
      else if (cell.v !== undefined && cell.v !== null) x.value = cell.t === 3 ? !!cell.v : cell.v
      const st = cell.s && typeof cell.s === 'object' ? cell.s : null
      if (st?.bl || st?.it) x.font = { bold: !!st.bl, italic: !!st.it }
      if (st?.n?.pattern) x.numFmt = st.n.pattern
    }
    for (const [c, col] of Object.entries(s.columnData || {})) if (col.w) ws.getColumn(+c + 1).width = col.w / 7
    let maxRow = 0, maxCol = 0
    for (const [r, row] of Object.entries(s.rowData || {})) {
      const L = row?.custom?.outlineLevel ?? 0
      if (L) { ws.getRow(+r + 1).outlineLevel = L; maxRow = Math.max(maxRow, L) }
      if (L || row?.custom?.outlineCollapsed) ws.getRow(+r + 1).collapsed = !!row?.custom?.outlineCollapsed
      if (row?.hd) ws.getRow(+r + 1).hidden = true
    }
    for (const [c, col] of Object.entries(s.columnData || {})) {
      const L = col?.custom?.outlineLevel ?? 0
      if (L) { ws.getColumn(+c + 1).outlineLevel = L; maxCol = Math.max(maxCol, L) }
      if (L || col?.custom?.outlineCollapsed) ws.getColumn(+c + 1).collapsed = !!col?.custom?.outlineCollapsed
      if (col?.hd) ws.getColumn(+c + 1).hidden = true
    }
    if (maxRow) ws.properties.outlineLevelRow = maxRow
    if (maxCol) ws.properties.outlineLevelCol = maxCol
    for (const m of s.mergeData || []) ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1)
  }
  return new Uint8Array(await wb.xlsx.writeBuffer())
}

// A group is collapsed when its summary row/column carries Excel's collapsed flag (the group one level deeper than
// the summary itself) or when its detail is all hidden: set that level's bit on the summary after it.
function markCollapsed(data) {
  const levels = new Map()
  for (const [i, d] of Object.entries(data)) if (d.custom?.outlineLevel) levels.set(+i, d.custom.outlineLevel)
  for (const g of outlineGroups(levels)) {
    const summary = data[g.end + 1]
    let hidden = true
    for (let i = g.start; i <= g.end && hidden; i++) hidden = data[i]?.hd === 1
    const flagged = summary?.collapsed && g.level === (summary.custom?.outlineLevel ?? 0) + 1
    if (!hidden && !flagged) continue
    const s = (data[g.end + 1] ??= {})
    s.custom = { ...(s.custom ?? {}), outlineCollapsed: (s.custom?.outlineCollapsed ?? 0) | (1 << g.level) }
  }
  for (const d of Object.values(data)) delete d.collapsed
}
