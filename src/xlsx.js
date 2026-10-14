// .xlsx <-> Univer snapshot via ExcelJS (values, formulas, basic styles, widths, merges).
import ExcelJS from 'exceljs'

const colorToArgb = (c) => (c ? 'FF' + c.replace('#', '').toUpperCase() : undefined)

export async function xlsxToSnapshot(buf) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf)
  const sheets = {}, order = []
  wb.eachSheet((ws, i) => {
    const id = `sheet${i}`, cellData = {}, columnData = {}, mergeData = []
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
    ws.columns?.forEach((col, i2) => { if (col.width) columnData[i2] = { w: Math.round(col.width * 7) } })
    Object.keys(ws._merges || {}).forEach((k) => {
      const m = ws._merges[k].model
      mergeData.push({ startRow: m.top - 1, endRow: m.bottom - 1, startColumn: m.left - 1, endColumn: m.right - 1 })
    })
    sheets[id] = { id, name: ws.name, cellData, columnData, mergeData,
      rowCount: Math.max(1000, ws.rowCount + 100), columnCount: Math.max(26, ws.columnCount + 10) }
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
    for (const m of s.mergeData || []) ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1)
  }
  return new Uint8Array(await wb.xlsx.writeBuffer())
}
