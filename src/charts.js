// Charts for F11 (chart sheet) and Alt+F1 (embedded chart): Excel's default clustered column chart
// of the current region, drawn with Chart.js in Univer's floating-DOM layer.
import { BarController, BarElement, CategoryScale, Chart, Legend, LinearScale, Title, Tooltip } from 'chart.js'
import { createElement, useEffect, useRef } from 'react'

Chart.register(BarController, BarElement, CategoryScale, LinearScale, Legend, Title, Tooltip)
// Office theme accent colors, in Excel's series order.
const SERIES_COLORS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
const COMPONENT = 'sheetly-chart'

function ChartView({ data }) {
  const ref = useRef(null)
  useEffect(() => {
    if (!ref.current || !data) return undefined
    const chart = new Chart(ref.current, {
      type: 'bar',
      data: { labels: data.labels, datasets: data.series.map((s, i) => ({ label: s.name, data: s.values, backgroundColor: SERIES_COLORS[i % SERIES_COLORS.length] })) },
      options: { responsive: true, maintainAspectRatio: false, animation: false,
        plugins: { title: { display: true, text: data.title }, legend: { position: 'bottom', display: data.series.length > 1 || !!data.series[0]?.name } } },
    })
    return () => chart.destroy()
  }, [data])
  return createElement('div', { className: 'sx-chart', role: 'img', 'aria-label': data?.title || 'Chart',
    style: { width: '100%', height: '100%', background: '#fff', border: '1px solid #d9d9d9', boxSizing: 'border-box', padding: '6px' } },
  createElement('canvas', { ref }))
}

export function createCharts(api, dialogs) {
  api.registerComponent(COMPONENT, ChartView)
  const wb = () => api.getActiveWorkbook()
  const ws = () => wb().getActiveSheet()

  /** Builds Excel-style series from the selection (or the current region of a single cell). */
  function chartData() {
    const g0 = ws().getActiveRange().getRange()
    const g = g0.startRow === g0.endRow && g0.startColumn === g0.endColumn ? dialogs.currentRegion(g0.startRow, g0.startColumn) : g0
    const v = (r, c) => ws().getRange(r, c).getValue()
    const isNum = (x) => typeof x === 'number'
    const headerRow = [...Array(g.endColumn - g.startColumn + 1)].some((_, i) => i > 0 && !isNum(v(g.startRow, g.startColumn + i)) && v(g.startRow, g.startColumn + i) != null)
    const labelCol = [...Array(g.endRow - g.startRow + 1)].some((_, i) => i > 0 && !isNum(v(g.startRow + i, g.startColumn)) && v(g.startRow + i, g.startColumn) != null)
    const r0 = g.startRow + (headerRow ? 1 : 0), c0 = g.startColumn + (labelCol ? 1 : 0)
    const rows = g.endRow - r0 + 1, cols = g.endColumn - c0 + 1
    // Excel plots series by columns when there are at least as many rows as columns.
    const byCols = rows >= cols
    const labels = [], series = []
    if (byCols) {
      for (let r = r0; r <= g.endRow; r++) labels.push(labelCol ? String(v(r, g.startColumn) ?? '') : String(r - r0 + 1))
      for (let c = c0; c <= g.endColumn; c++) series.push({ name: headerRow ? String(v(g.startRow, c) ?? '') : `Series${c - c0 + 1}`, values: labels.map((_, i) => num(v(r0 + i, c))) })
    } else {
      for (let c = c0; c <= g.endColumn; c++) labels.push(headerRow ? String(v(g.startRow, c) ?? '') : String(c - c0 + 1))
      for (let r = r0; r <= g.endRow; r++) series.push({ name: labelCol ? String(v(r, g.startColumn) ?? '') : `Series${r - r0 + 1}`, values: labels.map((_, i) => num(v(r, c0 + i))) })
    }
    const title = series.length === 1 ? series[0].name : 'Chart Title'
    return { data: { labels, series, title }, range: g }
  }
  const num = (x) => (typeof x === 'number' ? x : Number(x) || 0)
  const colX = (s, c) => { let x = 0; for (let i = 0; i < c; i++) x += s.getColumnWidth(i); return x }
  const rowY = (s, r) => { let y = 0; for (let i = 0; i < r; i++) y += s.getRowHeight(i); return y }

  function embeddedChart() {
    const { data, range } = chartData()
    const s = ws()
    const x = colX(s, range.endColumn + 2), y = rowY(s, range.startRow)
    s.addFloatDomToPosition({ componentKey: COMPONENT, allowTransform: true, data, initPosition: { startX: x, endX: x + 480, startY: y, endY: y + 288 } })
    return true
  }
  function chartSheet() {
    const { data } = chartData()
    const w = wb(), names = new Set(w.getSheets().map((s) => s.getSheetName()))
    let n = 1
    while (names.has(`Chart${n}`)) n++
    const index = w.getSheets().findIndex((s) => s.getSheetId() === ws().getSheetId())
    const sheet = w.insertSheet(`Chart${n}`, { index })
    w.setActiveSheet(sheet)
    sheet.setHiddenGridlines?.(true)
    setTimeout(() => sheet.addFloatDomToPosition({ componentKey: COMPONENT, allowTransform: false, data,
      initPosition: { startX: 0, endX: Math.max(640, window.innerWidth - 80), startY: 0, endY: Math.max(400, window.innerHeight - 200) } }), 0)
    return true
  }
  return { chartData, embeddedChart, chartSheet }
}
