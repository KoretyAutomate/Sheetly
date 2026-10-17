// Charts, drawn with Chart.js in Univer's floating-DOM layer. A chart is linked to its source range, as in Excel:
// it is redrawn whenever the sheet changes. The drawing's data holds the chart definition (spec), so it is saved
// with the workbook:
//   { v: 2, type, source: { sheetId, range }, plotBy: 'cols' | 'rows', title: null (automatic) | string,
//     legend: 'bottom' | 'right' | 'top' | 'left' | 'none', dataLabels, axisTitles: { x, y } | null }
// Charts made before the spec existed carry {labels, series, title} and are drawn as they were saved.
// F11 makes a chart sheet, Alt+F1 an embedded chart (Excel's default clustered column); Insert Chart and Change
// Chart Type dialogs; Switch Row/Column and chart elements act on the selected chart.
import { ArcElement, BarController, BarElement, CategoryScale, Chart, DoughnutController, Filler, Legend, LinearScale, LineController,
  LineElement, PieController, PointElement, ScatterController, Title, Tooltip } from 'chart.js'
import { CommandType, ICommandService, numfmt } from '@univerjs/core'
import { IDrawingManagerService } from '@univerjs/drawing'
import { DrawingApplyType, ISheetDrawingService, SetDrawingApplyMutation } from '@univerjs/sheets-drawing'
import { createElement, useEffect, useReducer, useRef } from 'react'
import { el, openDialog, radios, check } from './ui/dialog.js'
import { chartTitle, explicitSeries, resolveSource, seriesFrom } from './chart-model.js'

Chart.register(ArcElement, BarController, BarElement, CategoryScale, DoughnutController, Filler, Legend, LinearScale, LineController,
  LineElement, PieController, PointElement, ScatterController, Title, Tooltip)
// Office theme accents in Excel's series order; series 7-12 use the accents darkened, 13-18 lightened.
const ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
const shade = (hex, f) => '#' + [1, 3, 5].map((i) => Math.round(Math.min(255, Math.max(0, f < 0 ? parseInt(hex.slice(i, i + 2), 16) * (1 + f) : parseInt(hex.slice(i, i + 2), 16) + (255 - parseInt(hex.slice(i, i + 2), 16)) * f))).toString(16).padStart(2, '0')).join('')
export const seriesColor = (i) => { const base = ACCENTS[i % 6], round = Math.floor(i / 6) % 3; return round === 0 ? base : round === 1 ? shade(base, -0.4) : shade(base, 0.4) }
export const COMPONENT = 'sheetly-chart'
// The next frame, or a short timer when the window is in the background (Chromium holds animation frames back there).
const soon = (fn) => { let done = false; const run = () => { if (!done) { done = true; fn() } }; requestAnimationFrame(run); setTimeout(run, 60) }


/** Chart types offered, in Excel's Insert Chart order. */
export const CHART_TYPES = [
  ['column', 'Clustered Column', 'Column'], ['columnStacked', 'Stacked Column', 'Column'], ['columnStacked100', '100% Stacked Column', 'Column'],
  ['line', 'Line', 'Line'], ['lineMarkers', 'Line with Markers', 'Line'],
  ['pie', 'Pie', 'Pie'], ['doughnut', 'Doughnut', 'Pie'],
  ['bar', 'Clustered Bar', 'Bar'], ['barStacked', 'Stacked Bar', 'Bar'], ['barStacked100', '100% Stacked Bar', 'Bar'],
  ['area', 'Area', 'Area'], ['areaStacked', 'Stacked Area', 'Area'],
  ['scatter', 'Scatter', 'X Y (Scatter)'],
]
const typeName = (t) => CHART_TYPES.find((x) => x[0] === t)?.[1] ?? t

let apiRef = null, refsRef = null
const num = (x) => (typeof x === 'number' ? x : Number(x) || 0)

export { seriesFrom }

export const liveSource = (src) => resolveSource(src, (n) => refsRef?.resolve(n) ?? null)

/** Live data of a spec from the workbook (falls back to saved data for old charts). */
function dataOf(spec0) {
  if (!spec0?.source) return { labels: spec0?.labels ?? [], series: spec0?.series ?? [], title: spec0?.title ?? '' }
  const source = liveSource(spec0.source)
  if (!source) return { labels: [], series: [], title: spec0.title ?? '' }
  const spec = { ...spec0, source }
  const sheet = apiRef?.getActiveWorkbook()?.getSheetBySheetId(spec.source.sheetId)
  if (!sheet) return { labels: [], series: [], title: spec.title ?? '' }
  const w = sheet.getSheet()
  if (spec.source.series) {
    const d = explicitSeries((r, c) => sheet.getRange(r, c).getCalculatedValue(), spec.source, { hiddenRow: (r) => w.getRowVisible?.(r) === false, hiddenCol: (c) => w.getColVisible?.(c) === false })
    return { ...d, title: chartTitle(spec, d), pattern: null }
  }
  const d = seriesFrom((r, c) => sheet.getRange(r, c).getCalculatedValue(), spec.source.range, spec.plotBy,
    { hiddenRow: (r) => w.getRowVisible?.(r) === false, hiddenCol: (c) => w.getColVisible?.(c) === false })
  // The value axis shows numbers in the source's number format (sourceLinked, as in Excel).
  const pattern = d.series.length ? sheet.getRange(d.byCols ? d.r0 : d.series[0].line, d.byCols ? d.series[0].line : d.c0).getNumberFormat?.() : null
  return { ...d, title: chartTitle(spec, d), pattern: pattern && pattern !== 'General' ? pattern : null }
}

/** Chart.js configuration for a spec and its data. */
export function chartConfig(spec, d) {
  const type = spec?.type ?? 'column'
  const pie = type === 'pie' || type === 'doughnut', horizontal = type.startsWith('bar'), stacked = /Stacked/.test(type), pct = /100$/.test(type)
  const totals = d.labels.map((_, i) => d.series.reduce((s, x) => s + Math.abs(x.values[i] ?? 0), 0))
  const values = (s) => (pct ? s.values.map((x, i) => (x == null ? null : totals[i] ? (x / totals[i]) * 100 : 0)) : s.values)
  const base = { line: 'line', lineMarkers: 'line', area: 'line', areaStacked: 'line', pie: 'pie', doughnut: 'doughnut', scatter: 'scatter' }[type] ?? 'bar'
  let datasets
  if (pie) {
    const s = d.series[0] ?? { name: '', values: [] }
    datasets = [{ label: s.name, data: s.values, backgroundColor: d.labels.map((_, i) => seriesColor(i)), borderColor: '#fff', borderWidth: 1 }]
  } else if (type === 'scatter') {
    // X values come from the first series (Excel uses the first column as X when there is no label column).
    // Series picked one by one carry their X values as categories.
    const xNum = (l, i) => (l !== '' && Number.isFinite(Number(l)) ? Number(l) : i + 1)
    const xs = d.explicit ? d.labels.map(xNum) : d.series.length > 1 ? d.series[0].values : d.labels.map(xNum)
    const ys = d.explicit || d.series.length <= 1 ? d.series : d.series.slice(1)
    datasets = ys.map((s, i) => ({ label: s.name, data: s.values.map((y, k) => ({ x: s.xs ? s.xs[k] : xs[k], y })), backgroundColor: seriesColor(i), borderColor: seriesColor(i), pointRadius: 4 }))
  } else {
    datasets = d.series.map((s, i) => ({ label: s.name, data: values(s), backgroundColor: seriesColor(i), borderColor: seriesColor(i),
      ...(base === 'line' ? { borderWidth: 2.25, pointRadius: type === 'lineMarkers' ? 4 : 0, fill: type.startsWith('area') ? (stacked && i > 0 ? '-1' : 'origin') : false, tension: 0 } : {}),
      // Excel's default gap width 219% / overlap -27% (clustered) or 100% (stacked), approximated.
      ...(base === 'bar' ? { categoryPercentage: stacked ? 0.6 : 0.62, barPercentage: stacked ? 1 : 0.85 } : {}) }))
  }
  const legend = spec?.legend ?? (pie || d.series.length > 1 ? 'bottom' : 'none')
  const axisTitle = (t) => (t ? { display: true, text: t } : { display: false })
  // Excel starts the value axis at 0 unless the data are all on one side of 0 and their spread is under a sixth of
  // the largest magnitude.
  const all = d.series.flatMap((x) => x.values).filter((x) => x != null)
  const lo = Math.min(...all), hi = Math.max(...all)
  const zero = stacked || !all.length || !(lo > 0 ? hi - lo < hi / 6 : hi < 0 ? hi - lo < -lo / 6 : false)
  const fmt = (x) => { try { return d.pattern ? numfmt.format(d.pattern, x) : x } catch { return x } }
  const valueAxis = { stacked, beginAtZero: zero, grid: { color: '#d9d9d9' }, border: { display: false },
    ticks: { callback: pct ? (x) => `${x}%` : (x) => fmt(x) }, ...(pct ? { max: 100 } : {}) }
  const catAxis = { stacked, grid: { display: false } }
  const scales = pie ? {} : type === 'scatter'
    ? { x: { type: 'linear', grid: { color: '#d9d9d9' }, title: axisTitle(spec?.axisTitles?.x) }, y: { ...valueAxis, title: axisTitle(spec?.axisTitles?.y) } }
    : horizontal ? { x: { ...valueAxis, title: axisTitle(spec?.axisTitles?.y) }, y: { ...catAxis, title: axisTitle(spec?.axisTitles?.x) } }
      : { x: { ...catAxis, title: axisTitle(spec?.axisTitles?.x) }, y: { ...valueAxis, title: axisTitle(spec?.axisTitles?.y) } }
  return {
    type: base,
    // Multi-level categories: the inner label with the outer ones under it (shown once per group).
    data: { labels: type === 'scatter' ? undefined : d.levels?.length ? d.labels.map((l, i) => [l, ...d.levels.map((x) => x[i])]) : d.labels, datasets },
    options: { responsive: true, maintainAspectRatio: false, animation: false, indexAxis: horizontal ? 'y' : 'x', scales,
      plugins: { title: { display: spec?.title !== '', text: d.title, font: { size: 14, weight: 'normal' }, color: '#595959' },
        legend: { display: legend !== 'none', position: legend === 'none' ? 'bottom' : legend }, tooltip: { enabled: true } } },
    plugins: spec?.dataLabels ? [dataLabelsPlugin] : [],
  }
}
// Data labels: the value at each bar / point / slice.
const dataLabelsPlugin = {
  id: 'sheetlyDataLabels',
  afterDatasetsDraw(chart) {
    const { ctx } = chart
    ctx.save(); ctx.font = '11px Calibri, Carlito, sans-serif'; ctx.fillStyle = '#404040'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'
    chart.data.datasets.forEach((ds, i) => chart.getDatasetMeta(i).data.forEach((elem, k) => {
      const v = ds.data[k], p = elem.tooltipPosition()
      if (v == null) return
      ctx.fillText(String(typeof v === 'object' ? v.y : Math.round(v * 100) / 100), p.x, p.y - 2)
    }))
    ctx.restore()
  },
}

function ChartView({ data }) {
  const ref = useRef(null)
  const [tick, redraw] = useReducer((x) => x + 1, 0)
  // Redraw when the workbook changes (values, formula results), batched per frame.
  useEffect(() => {
    if (!data?.source || !apiRef) return undefined
    let pending = false
    const sub = apiRef._injector.get(ICommandService).onCommandExecuted((cmd) => {
      if (cmd.type !== CommandType.MUTATION || pending) return
      pending = true
      soon(() => { pending = false; redraw() })
    })
    return () => sub.dispose()
  }, [data])
  useEffect(() => {
    if (!ref.current || !data) return undefined
    const chart = new Chart(ref.current, chartConfig(data, dataOf(data)))
    ref.current.__chart = chart
    return () => chart.destroy()
  }, [data, tick])
  const d = data?.source ? dataOf(data) : data
  return createElement('div', { className: 'sx-chart', role: 'img', 'aria-label': `${typeName(data?.type ?? 'column')} chart: ${d?.title || 'Chart'}`,
    'data-chart-type': data?.type ?? 'column',
    style: { width: '100%', height: '100%', background: '#fff', border: '1px solid #d9d9d9', boxSizing: 'border-box', padding: '6px' } },
  createElement('canvas', { ref }))
}

export function createCharts(api, dialogs, refs) {
  apiRef = api
  refsRef = refs
  const commands = api._injector.get(ICommandService)
  api.registerComponent(COMPONENT, ChartView)
  const wb = () => api.getActiveWorkbook()
  const ws = () => wb().getActiveSheet()
  const drawings = (() => { try { return api._injector.get(IDrawingManagerService) } catch { return null } })()

  /** Source range of a new chart: the selection, or the current region of a single cell. */
  function sourceRange() {
    const g0 = ws().getActiveRange().getRange()
    return g0.startRow === g0.endRow && g0.startColumn === g0.endColumn ? dialogs.currentRegion(g0.startRow, g0.startColumn) : g0
  }
  const newSpec = (type = 'column') => {
    const range = sourceRange()
    const { startRow, endRow, startColumn, endColumn } = range
    const g = { startRow, endRow, startColumn, endColumn }
    // The source range lives on as a hidden defined name, so it moves with inserted and deleted rows and columns.
    // local: the data is on the chart's own sheet (true for an embedded chart; F11's chart sheet sets it false).
    return { v: 2, type, source: { sheetId: ws().getSheetId(), range: g, ref: refs?.track(ws().getSheetId(), g) ?? undefined, local: true }, plotBy: null, title: null, legend: null, dataLabels: false, axisTitles: null }
  }
  /** For tests and Quick Analysis: the series Excel would plot from the selection. */
  function chartData() { const spec = newSpec(); return { data: dataOf(spec), range: spec.source.range, spec } }
  const colX = (s, c) => { let x = 0; for (let i = 0; i < c; i++) x += s.getColumnWidth(i); return x }
  const rowY = (s, r) => { let y = 0; for (let i = 0; i < r; i++) y += s.getRowHeight(i); return y }

  function embeddedChart(type = 'column') {
    const spec = newSpec(type), s = ws(), range = spec.source.range
    const x = colX(s, range.endColumn + 2), y = rowY(s, range.startRow)
    // Excel's default embedded chart: 5 x 3 inches (480 x 288 px).
    return s.addFloatDomToPosition({ componentKey: COMPONENT, allowTransform: true, data: spec, initPosition: { startX: x, endX: x + 480, startY: y, endY: y + 288 } })?.id ?? null
  }
  function chartSheet() {
    const spec = newSpec('column')
    spec.source.local = false
    const w = wb(), names = new Set(w.getSheets().map((s) => s.getSheetName()))
    let n = 1
    while (names.has(`Chart${n}`)) n++
    const index = w.getSheets().findIndex((s) => s.getSheetId() === ws().getSheetId())
    const sheet = w.insertSheet(`Chart${n}`, { index })
    w.setActiveSheet(sheet)
    sheet.setHiddenGridlines?.(true)
    setTimeout(() => sheet.addFloatDomToPosition({ componentKey: COMPONENT, allowTransform: false, data: spec,
      initPosition: { startX: 0, endX: Math.max(640, window.innerWidth - 80), startY: 0, endY: Math.max(400, window.innerHeight - 200) } }), 0)
    return true
  }

  // ---- the selected chart (a focused chart drawing, else the only chart on the sheet)
  const charts = () => (ws().getAllFloatDoms?.() ?? []).filter((d) => d.componentKey === COMPONENT)
  let lastFocused = null
  drawings?.focus$?.subscribe((list) => { const f = (list ?? []).find((d) => d.componentKey === COMPONENT); if (f) lastFocused = f.drawingId })
  function selected() {
    const all = charts()
    const focus = (drawings?.getFocusDrawings?.() ?? []).find((d) => d.componentKey === COMPONENT)
    return all.find((c) => c.id === focus?.drawingId) ?? all.find((c) => c.id === lastFocused) ?? (all.length === 1 ? all[0] : null)
  }
  /** The chart drawing that is selected right now (for the contextual Chart Design KeyTips). */
  const selectedNow = () => { const f = (drawings?.getFocusDrawings?.() ?? []).find((d) => d.componentKey === COMPONENT); return f ? charts().find((c) => c.id === f.drawingId) ?? null : null }
  const message = (text) => { openDialog({ title: 'Sheetly', width: 300, body: el('p', { style: { margin: 0 } }, text), cancelLabel: 'Close' }); return true }
  function update(change) {
    const c = selected()
    if (!c) return message('Select a chart first.')
    const spec = c.data?.source ? { ...c.data } : { ...newSpec(), ...c.data }
    ws().updateFloatDom(c.id, { data: { ...spec, ...change(spec) } })
    return true
  }
  const setType = (type) => update(() => ({ type }))
  // Switch Row/Column swaps what the series are made of.
  // (A chart whose series were picked one by one becomes one block over their range, as when Excel switches it.)
  const switchRowColumn = () => update((s) => ({ plotBy: (s.plotBy ?? (dataOf(s).byCols ? 'cols' : 'rows')) === 'cols' ? 'rows' : 'cols',
    source: { ...s.source, series: undefined, categories: undefined, categoriesRef: undefined } }))
  const toggle = (what) => update((s) => (what === 'title' ? { title: s.title === '' ? null : '' }
    : what === 'legend' ? { legend: (s.legend ?? 'bottom') === 'none' ? 'bottom' : 'none' }
      : what === 'dataLabels' ? { dataLabels: !s.dataLabels }
        : { axisTitles: s.axisTitles ? null : { x: 'Axis Title', y: 'Axis Title' } }))

  function typeDialog(title, current, onPick) {
    const r = radios('chart-type', CHART_TYPES.map(([id, name]) => [id, name]), current)
    const groups = []
    let last = null
    CHART_TYPES.forEach(([, , group], i) => { if (group !== last) { groups.push(el('div', { style: { fontWeight: 600, marginTop: groups.length ? '6px' : 0 } }, group)); last = group } groups.push(r.nodes[i]) })
    openDialog({ title, width: 300, body: el('div', { style: { maxHeight: '60vh', overflow: 'auto' } }, ...groups), onOk: () => { onPick(r.value) } })
    return true
  }
  const insertChartDialog = () => typeDialog('Insert Chart', 'column', (t) => embeddedChart(t))
  const changeTypeDialog = () => { const c = selected(); if (!c) return message('Select a chart first.'); return typeDialog('Change Chart Type', c.data?.type ?? 'column', setType) }
  function elementsDialog() {
    const c = selected()
    if (!c) return message('Select a chart first.')
    const s = c.data ?? {}
    const boxes = { title: check('Chart &Title', s.title !== ''), legend: check('&Legend', (s.legend ?? 'bottom') !== 'none'),
      dataLabels: check('&Data Labels', !!s.dataLabels), axisTitles: check('&Axis Titles', !!s.axisTitles) }
    openDialog({ title: 'Chart Elements', width: 260, body: el('div', {}, ...Object.values(boxes).map((b) => b.node)), onOk: () => {
      update((sp) => ({ title: boxes.title.input.checked ? (sp.title === '' ? null : sp.title) : '', legend: boxes.legend.input.checked ? (sp.legend && sp.legend !== 'none' ? sp.legend : 'bottom') : 'none',
        dataLabels: boxes.dataLabels.input.checked, axisTitles: boxes.axisTitles.input.checked ? (sp.axisTitles ?? { x: 'Axis Title', y: 'Axis Title' }) : null }))
    } })
    return true
  }

  /** A chart whose data is on its own sheet (source.local) keeps it there: on a sheet whose such charts point at
   *  another sheet (a copied sheet, also on redo) they are pointed at this sheet. The drawing is updated through its
   *  mutation alone, so this is part of the copy, not an undo step of its own. */
  function fixLocal() {
    const book = wb()
    if (!book) return
    const unitId = book.getId(), drawingsSvc = api._injector.get(ISheetDrawingService)
    for (const sheet of book.getSheets()) {
      const subUnitId = sheet.getSheetId()
      for (const c of (sheet.getAllFloatDoms?.() ?? []).filter((x) => x.componentKey === COMPONENT)) {
        const src = c.data?.source
        if (!src?.local) continue
        const main = src.ref ? refs?.resolve(src.ref) : { sheetId: src.sheetId, range: src.range }
        if (!main || main.sheetId === subUnitId) continue
        const from = main.sheetId
        // A copy leaves the original chart on the original sheet; a chart cut to another sheet keeps its data there.
        const original = book.getSheetBySheetId(from)
        if (!(original?.getAllFloatDoms?.() ?? []).some((x) => x.componentKey === COMPONENT && x.data?.source?.ref === src.ref)) continue
        const re = (name, range) => { const at = name ? refs?.resolve(name) : range ? { sheetId: from, range } : null; return at && at.sheetId === from ? refs?.track(subUnitId, at.range) ?? name : name }
        const next = { ...src, sheetId: subUnitId, ref: re(src.ref, src.range) }
        if (src.series) {
          next.series = src.series.map((x) => ({ ...x, valuesRef: re(x.valuesRef, x.values), xValuesRef: x.xValuesRef || x.xValues ? re(x.xValuesRef, x.xValues) : undefined, nameRef: x.nameRef ? re(x.nameRef) : undefined }))
          next.categoriesRef = src.categoriesRef || src.categories ? re(src.categoriesRef, src.categories) : undefined
        }
        const drawing = drawingsSvc.getDrawingByParam?.({ unitId, subUnitId, drawingId: c.id })
        if (!drawing) continue
        const { redo, objects } = drawingsSvc.getBatchUpdateOp([{ ...drawing, data: { ...drawing.data, source: next } }])
        commands.syncExecuteCommand(SetDrawingApplyMutation.id, { unitId, subUnitId, type: DrawingApplyType.UPDATE, op: redo, objects })
      }
    }
  }

  return { fixLocal, chartData, embeddedChart, chartSheet, insertChartDialog, changeTypeDialog, elementsDialog, setType, switchRowColumn, toggle, selected, selectedNow, charts,
    /** Where a chart's data is now (its references follow structural edits). */
    sourceOf: (c) => liveSource(c?.data?.source) }
}
