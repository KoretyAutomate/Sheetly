// Ctrl+Q Quick Analysis: a gallery at the bottom-right of the selection with Excel's tabs — Formatting, Charts,
// Totals, Tables. Left/Right move between the options of a tab, Tab / Shift+Tab move between tabs, Enter (or a
// click) applies the option, Esc closes. A single cell expands to its current region; an empty one does nothing.
import { ILayoutService } from '@univerjs/ui'
import { el, openDialog, field } from './ui/dialog.js'

// Excel's default "Light Red Fill with Dark Red Text" highlight.
const LIGHT_RED = '#FFC7CE', DARK_RED = '#9C0006'

export function createQuickAnalysis(api, { dialogs, charts }) {
  const ws = () => api.getActiveWorkbook().getActiveSheet()
  const a1 = (r, c, r2 = r, c2 = c) => ws().getRange(r, c, r2 - r + 1, c2 - c + 1).getA1Notation()
  const abs = (r, c) => ws().getRange(r, c).getA1Notation().replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2')
  let panel = null

  const cf = (build) => (g) => {
    const s = ws(), rule = build(s.newConditionalFormattingRule()).setRanges([g]).build()
    s.addConditionalFormattingRule(rule)
    return true
  }
  const average = (g) => {
    let sum = 0, n = 0
    for (let r = g.startRow; r <= g.endRow; r++) for (let c = g.startColumn; c <= g.endColumn; c++) {
      const v = ws().getRange(r, c).getValue()
      if (typeof v === 'number') { sum += v; n++ }
    }
    return n ? Math.round((sum / n) * 100) / 100 : 0
  }
  function greaterThan(g) {
    const input = el('input', { type: 'number', value: String(average(g)), step: 'any', style: { width: '120px' } })
    openDialog({ title: 'Greater Than', width: 340, initialFocus: input,
      body: el('div', {}, field('Format cells that are GREATER THAN:', input), el('div', { style: { marginTop: '6px' } }, 'with Light Red Fill with Dark Red Text')),
      onOk: () => { cf((b) => b.whenNumberGreaterThan(Number(input.value)).setBackground(LIGHT_RED).setFontColor(DARK_RED))(g) } })
    return true
  }

  function textRule(title, label, apply) {
    return (g) => {
      const input = el('input', { type: 'text', value: '', style: { width: '160px' } })
      openDialog({ title, width: 340, initialFocus: input,
        body: el('div', {}, field(label, input), el('div', { style: { marginTop: '6px' } }, 'with Light Red Fill with Dark Red Text')),
        onOk: () => { cf((b) => apply(b, input.value).setBackground(LIGHT_RED).setFontColor(DARK_RED))(g) } })
      return true
    }
  }
  const hasNumbers = (g) => {
    for (let r = g.startRow; r <= g.endRow; r++) for (let c = g.startColumn; c <= g.endColumn; c++) if (typeof ws().getRange(r, c).getValue() === 'number') return true
    return false
  }
  const clearFormat = (g) => { ws().getRange(g.startRow, g.startColumn, g.endRow - g.startRow + 1, g.endColumn - g.startColumn + 1).clearConditionalFormatRules(); return true }

  // Totals: a results row below the selection (one formula per column) or a results column to its right.
  const totals = (kind, below) => (g) => {
    const s = ws()
    if (below) {
      const r = g.endRow + 1
      for (let c = g.startColumn; c <= g.endColumn; c++) {
        const col = a1(g.startRow, c, g.endRow, c)
        const f = { sum: `=SUM(${col})`, average: `=AVERAGE(${col})`, count: `=COUNTA(${col})`,
          percent: `=SUM(${col})/SUM(${a1(g.startRow, g.startColumn, g.endRow, g.endColumn).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')})`,
          running: `=SUM(${abs(g.startRow, g.startColumn)}:${s.getRange(g.endRow, c).getA1Notation()})` }[kind]
        s.getRange(r, c).setFormula(f)
        if (kind === 'percent') s.getRange(r, c).setNumberFormat('0.00%')
      }
    } else {
      const c = g.endColumn + 1
      for (let r = g.startRow; r <= g.endRow; r++) {
        const row = a1(r, g.startColumn, r, g.endColumn)
        const f = { sum: `=SUM(${row})`, average: `=AVERAGE(${row})`, count: `=COUNTA(${row})`,
          percent: `=SUM(${row})/SUM(${a1(g.startRow, g.startColumn, g.endRow, g.endColumn).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')})`,
          running: `=SUM(${abs(g.startRow, g.startColumn)}:${s.getRange(r, g.endColumn).getA1Notation()})` }[kind]
        s.getRange(r, c).setFormula(f)
        if (kind === 'percent') s.getRange(r, c).setNumberFormat('0.00%')
      }
    }
    return true
  }

  const TABS = [
    ['Formatting', [
      ['Data Bars', cf((b) => b.setDataBar({ min: { type: 'min' }, max: { type: 'max' }, isGradient: true, positiveColor: '#638EC6', nativeColor: '#FF0000', isShowValue: true }))],
      ['Color Scale', cf((b) => b.setColorScale([{ index: 0, color: '#F8696B', value: { type: 'min' } },
        { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } }, { index: 2, color: '#63BE7B', value: { type: 'max' } }]))],
      ['Icon Set', cf((b) => b.setIconSet({ isShowValue: true, iconConfigs: [
        { iconType: '3Arrows', iconId: '0', operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 } },
        { iconType: '3Arrows', iconId: '1', operator: 'greaterThanOrEqual', value: { type: 'percent', value: 33 } },
        { iconType: '3Arrows', iconId: '2', operator: 'lessThan', value: { type: 'percent', value: 33 } }] }))],
      ['Greater Than', greaterThan],
      ['Top 10%', cf((b) => b.setRank({ isBottom: false, isPercent: true, value: 10 }).setBackground(LIGHT_RED).setFontColor(DARK_RED))],
      ['Clear Format', clearFormat],
    ]],
    ['Charts', [['Clustered Column', () => charts.embeddedChart()]]],
    ['Totals', [
      ['Sum', totals('sum', true)], ['Average', totals('average', true)], ['Count', totals('count', true)],
      ['% Total', totals('percent', true)], ['Running Total', totals('running', true)],
      ['Sum (columns)', totals('sum', false)], ['Average (columns)', totals('average', false)], ['Count (columns)', totals('count', false)],
      ['% Total (columns)', totals('percent', false)], ['Running Total (columns)', totals('running', false)],
    ]],
    ['Tables', [['Table', () => dialogs.createTable()]]],
    // Excel's fifth tab; Sheetly has no sparklines, so its options are shown unavailable.
    ['Sparklines', [['Line', null], ['Column', null], ['Win/Loss', null]]],
  ]
  const NUMBER_FORMATTING = TABS[0][1]
  // Formatting options for text (Excel offers these when the selection holds no numbers).
  const TEXT_FORMATTING = [
    ['Text Contains', textRule('Text That Contains', 'Format cells that contain the text:', (b, v) => b.whenTextContains(v))],
    ['Duplicate Values', cf((b) => b.setDuplicateValues().setBackground(LIGHT_RED).setFontColor(DARK_RED))],
    ['Unique Values', cf((b) => b.setUniqueValues().setBackground(LIGHT_RED).setFontColor(DARK_RED))],
    ['Equal To', textRule('Equal To', 'Format cells that are EQUAL TO:', (b, v) => b.whenTextEqualTo(v))],
    ['Clear Format', clearFormat],
  ]

  const focusGrid = () => api._injector.get(ILayoutService).focus()
  function close() { panel?.remove(); panel = null }

  function open() {
    close()
    let g = ws().getActiveRange().getRange()
    if (g.startRow === g.endRow && g.startColumn === g.endColumn) {
      const v = ws().getRange(g.startRow, g.startColumn).getValue()
      if (v == null || v === '') return true
      g = dialogs.currentRegion(g.startRow, g.startColumn)
      if (g.startRow === g.endRow && g.startColumn === g.endColumn) return true
      dialogs.select([g])
    }
    let tab = 0, item = 0
    TABS[0][1] = hasNumbers(g) ? NUMBER_FORMATTING : TEXT_FORMATTING
    const tabsRow = el('div', { class: 'sx-qa-tabs', role: 'tablist' })
    const items = el('div', { class: 'sx-qa-items', role: 'toolbar' })
    const run = (i) => { const fn = TABS[tab][1][i][1]; if (!fn) return true; close(); focusGrid(); fn(g); return true }
    const render = () => {
      tabsRow.replaceChildren(...TABS.map(([name], i) => el('button', { type: 'button', role: 'tab', class: 'sx-qa-tab', 'aria-selected': String(i === tab),
        onClick: () => { tab = i; item = 0; render() } }, name.toUpperCase())))
      items.replaceChildren(...TABS[tab][1].map(([name, fn], i) => el('button', { type: 'button', class: 'sx-qa-item', tabindex: i === item ? '0' : '-1',
        'aria-disabled': fn ? null : 'true', title: fn ? null : 'Not available in Sheetly',
        'aria-current': String(i === item), onClick: () => run(i) }, name)))
      items.children[item]?.focus()
    }
    panel = el('div', { class: 'sx-qa', role: 'dialog', 'aria-label': 'Quick Analysis' }, tabsRow, items,
      el('div', { class: 'sx-qa-hint' }, 'Quick Analysis'))
    panel.addEventListener('keydown', (e) => {
      e.stopPropagation()
      const n = TABS[tab][1].length
      if (e.key === 'Escape') { e.preventDefault(); close(); focusGrid() }
      else if (e.key === 'ArrowRight') { e.preventDefault(); item = (item + 1) % n; render() }
      else if (e.key === 'ArrowLeft') { e.preventDefault(); item = (item - 1 + n) % n; render() }
      else if (e.key === 'Tab') { e.preventDefault(); tab = (tab + (e.shiftKey ? -1 : 1) + TABS.length) % TABS.length; item = 0; render() }
      else if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); run(item) }
    })
    for (const t of ['keyup', 'keypress']) panel.addEventListener(t, (e) => e.stopPropagation())
    document.body.append(panel)
    // Below and right of the selection, kept on screen.
    const rect = ws().getRange(g.endRow, g.endColumn).getCellRect?.()
    const canvas = [...document.querySelectorAll('canvas')].sort((a, b) => b.width * b.height - a.width * a.height)[0]?.getBoundingClientRect()
    let x = (canvas?.left ?? 0) + (rect ? rect.right : 200), y = (canvas?.top ?? 0) + (rect ? rect.bottom : 200)
    const pw = panel.offsetWidth, ph = panel.offsetHeight
    x = Math.max(8, Math.min(x, window.innerWidth - pw - 8)); y = Math.max(8, Math.min(y + 4, window.innerHeight - ph - 8))
    Object.assign(panel.style, { left: x + 'px', top: y + 'px' })
    render()
    const outside = (e) => { if (panel && !panel.contains(e.target)) close(); if (!panel) window.removeEventListener('pointerdown', outside, true) }
    window.addEventListener('pointerdown', outside, true)
    return true
  }

  return { open, close, isOpen: () => !!panel, tabs: () => TABS.map(([t, list]) => [t, list.map(([n]) => n)]) }
}
