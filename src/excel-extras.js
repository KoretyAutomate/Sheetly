// Remaining Excel shortcuts: Ctrl+` (show formulas), Shift+F3 (Insert Function), F3 (Paste Name),
// Ctrl+Shift+A (argument names), Ctrl+E (Flash Fill), F7 (Spelling), Shift+F10 (context menu),
// Ctrl+6 (hide/show objects).
import { CellValueType, ICommandService, IUniverInstanceService, LocaleService } from '@univerjs/core'
import { IContextMenuService } from '@univerjs/ui'
import { IFunctionService } from '@univerjs/engine-formula'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'
import { el, field, openDialog } from './ui/dialog.js'
import { learn, run } from './flashfill.js'

const EDITOR_UNIT = '__INTERNAL_EDITOR__DOCS_NORMAL'
const CATEGORY_NAMES = ['Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Database', 'Text', 'Logical',
  'Information', 'Engineering', 'Cube', 'Compatibility', 'Web', 'Array', 'Univer', 'User Defined', 'Defined Names']

export function createExtras(api) {
  const commands = api._injector.get(ICommandService)
  const locale = api._injector.get(LocaleService)
  const fns = () => api._injector.get(IFunctionService)
  const wb = () => api.getActiveWorkbook()
  const ws = () => wb().getActiveSheet()
  const t = (key) => (key ? locale.t(key) : '')

  // ---- editing helpers
  const editorText = () => api._injector.get(IUniverInstanceService).getUnit(EDITOR_UNIT)?.getBody?.()?.dataStream?.replace(/\r?\n$/, '') ?? ''
  async function insertInEditor(text, prefixIfNotEditing = '') {
    if (!wb().isCellEditing()) {
      commands.executeCommand('sheet.operation.set-cell-edit-visible-f2', { visible: true, eventType: 4, keycode: 113 })
      for (let i = 0; i < 40 && !wb().isCellEditing(); i++) await new Promise((r) => setTimeout(r, 25))
      if (!editorText().startsWith('=')) text = prefixIfNotEditing + text
    }
    await window.sheetly?.typeText?.(text)
    return true
  }

  // ---- Ctrl+`: show formulas instead of results (display only, like Excel's Show Formulas).
  // The sheets plugin starts with the first workbook, so the interceptor is registered on first use.
  let showFormulas = false, intercepting = false
  const ensureInterceptor = () => {
    if (intercepting) return
    intercepting = true
    api._injector.get(SheetInterceptorService).intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
      priority: 1000,
      handler: (cell, ctx, next) => {
        if (!showFormulas || !cell?.f) return next(cell)
        return next({ ...cell, v: cell.f, t: CellValueType.STRING, p: undefined })
      },
    })
  }
  function toggleShowFormulas() {
    ensureInterceptor()
    showFormulas = !showFormulas
    for (const s of wb().getSheets()) s.refreshCanvas?.()
    return true
  }

  // ---- Shift+F3: Insert Function
  function insertFunction() {
    const infos = [...fns().getDescriptions().values()].filter((d) => !d.functionName.startsWith('_'))
    const search = el('input', { type: 'text', placeholder: 'Type a brief description of what you want to do', 'aria-label': 'Search for a function' })
    const cat = el('select', { 'aria-label': 'Category' }, el('option', { value: 'all' }, 'All'),
      ...CATEGORY_NAMES.map((n, i) => el('option', { value: String(i) }, n)))
    const list = el('div', { class: 'sx-list', role: 'listbox', 'aria-label': 'Select a function' })
    const desc = el('div', { class: 'sx-sample', style: { minHeight: '48px' } })
    let chosen = null
    const signature = (d) => `${d.functionName}(${(d.functionParameter ?? []).map((p) => t(p.name)).join(',')})`
    const pick = (d, row) => {
      chosen = d
      list.querySelectorAll('.sel').forEach((x) => x.classList.remove('sel'))
      row?.classList.add('sel')
      desc.replaceChildren(el('strong', {}, signature(d)), el('br'), t(d.description) || t(d.abstract))
    }
    const fill = () => {
      const q = search.value.trim().toLowerCase()
      // Rank like Excel's search: exact name, then names starting with the text, then other matches.
      const rank = (d) => { const n = d.functionName.toLowerCase(); return n === q ? 0 : n.startsWith(q) ? 1 : n.includes(q) ? 2 : 3 }
      const items = infos.filter((d) => (cat.value === 'all' || String(d.functionType) === cat.value)
        && (!q || d.functionName.toLowerCase().includes(q) || (t(d.abstract) || '').toLowerCase().includes(q)))
        .sort((a, b) => (q ? rank(a) - rank(b) : 0) || a.functionName.localeCompare(b.functionName))
      list.replaceChildren(...items.slice(0, 600).map((d) => {
        const row = el('div', { role: 'option', onClick: () => pick(d, row), onDblclick: () => { pick(d, row); dlg.ok() } }, d.functionName)
        return row
      }))
      if (items[0]) pick(items[0], list.firstChild)
    }
    search.addEventListener('input', fill)
    cat.addEventListener('change', fill)
    const dlg = openDialog({ title: 'Insert Function', width: 460, initialFocus: search,
      body: el('div', {}, field('Search for a function:', search), field('Or select a category:', cat), field('Select a function:', list), desc),
      onOk: () => { if (chosen) insertInEditor(`${chosen.functionName}(`, '=') } })
    fill()
    return true
  }

  // ---- Ctrl+Shift+A: after a function name in the editor, insert its argument names.
  function insertArgumentNames() {
    const m = /([A-Za-z][A-Za-z0-9.]*)\(?$/.exec(editorText())
    if (!m) return true
    const d = fns().getDescriptions().get(m[1].toUpperCase())
    if (!d) return true
    const args = (d.functionParameter ?? []).map((p) => (p.require ? t(p.name) : `[${t(p.name)}]`)).join(',')
    window.sheetly?.typeText?.((editorText().endsWith('(') ? '' : '(') + args + ')')
    return true
  }

  // ---- F3: Paste Name
  function pasteName() {
    const names = (wb().getDefinedNames?.() ?? []).filter((n) => !n.getName().startsWith('SheetlyRef_')).map((n) => n.getName()).sort((a, b) => a.localeCompare(b))
    if (!names.length) return true // Excel shows nothing when the workbook has no names
    let chosen = names[0]
    const list = el('div', { class: 'sx-list', role: 'listbox', 'aria-label': 'Paste name' }, ...names.map((n, i) => {
      const row = el('div', { role: 'option', class: i === 0 ? 'sel' : '', onClick: () => { chosen = n; list.querySelectorAll('.sel').forEach((x) => x.classList.remove('sel')); row.classList.add('sel') },
        onDblclick: () => { chosen = n; dlg.ok() } }, n)
      return row
    }))
    const dlg = openDialog({ title: 'Paste Name', width: 280, body: field('Paste name', list),
      extraButtons: [el('button', { class: 'sx-btn', type: 'button', onClick: () => { pasteList(); dlg.close() } }, 'Paste List')],
      onOk: () => { insertInEditor(chosen, '=') } })
    return true
  }
  function pasteList() {
    const g = ws().getActiveRange().getRange()
    ;(wb().getDefinedNames?.() ?? []).filter((n) => !n.getName().startsWith('SheetlyRef_')).forEach((n, i) => {
      ws().getRange(g.startRow + i, g.startColumn).setValue(n.getName())
      ws().getRange(g.startRow + i, g.startColumn + 1).setValue(n.getFormulaOrRefString().startsWith('=') ? n.getFormulaOrRefString() : '=' + n.getFormulaOrRefString())
    })
  }

  // ---- Ctrl+E: Flash Fill
  function flashFill() {
    const s = ws(), g = s.getActiveRange().getRange()
    const col = g.startColumn
    const text = (r, c) => { const v = s.getRange(r, c).getValue(); return v == null ? '' : String(v) }
    const filled = (r, c) => text(r, c) !== ''
    // Source columns: the contiguous non-empty columns next to the target (left first, then right).
    const probeRow = Math.max(0, g.startRow - 1)
    const sources = []
    for (let c = col - 1; c >= 0 && filled(probeRow, c); c--) sources.unshift(c)
    for (let c = col + 1; filled(probeRow, c); c++) sources.push(c)
    if (!sources.length) return noPattern()
    let top = g.startRow
    while (top > 0 && sources.some((c) => filled(top - 1, c))) top--
    let bottom = g.startRow
    while (sources.some((c) => filled(bottom + 1, c))) bottom++
    const examples = []
    for (let r = top; r <= bottom; r++) if (filled(r, col)) examples.push({ inputs: sources.map((c) => text(r, c)), out: text(r, col) })
    if (!examples.length) return noPattern()
    const program = learn(examples)
    if (!program) return noPattern()
    let count = 0
    for (let r = top; r <= bottom; r++) {
      if (filled(r, col)) continue
      const inputs = sources.map((c) => text(r, c))
      if (inputs.every((x) => x === '')) continue
      const v = run(program, inputs)
      if (v != null) { s.getRange(r, col).setValue(v); count++ }
    }
    return count ? true : noPattern()
  }
  function noPattern() {
    openDialog({ title: 'Sheetly', width: 340, cancelLabel: 'Close',
      body: el('p', { style: { margin: 0 } }, 'Flash Fill didn’t find a pattern. Type a few more examples in the column, then try again.') })
    return true
  }

  // ---- F7: Spelling
  async function spelling() {
    const s = ws(), sheet = s.getSheet()
    const ignored = new Set(), changes = new Map()
    const issues = []
    const matrix = sheet.getCellMatrix()
    const cells = []
    matrix.forValue((r, c, cell) => { if (cell && cell.f == null && typeof cell.v === 'string') cells.push([r, c]) })
    cells.sort((a, b) => a[0] - b[0] || a[1] - b[1])
    for (const [r, c] of cells) {
      const v = String(s.getRange(r, c).getValue() ?? '')
      for (const m of v.matchAll(/\p{L}[\p{L}']*/gu)) if (await window.sheetly?.spell?.isMisspelled?.(m[0])) issues.push({ r, c, word: m[0] })
    }
    if (!issues.length) {
      openDialog({ title: 'Sheetly', width: 300, cancelLabel: 'Close', body: el('p', { style: { margin: 0 } }, 'Spell check complete. You’re good to go!') })
      return true
    }
    let i = 0
    const notIn = el('input', { type: 'text', readOnly: true, 'aria-label': 'Not in dictionary' })
    const change = el('input', { type: 'text', 'aria-label': 'Change to' })
    const sugg = el('div', { class: 'sx-list', role: 'listbox', 'aria-label': 'Suggestions', style: { height: '100px' } })
    const next = async () => {
      while (i < issues.length && (ignored.has(issues[i].word) || changes.has(issues[i].word))) {
        if (changes.has(issues[i].word)) replaceWord(issues[i], changes.get(issues[i].word))
        i++
      }
      if (i >= issues.length) { dlg.close(); return }
      const it = issues[i]
      s.getRange(it.r, it.c).activate()
      notIn.value = it.word
      const list = (await window.sheetly?.spell?.suggestions?.(it.word)) ?? []
      change.value = list[0] ?? it.word
      sugg.replaceChildren(...list.map((w) => el('div', { role: 'option', onClick: () => { change.value = w } }, w)))
    }
    const replaceWord = (it, to) => {
      const r = s.getRange(it.r, it.c), v = String(r.getValue() ?? '')
      r.setValue(v.replace(new RegExp(`(?<![\\p{L}'])${it.word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![\\p{L}'])`, 'u'), to))
    }
    const btn = (label, fn) => el('button', { class: 'sx-btn', type: 'button', onClick: async () => { await fn(); await next() } }, label)
    const dlg = openDialog({ title: 'Spelling: English (United States)', width: 440, noButtons: true,
      body: el('div', {}, field('Not in Dictionary:', notIn), field('Suggestions:', sugg), field('Change to:', change),
        el('div', { class: 'sx-buttons', style: { borderTop: 0, padding: 0, flexWrap: 'wrap' } },
          btn('Ignore Once', () => { i++ }), btn('Ignore All', () => { ignored.add(issues[i].word); i++ }),
          btn('Change', () => { replaceWord(issues[i], change.value); i++ }), btn('Change All', () => { changes.set(issues[i].word, change.value) }),
          el('button', { class: 'sx-btn', type: 'button', onClick: () => dlg.close() }, 'Cancel'))) })
    await next()
    return true
  }

  // ---- Shift+F10: context menu at the active cell (a real right-click there).
  function contextMenu() {
    const r = ws().getActiveRange()
    const rect = r.getCellRect?.()
    // getCellRect is relative to the grid canvas; convert to window coordinates.
    const canvas = [...document.querySelectorAll('canvas')].sort((a, b) => b.width * b.height - a.width * a.height)[0]
    if (!rect || !canvas) return true
    const box = canvas.getBoundingClientRect()
    // Same call Univer makes for a long-press: open the grid's context menu at the cell.
    const ev = new MouseEvent('contextmenu', { clientX: box.left + rect.left + Math.min(20, rect.width / 2), clientY: box.top + rect.top + rect.height / 2 })
    api._injector.get(IContextMenuService).triggerContextMenu(ev, 'contextMenu.mainArea', { unitId: wb().getId() })
    return true
  }

  // ---- Ctrl+6: hide / show objects (charts and other floating objects).
  let objectsHidden = false
  const objectsStyle = document.createElement('style')
  document.head.append(objectsStyle)
  function toggleObjects() {
    objectsHidden = !objectsHidden
    objectsStyle.textContent = objectsHidden ? '.sx-chart { visibility: hidden !important; }' : ''
    return true
  }

  // ---- Ctrl+Shift+U: expand / collapse the formula bar (Univer's own toggle button).
  function toggleFormulaBar() {
    const icon = document.querySelector('[data-u-comp="formula-bar"] svg.univerjs-icon-down-icon, [data-u-comp="formula-bar"] svg.univerjs-icon-up-icon')
    const target = icon?.closest('button,[role=button]') ?? icon?.parentElement
    target?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    return true
  }

  return { editorText, toggleShowFormulas, isShowingFormulas: () => showFormulas, insertFunction, insertArgumentNames, pasteName, flashFill, spelling,
    contextMenu, toggleObjects, toggleFormulaBar }
}

