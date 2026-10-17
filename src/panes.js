// F6 / Shift+F6: move the keyboard focus between the worksheet, the ribbon, an open task pane and the status bar
// (sheet tabs and zoom), as Excel does; Esc in the ribbon or status bar returns to the worksheet.
// Ctrl+F1: collapse the ribbon to its tab row, or expand it again (remembered per viewer).
// Window capture listeners, so they work wherever the focus is; inactive while a Sheetly dialog is open.
import { ILayoutService } from '@univerjs/ui'

const RIBBON = 'header[data-u-comp="headerbar"]', TASK_PANE = '[data-u-comp="right-sidebar"], .sx-task-pane:not([hidden])', STATUS = 'footer'
const STORE = 'sheetly.ribbonCollapsed'

export function installPanes(api) {
  const layout = api._injector.get(ILayoutService)
  const q = (sel) => document.querySelector(sel)
  const visible = (n) => { if (!n) return false; const r = n.getBoundingClientRect(); return r.width > 4 && r.height > 4 }
  const focusables = (root) => [...(root?.querySelectorAll('button, [tabindex]:not([tabindex="-1"]), input, select') ?? [])].filter((n) => !n.disabled && visible(n))

  const GIVES_BACK = new Set(['button', 'sheet-bar-append-button'])
  const statusTarget = () => {
    const list = focusables(q(STATUS)).filter((n) => !GIVES_BACK.has(n.closest('[data-u-comp]')?.dataset.uComp))
    return list.find((n) => n.tagName === 'INPUT') ?? list[0]
  }
  const regions = [
    { name: 'worksheet', has: () => true, focus: () => { layout.focus(); return true } },
    { name: 'ribbon', has: () => visible(q(RIBBON)),
      focus: () => { const tab = q(`${RIBBON} [role="tab"][aria-selected="true"]`) ?? focusables(q(RIBBON))[0]; tab?.focus(); return !!tab } },
    { name: 'taskpane', has: () => visible(q(TASK_PANE)) && focusables(q(TASK_PANE)).length > 0, focus: () => { const n = focusables(q(TASK_PANE))[0]; n?.focus(); return !!n } },
    // Excel's status-bar stop is the zoom control. Univer hands the focus straight back to the grid from its own
    // buttons, so only controls that keep it (zoom box, sheet tabs) are candidates.
    { name: 'statusbar', has: () => !!statusTarget(), focus: () => { const n = statusTarget(); n?.focus(); return !!n && document.activeElement === n } },
  ]
  const current = () => {
    const a = document.activeElement
    if (a?.closest?.(RIBBON)) return 'ribbon'
    if (a?.closest?.(TASK_PANE)) return 'taskpane'
    if (a?.closest?.(STATUS)) return 'statusbar'
    return 'worksheet'
  }
  function cycle(step) {
    const list = regions.filter((r) => r.has())
    let i = list.findIndex((r) => r.name === current())
    for (let n = 0; n < list.length; n++) {
      i = (i + step + list.length) % list.length
      if (list[i].focus()) return list[i].name
    }
    return current()
  }

  const setRibbon = (collapsed) => {
    document.body.classList.toggle('sx-ribbon-collapsed', collapsed)
    try { localStorage.setItem(STORE, collapsed ? '1' : '0') } catch { /* storage unavailable: not remembered */ }
    // Univer sizes the canvas from its container; let it re-measure.
    requestAnimationFrame(() => window.dispatchEvent(new Event('resize')))
  }
  let saved = false
  try { saved = localStorage.getItem(STORE) === '1' } catch { saved = false }
  if (saved) setRibbon(true)
  const style = document.createElement('style')
  style.textContent = `body.sx-ribbon-collapsed ${RIBBON} > div:has(> [data-u-comp="ribbon-toolbar"]) { display: none; }`
  document.head.append(style)

  const dialogOpen = () => !!document.querySelector('.sx-overlay')
  const menuOpen = () => !!document.querySelector('[role="menu"], [data-slot="dropdown-menu-content"]')
  window.addEventListener('keydown', (e) => {
    if (dialogOpen()) return
    if (e.key === 'F6' && !e.ctrlKey && !e.altKey && !e.metaKey) {
      e.preventDefault(); e.stopImmediatePropagation()
      cycle(e.shiftKey ? -1 : 1)
    } else if (e.key === 'F1' && e.ctrlKey && !e.shiftKey && !e.altKey && !e.metaKey) {
      e.preventDefault(); e.stopImmediatePropagation()
      setRibbon(!document.body.classList.contains('sx-ribbon-collapsed'))
    } else if (e.key === 'Escape' && ['ribbon', 'statusbar'].includes(current()) && !menuOpen()) {
      e.preventDefault(); e.stopImmediatePropagation()
      layout.focus()
    }
  }, true)

  return { cycle, current, ribbonCollapsed: () => document.body.classList.contains('sx-ribbon-collapsed'), setRibbon }
}
