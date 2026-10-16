// Excel ribbon KeyTips: press and release Alt (or F10) to show the tab letters, then type the letters of a command,
// e.g. Alt H O I = AutoFit Column Width. Holding Alt while pressing the first letter works too. Esc goes back one
// level; Alt or F10 again, a click elsewhere, or leaving the window hides them. Office 2003 menu keys (Alt E S, ...)
// are accepted where no ribbon tab uses the letter, as in Excel.
//
// The key listener is registered when this module is imported, before Univer creates its own window listener, so
// it runs first and can keep KeyTip letters away from the sheet. The tree itself is supplied later (setKeyTipTree).
import './keytips.css'
import { el } from './ui/dialog.js'

let tree = null // { tabs: Node[], legacy: Node[] }  Node = { k, label, run?, items? }
let canStart = () => false
const state = { active: false, path: [], buffer: '', altDown: false, altOther: false, altAgain: false, legacy: false, typed: [] }

export function setKeyTipTree(t, { canStart: c }) { tree = t; canStart = c }
export const keyTipState = () => ({ active: state.active, path: state.path.map((n) => n.k), buffer: state.buffer, legacy: state.legacy })

const panel = el('div', { class: 'sx-keytips', role: 'menubar', 'aria-label': 'KeyTips', hidden: true })
const mount = () => { if (!panel.isConnected) document.body.append(panel) }
const overlayOpen = () => !!document.querySelector('.sx-overlay')
const keyChar = (e) => (/^Key[A-Z]$/.test(e.code) ? e.code.slice(3) : /^Digit\d$/.test(e.code) ? e.code.slice(5) : null)
const level = () => (state.path.length ? state.path[state.path.length - 1].items : state.legacy ? tree.legacy : tree.tabs)

function start(first) {
  state.active = true; state.path = []; state.buffer = ''; state.legacy = false; state.typed = []
  render()
  if (first) type(first)
}
function stop() {
  state.active = false; state.path = []; state.buffer = ''; state.legacy = false; state.typed = []
  panel.hidden = true
}
function back() {
  if (state.buffer) state.buffer = ''
  else if (state.path.length) { state.path.pop(); state.typed.pop() }
  else if (state.legacy) { state.legacy = false; state.typed = [] }
  else return stop()
  render()
}
function choose(node) {
  state.buffer = ''
  if (node.items) { state.path.push(node); state.typed.push(node.k); render(); return }
  stop()
  node.run?.()
}
function type(ch) {
  const buf = state.buffer + ch
  // At the top level a letter that no tab uses reaches the Office 2003 menus.
  if (!state.path.length && !state.legacy && !state.buffer) {
    const tab = tree.tabs.find((n) => n.k === ch)
    if (!tab) { const menu = tree.legacy.find((n) => n.k === ch); if (menu) { state.legacy = true; state.typed = []; return choose(menu) } return }
  }
  const items = level() ?? []
  const exact = items.find((n) => n.k === buf)
  if (exact) return choose(exact)
  if (items.some((n) => n.k.startsWith(buf))) { state.buffer = buf; render() }
}

function badge(text) { return el('span', { class: 'sx-keytip' }, text) }
function render() {
  mount()
  panel.hidden = false
  const tabs = el('div', { class: 'sx-keytips-tabs' }, ...tree.tabs.map((t) => el('button', { type: 'button', class: 'sx-keytips-tab',
    'aria-current': state.path[0] === t ? 'true' : null, onClick: () => { state.path = []; state.legacy = false; choose(t) } }, t.label, badge(t.k))))
  const rows = [tabs]
  if (state.legacy) {
    rows.push(el('div', { class: 'sx-keytips-legacy' }, `Office 2003 access key: ALT, ${state.typed.join(', ')}, `,
      'continue typing the menu key sequence from an earlier version of Office or press Escape to cancel.'))
  } else if (state.path.length) {
    const crumb = state.path.map((n) => n.label).join(' › ')
    const items = level().filter((n) => n.k.startsWith(state.buffer))
    rows.push(el('div', { class: 'sx-keytips-crumb' }, crumb),
      el('div', { class: 'sx-keytips-items', role: 'menu' }, ...items.map((n) => el('button', { type: 'button', role: 'menuitem', class: 'sx-keytips-item',
        onClick: () => choose(n) }, badge(n.k), n.label, n.items ? ' ▸' : ''))))
  }
  panel.replaceChildren(...rows)
}

window.addEventListener('keydown', (e) => {
  if (!tree || overlayOpen()) { if (state.active) stop(); return }
  if (state.active) {
    e.preventDefault(); e.stopImmediatePropagation()
    if (e.repeat && e.key === 'Alt') return
    if (e.key === 'Alt') { state.altDown = true; state.altOther = false; state.altAgain = true; return }
    state.altOther = true; state.altAgain = false
    if (e.key === 'Escape') return back()
    if (e.key === 'F10' && !e.shiftKey && !e.ctrlKey) return stop()
    const ch = keyChar(e)
    if (ch && !e.ctrlKey && !e.metaKey) type(ch)
    return
  }
  if (e.key === 'Alt') {
    if (!e.repeat) { state.altDown = true; state.altOther = false }
    e.preventDefault() // keep the window menu bar from taking Alt
    return
  }
  if (state.altDown) state.altOther = true
  // Alt+letter (Alt still held): straight into that tab or Office 2003 menu.
  const ch = keyChar(e)
  if (e.altKey && !e.ctrlKey && !e.shiftKey && !e.metaKey && ch && (tree.tabs.some((n) => n.k === ch) || tree.legacy.some((n) => n.k === ch)) && canStart()) {
    e.preventDefault(); e.stopImmediatePropagation()
    start(ch)
    return
  }
  if (e.key === 'F10' && !e.altKey && !e.ctrlKey && !e.shiftKey && !e.metaKey && canStart()) {
    e.preventDefault(); e.stopImmediatePropagation()
    start()
  }
}, true)

window.addEventListener('keyup', (e) => {
  if (e.key !== 'Alt') return
  e.preventDefault()
  const alone = state.altDown && !state.altOther
  state.altDown = false
  if (!tree || overlayOpen()) return
  if (state.active) { if (alone && state.altAgain) stop(); state.altAgain = false; return }
  if (alone && canStart()) start()
}, true)

window.addEventListener('blur', () => { state.altDown = false; if (state.active) stop() })
window.addEventListener('pointerdown', (e) => {
  if (state.altDown) state.altOther = true
  if (state.active && !panel.contains(e.target)) stop()
}, true)
window.addEventListener('wheel', () => { if (state.altDown) state.altOther = true }, true)
