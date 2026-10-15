// Minimal Excel-style modal dialogs: title bar, body, OK/Cancel; Enter = OK, Esc = Cancel.
// Sheet shortcuts are suspended while a dialog is open so typing goes to the dialog.
import './dialog.css'

let shortcutService = null
export function setDialogShortcutService(svc) { shortcutService = svc }

export function el(tag, attrs = {}, ...children) {
  const n = document.createElement(tag)
  for (const [k, v] of Object.entries(attrs)) {
    if (v == null || v === false) continue
    if (k === 'class') n.className = v
    else if (k === 'style' && typeof v === 'object') Object.assign(n.style, v)
    else if (k.startsWith('on')) n.addEventListener(k.slice(2).toLowerCase(), v)
    else if (k in n && typeof v !== 'string') n[k] = v
    else n.setAttribute(k, v === true ? '' : v)
  }
  for (const c of children.flat()) if (c != null && c !== false) n.append(c instanceof Node ? c : document.createTextNode(String(c)))
  return n
}

/**
 * @param {{title: string, body: Node, onOk?: () => (boolean|void), okLabel?: string, extraButtons?: Node[],
 *          width?: number, initialFocus?: HTMLElement, noButtons?: boolean, className?: string}} opts
 * onOk returning false keeps the dialog open.
 */
export function openDialog(opts) {
  const disable = shortcutService?.forceDisable?.()
  const previousFocus = document.activeElement
  const close = () => {
    overlay.remove()
    disable?.dispose?.()
    previousFocus?.focus?.()
    opts.onClose?.()
  }
  const ok = () => { if (opts.onOk?.() !== false) close() }
  const buttons = opts.noButtons ? null : el('div', { class: 'sx-buttons' },
    ...(opts.extraButtons ?? []),
    el('button', { class: 'sx-btn sx-primary', type: 'button', onClick: ok }, opts.okLabel ?? 'OK'),
    el('button', { class: 'sx-btn', type: 'button', onClick: close }, opts.cancelLabel ?? 'Cancel'))
  const box = el('div', { class: `sx-dialog ${opts.className ?? ''}`, role: 'dialog', 'aria-modal': 'true', 'aria-label': opts.title,
    style: { width: (opts.width ?? 380) + 'px' } },
  el('div', { class: 'sx-title' }, el('span', {}, opts.title),
    el('button', { class: 'sx-close', type: 'button', 'aria-label': 'Close', onClick: close }, '×')),
  el('div', { class: 'sx-body' }, opts.body), buttons)
  const overlay = el('div', { class: 'sx-overlay' }, box)
  overlay.addEventListener('keydown', (e) => {
    e.stopPropagation()
    if (e.key === 'Escape') { e.preventDefault(); close() }
    else if (e.key === 'Enter' && !opts.noButtons && !(e.target instanceof HTMLTextAreaElement) && !(e.target instanceof HTMLButtonElement)) { e.preventDefault(); ok() }
  })
  for (const t of ['keyup', 'keypress']) overlay.addEventListener(t, (e) => e.stopPropagation())
  document.body.append(overlay)
  const focus = opts.initialFocus ?? box.querySelector('input,select,textarea,button.sx-primary')
  setTimeout(() => focus?.focus?.(), 0)
  return { close, ok, root: box }
}

export const field = (label, control) => el('label', { class: 'sx-field' }, el('span', {}, label), control)
export const check = (label, checked = false) => {
  const input = el('input', { type: 'checkbox', checked })
  return { input, node: el('label', { class: 'sx-check' }, input, label) }
}
export function radios(name, options, value) {
  const inputs = options.map(([v, label]) => {
    const input = el('input', { type: 'radio', name, value: v, checked: v === value })
    return { v, input, node: el('label', { class: 'sx-check' }, input, label) }
  })
  return { nodes: inputs.map((i) => i.node), get value() { return inputs.find((i) => i.input.checked)?.v } }
}
export const group = (legend, ...children) => el('fieldset', { class: 'sx-group' }, el('legend', {}, legend), ...children)
