// PivotTable model (no DOM / Univer): reads a source list, lays the report out as Excel's default Compact Form,
// and answers GETPIVOTDATA. Used by pivot.js and the GETPIVOTDATA function.
//
// Definition: { name, rows: [field], cols: [field], values: [{ field, fn, name? }], filters: [{ field, items: null | [item] }] }
// Fields are source column headers. Layout (Excel's defaults): row and column items sorted A to Z with "(blank)"
// last, subtotals at the top of each outer row group, column subtotals "<item> Total" after each outer column group,
// Grand Total row and column, several value fields shown side by side as an innermost column level ("Values").

// Largest / smallest of a list of any length (spreading a long list into Math.max overflows the call stack).
const maxOf = (list, start = -Infinity) => { let m = start; for (const x of list) if (x > m) m = x; return m }
const minOf = (list, start = Infinity) => { let m = start; for (const x of list) if (x < m) m = x; return m }
export const FUNCS = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp']
// Excel's caption prefixes; Count Numbers is captioned "Count of" too.
export const FN_LABEL = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' }
export const FN_MENU = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count Numbers', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' }
export const BLANK = '(blank)'
export const valueName = (v) => v.name || `${FN_LABEL[v.fn] ?? 'Sum'} of ${v.field}`

const isBlank = (x) => x == null || x === ''
const label = (x) => (isBlank(x) ? BLANK : x)
// Excel's order: numbers, text, logicals (FALSE, TRUE), then (blank).
const rank = (x) => (x === BLANK ? 3 : typeof x === 'number' ? 0 : typeof x === 'boolean' ? 2 : 1)
const cmpItems = (a, b) => {
  if (rank(a) !== rank(b)) return rank(a) - rank(b)
  if (typeof a === 'boolean') return Number(a) - Number(b)
  if (a === BLANK) return b === BLANK ? 0 : 1
  if (b === BLANK) return -1
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (typeof a === 'number') return -1
  if (typeof b === 'number') return 1
  return String(a).localeCompare(String(b), undefined, { sensitivity: 'base' })
}
// Items as Excel groups them: a number and a text that look alike are different items; text ignores case.
const sameItem = (a, b) => typeof a === typeof b && (typeof a === 'string' ? a.toLowerCase() === b.toLowerCase() : a === b)
const itemKey = (x) => `${typeof x}:${typeof x === 'string' ? x.toLowerCase() : String(x)}`
// Matching what a user typed (GETPIVOTDATA items, filter choices): "2024" finds the number 2024.
const same = (a, b) => (typeof a === 'number' || typeof b === 'number' ? Number(a) === Number(b) : String(a).toLowerCase() === String(b).toLowerCase())

/** Source list: the first row holds the field names. Returns null when a header is blank (Excel refuses those). */
export function readSource(get, g) {
  const fields = []
  for (let c = g.startColumn; c <= g.endColumn; c++) fields.push(get(g.startRow, c))
  if (fields.some(isBlank)) return null
  // Repeated headers become separate fields, numbered as Excel does (Sales, Sales2, ...).
  const seen = new Map()
  for (let k = 0; k < fields.length; k++) {
    const base = String(fields[k]), key = base.toLowerCase(), n = (seen.get(key) ?? 0) + 1
    seen.set(key, n)
    if (n > 1) { let name = `${base}${n}`; while (fields.some((f) => String(f).toLowerCase() === name.toLowerCase())) name += '_'; fields[k] = name }
  }
  const records = []
  for (let r = g.startRow + 1; r <= g.endRow; r++) {
    const rec = []
    for (let c = g.startColumn; c <= g.endColumn; c++) rec.push(get(r, c) ?? null)
    if (rec.some((x) => !isBlank(x))) records.push(rec)
  }
  return { fields: fields.map(String), records }
}

/** A field's values: numbers when every non-blank value is a number. */
export const isNumericField = (src, field) => {
  const i = src.fields.indexOf(field)
  const vals = src.records.map((r) => r[i]).filter((x) => !isBlank(x))
  return vals.length > 0 && vals.every((x) => typeof x === 'number')
}

/** Excel's choice for a value field added from the list: Sum for numbers, Count otherwise. */
export const defaultFn = (src, field) => (isNumericField(src, field) ? 'sum' : 'count')

// Error values (a source formula returning #DIV/0!, say) come through as their text.
const ERROR_VALUES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#GETTING_DATA', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!'])
export const isErrorValue = (x) => typeof x === 'string' && ERROR_VALUES.has(x.toUpperCase())

/** A value field's result; an error in the values is the result for every function but Count (as in Excel). */
export function aggregate(fn, values) {
  if (fn !== 'count') { const e = values.find(isErrorValue); if (e !== undefined) return e }
  const nums = values.filter((x) => typeof x === 'number')
  const n = nums.length, sum = nums.reduce((a, b) => a + b, 0)
  // A sample statistic of one number cannot be computed: #DIV/0!, as Excel shows it (no numbers at all: empty).
  const varOf = (pop) => { if (!n) return null; if (!pop && n < 2) return '#DIV/0!'; const m = sum / n; return nums.reduce((a, b) => a + (b - m) ** 2, 0) / (pop ? n : n - 1) }
  switch (fn) {
    case 'count': { const k = values.filter((x) => !isBlank(x)).length; return k || null }
    case 'countNums': return n || null
    case 'average': return n ? sum / n : null
    case 'max': return n ? maxOf(nums) : null
    case 'min': return n ? minOf(nums) : null
    case 'product': return n ? nums.reduce((a, b) => a * b, 1) : null
    case 'stdDev': { const v = varOf(false); return typeof v === 'number' ? Math.sqrt(v) : v }
    case 'stdDevp': { const v = varOf(true); return v == null ? null : Math.sqrt(v) }
    case 'var': return varOf(false)
    case 'varp': return varOf(true)
    default: return values.length ? sum : null
  }
}

/** The report item a typed value names: the item of the same type first ("1" is the text item when there is one),
 *  else one that matches loosely (the number 2024 for "2024"). */
export function resolveItem(src, records, field, wanted) {
  const i = src.fields.indexOf(field)
  const items = [...new Map(records.map((r) => [itemKey(label(r[i])), label(r[i])])).values()]
  return items.find((x) => sameItem(x, wanted)) ?? items.find((x) => same(x, wanted)) ?? null
}

/** Records left after the report filters. */
export function filtered(def, src) {
  const idx = (f) => src.fields.indexOf(f)
  const chosen = (def.filters ?? []).filter((fl) => fl.items).map((fl) => [idx(fl.field), fl.items.map((it) => resolveItem(src, src.records, fl.field, it)).filter((x) => x !== null)])
  return src.records.filter((rec) => chosen.every(([i, items]) => items.some((it) => sameItem(label(rec[i]), it))))
}

// Item tree of some fields over records: [{ item, path, children }].
function tree(fields, records, src, path = []) {
  if (!fields.length) return []
  const i = src.fields.indexOf(fields[0])
  // One pass groups the records by item (no rescanning per item).
  const groups = new Map()
  for (const r of records) { const it = label(r[i]), k = itemKey(it); let g = groups.get(k); if (!g) groups.set(k, g = { item: it, recs: [] }); g.recs.push(r) }
  return [...groups.values()].sort((a, b) => cmpItems(a.item, b.item))
    .map(({ item, recs }) => ({ item, path: [...path, item], children: tree(fields.slice(1), recs, src, [...path, item]) }))
}
/** A field's distinct items in report order (blank last, as null), and the index of a value among them: what an
 *  .xlsx PivotCache lists as the field's shared items. */
export function cacheItems(src, field) {
  const i = src.fields.indexOf(field), seen = new Map()
  for (const r of src.records) { const k = itemKey(label(r[i])); if (!seen.has(k)) seen.set(k, isBlank(r[i]) ? null : r[i]) }
  const items = [...seen.values()].sort((a, b) => cmpItems(label(a), label(b)))
  const at = new Map(items.map((x, k) => [itemKey(label(x)), k]))
  return { items, indexOf: (x) => at.get(itemKey(label(x))) }
}
const pathKey = (path) => path.map(itemKey).join('\u0001')

/**
 * The report as a grid of cells { v, kind, indent? } plus where things are:
 * kind: 'corner' | 'colHeader' | 'rowHeader' | 'rowLabel' | 'value' | 'subtotal' | 'grand' | 'filterLabel' | 'filterValue'
 */
export function layout(def, src) {
  const recs = filtered(def, src)
  const V = (def.values ?? []).filter((v) => src.fields.includes(v.field))
  const rowF = (def.rows ?? []).filter((f) => src.fields.includes(f)), colF = (def.cols ?? []).filter((f) => src.fields.includes(f))
  // The values of every (row path prefix, column path prefix) pair, gathered in one pass over the records.
  const rIdx = rowF.map((f) => src.fields.indexOf(f)), cIdx = colF.map((f) => src.fields.indexOf(f))
  const buckets = new Map()
  for (const rec of recs) {
    const rk = [''], ck = ['']
    for (const i of rIdx) rk.push((rk.length > 1 ? rk[rk.length - 1] + '\u0001' : '') + itemKey(label(rec[i])))
    for (const i of cIdx) ck.push((ck.length > 1 ? ck[ck.length - 1] + '\u0001' : '') + itemKey(label(rec[i])))
    for (const a of rk) for (const b of ck) { const key = `${a}\u0002${b}`; let list = buckets.get(key); if (!list) buckets.set(key, list = []); list.push(rec) }
  }
  const valueOf = (rpath, cpath, v) => {
    const vi = src.fields.indexOf(v.field)
    const sel = buckets.get(`${pathKey(rpath)}\u0002${pathKey(cpath)}`) ?? []
    return aggregate(v.fn, sel.map((r) => r[vi]))
  }
  // Column leaves: each a { cpath, v, kind } — items, subtotals of outer groups, the grand total, per value field.
  const colTree = tree(colF, recs, src)
  const cols = []
  const walkCols = (nodes, depth) => {
    for (const n of nodes) {
      if (n.children.length) { walkCols(n.children, depth + 1); for (const v of V.length ? V : [null]) cols.push({ cpath: n.path, v, kind: 'subtotal', label: `${n.item} Total` }) }
      else for (const v of V.length ? V : [null]) cols.push({ cpath: n.path, v, kind: 'value' })
    }
  }
  walkCols(colTree, 0)
  // Outer-level subtotals only between groups: the innermost level has none.
  if (colF.length <= 1) for (let i = cols.length - 1; i >= 0; i--) if (cols[i].kind === 'subtotal') cols.splice(i, 1)
  if (!colF.length) for (const v of V) cols.push({ cpath: [], v, kind: 'value' })
  else for (const v of V.length ? V : [null]) cols.push({ cpath: [], v, kind: 'grand', label: V.length > 1 ? `Total ${valueName(v)}` : 'Grand Total' })
  const grid = []
  // Header rows: one per column field, plus the value captions when there are several value fields. With no column
  // field the value captions share the Row Labels row.
  const headerLevels = colF.length ? colF.length + (V.length > 1 ? 1 : 0) : 0
  if (!rowF.length && !colF.length && V.length) {
    // Values only: one header row of value names and one row of totals.
    grid.push(V.map((v) => ({ v: valueName(v), kind: 'colHeader' })))
    grid.push(V.map((v) => ({ v: valueOf([], [], v), kind: 'grand' })))
    return { grid, rows: rowF, cols: colF, values: V }
  }
  if (headerLevels) {
    grid.push([{ v: V.length === 1 ? valueName(V[0]) : null, kind: 'corner' }, { v: 'Column Labels', kind: 'colHeader' }, ...cols.slice(1).map(() => ({ v: null, kind: 'colHeader' }))])
    for (let lvl = 0; lvl < headerLevels; lvl++) {
      const last = lvl === headerLevels - 1
      const row = [{ v: last ? (rowF.length ? 'Row Labels' : null) : null, kind: 'rowHeader' }]
      let prev = null
      for (const c of cols) {
        let text = null
        if (lvl < colF.length) {
          if (c.kind === 'grand') text = lvl === 0 ? c.label : null
          else if (c.kind === 'subtotal') text = lvl === c.cpath.length - 1 ? c.label : null
          else { const key = c.cpath.slice(0, lvl + 1).join('\u0001'); text = key !== prev ? c.cpath[lvl] : null; prev = key }
        } else text = c.kind === 'grand' ? null : c.v ? valueName(c.v) : null
        row.push({ v: text, kind: 'colHeader' })
      }
      grid.push(row)
    }
  } else grid.push([{ v: rowF.length ? 'Row Labels' : null, kind: 'rowHeader' }, ...cols.map((c) => ({ v: c.v ? valueName(c.v) : null, kind: 'colHeader' }))])
  // Body: compact form, one column of row labels indented by level, subtotals on the group's own row.
  const walkRows = (nodes, depth) => {
    for (const n of nodes) {
      const outer = n.children.length > 0
      grid.push([{ v: n.item, kind: outer ? 'subtotalLabel' : 'rowLabel', indent: depth },
        ...cols.map((c) => ({ v: c.v ? valueOf(n.path, c.cpath, c.v) : null, kind: outer ? 'subtotal' : 'value' }))])
      if (outer) walkRows(n.children, depth + 1)
    }
  }
  walkRows(tree(rowF, recs, src), 0)
  // The totals row: "Grand Total" under row items; with column fields only, the one values row has no label.
  if (rowF.length || colF.length) grid.push([{ v: rowF.length ? 'Grand Total' : null, kind: 'grandLabel' }, ...cols.map((c) => ({ v: c.v ? valueOf([], c.cpath, c.v) : null, kind: 'grand' }))])
  return { grid, rows: rowF, cols: colF, values: V }
}

/** Report filter cells above the report: [field, "(All)" | item | "(Multiple Items)"]. */
export const filterCells = (def) => (def.filters ?? []).map((f) => [f.field, !f.items ? '(All)' : f.items.length === 1 ? f.items[0] : '(Multiple Items)'])

/**
 * GETPIVOTDATA(data_field, pivot, field1, item1, ...): the value for a combination of items shown in the report.
 * data_field is the value field's caption ("Sum of Sales") or its source field ("Sales"). Fields must be on rows or
 * columns and items must be in the report, else #REF! (as in Excel).
 */
export function getPivotData(def, src, dataField, pairs) {
  const V = (def.values ?? [])
  const v = V.find((x) => same(valueName(x), dataField)) ?? V.find((x) => same(x.field, dataField))
  if (!v) return { error: '#REF!' }
  // Fields on rows or columns, and report filters (whose shown items are the selected ones).
  const shown = [...(def.rows ?? []), ...(def.cols ?? []), ...(def.filters ?? []).map((f) => f.field)]
  const recs = filtered(def, src)
  let sel = recs
  for (const [field, item] of pairs) {
    const f = shown.find((x) => same(x, field))
    if (!f) return { error: '#REF!' }
    const i = src.fields.indexOf(f)
    const it = resolveItem(src, recs, f, item)
    if (it === null) return { error: '#REF!' }
    sel = sel.filter((r) => sameItem(label(r[i]), it))
  }
  const out = aggregate(v.fn, sel.map((r) => r[src.fields.indexOf(v.field)]))
  return out == null ? { error: '#REF!' } : isErrorValue(out) ? { error: out.toUpperCase() } : { value: out }
}
