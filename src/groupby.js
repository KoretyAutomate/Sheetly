// GROUPBY and PIVOTBY, following Excel's documented argument semantics.
// Replaces Univer's GROUPBY, which only accepted 7 hard-coded aggregators; here the function
// argument may be any function name (eta lambda), a LAMBDA, or HSTACK of several.
import {
  ArrayValueObject, BaseFunction, ErrorValueObject, FunctionType, IFunctionService, NullValueObject,
  StringValueObject,
} from '@univerjs/engine-formula'

const NULL = () => NullValueObject.create()
// Empty cells inside a spilled result display as 0; Excel's GROUPBY/PIVOTBY show them blank.
const BLANK = () => StringValueObject.create('')
const err = (e) => ErrorValueObject.create(e)

// Excel's sort order: numbers < text (case-insensitive) < FALSE < TRUE < errors < blanks.
function rank(v) {
  if (v == null || v.isNull()) return 5
  if (v.isError()) return 4
  if (v.isBoolean()) return 3
  if (v.isNumber()) return 0
  return 1
}
export function compareValues(a, b) {
  const ra = rank(a), rb = rank(b)
  if (ra !== rb) return ra - rb
  if (ra === 0) return +a.getValue() - +b.getValue()
  if (ra === 1) return String(a.getValue()).toLowerCase().localeCompare(String(b.getValue()).toLowerCase())
  if (ra === 3) return (a.getValue() ? 1 : 0) - (b.getValue() ? 1 : 0)
  return 0
}
const keyOf = (v) => (v == null || v.isNull() ? 'n:' : v.isError() ? 'e:' + v.getValue()
  : v.isNumber() ? 'd:' + +v.getValue() : v.isBoolean() ? 'b:' + v.getValue() : 's:' + String(v.getValue()).toLowerCase())

function toGrid(variant) {
  let v = variant
  if (v == null) return null
  if (v.isReferenceObject?.()) v = v.toArrayValueObject()
  if (v.isError()) return v
  if (!v.isArray()) return { rows: 1, cols: 1, at: () => v }
  return { rows: v.getRowCount(), cols: v.getColumnCount(), at: (r, c) => v.get(r, c) ?? NULL() }
}
function scalarNumber(variant, dflt) {
  if (variant == null) return dflt
  let v = variant
  if (v.isReferenceObject?.()) v = v.getFirstCell()
  if (v.isArray?.()) v = v.get(0, 0)
  if (v == null || v.isNull()) return dflt
  if (v.isError()) return v
  const n = Number(v.getValue())
  return Number.isFinite(n) ? Math.trunc(n) : err('#VALUE!')
}
function numberList(variant) {
  if (variant == null) return []
  const g = toGrid(variant)
  if (g == null || g.isError?.()) return g
  const out = []
  for (let r = 0; r < g.rows; r++) for (let c = 0; c < g.cols; c++) {
    const x = g.at(r, c)
    if (x.isNull()) continue
    const n = Number(x.getValue())
    if (!Number.isFinite(n) || n === 0) return err('#VALUE!')
    out.push(Math.trunc(n))
  }
  return out
}

// Builds the list of aggregators from the AST node of the `function` argument.
function aggregatorsFrom(node, getVariant, fns) {
  const token = String(node.getToken?.() ?? '').toUpperCase().replace(/^(_XLFN\.|_XLWS\.|_XLETA\.)+/, '')
  if (token === 'HSTACK') return node.getChildren().flatMap((ch) => aggregatorsFrom(ch, getVariant, fns))
  if (token !== 'LAMBDA' && node.getChildren().length === 0 && fns.hasExecutor(token)) {
    const ex = fns.getExecutor(token)
    return [{ name: token, run: (subset, all) => (token === 'PERCENTOF' ? ex.calculate(subset, all) : ex.calculate(subset)) }]
  }
  const lam = getVariant(node)
  if (lam?.isLambda?.()) {
    const n = lam.getLambdaPrivacyVarKeys().length
    if (n < 1 || n > 2) return [{ error: err('#VALUE!') }]
    return [{ name: 'LAMBDA', run: (subset, all) => (n === 1 ? lam.execute(subset) : lam.execute(subset, all)) }]
  }
  if (lam?.isError?.()) return [{ error: lam }]
  return [{ error: err('#VALUE!') }]
}

function column(values) {
  return ArrayValueObject.create({ calculateValueList: values.length ? values.map((v) => [v]) : [[NULL()]],
    rowCount: Math.max(1, values.length), columnCount: 1, unitId: '', sheetId: '', row: 0, column: 0 })
}
function scalarResult(v) {
  if (v == null) return NULL()
  if (v.isReferenceObject?.()) v = v.toArrayValueObject()
  if (v.isArray?.()) return v.getRowCount() === 1 && v.getColumnCount() === 1 ? v.get(0, 0) : err('#CALC!')
  return v
}

// Hierarchical group tree for one axis (rows or columns).
function buildAxis(records, keyIndexOf, levels, depth, relationship, sortSpec, sortValue) {
  // depth: 0 none, 1 grand, 2 grand+subtotals; negative = totals placed before their group.
  const subtotals = Math.abs(depth) >= 2 && relationship === 0 && levels > 1
  const before = depth < 0
  const slots = []
  const groupsAt = (recs, level) => {
    const map = new Map()
    for (const r of recs) {
      const k = keyIndexOf(r)[level], id = keyOf(k)
      if (!map.has(id)) map.set(id, { key: k, recs: [] })
      map.get(id).recs.push(r)
    }
    const gs = [...map.values()]
    const spec = sortSpec.find((s) => Math.abs(s) - 1 === level) ?? sortSpec.find((s) => Math.abs(s) - 1 >= levels)
    gs.sort((a, b) => {
      if (spec != null && Math.abs(spec) - 1 >= levels) {
        const c = compareValues(sortValue(a.recs, Math.abs(spec) - 1 - levels), sortValue(b.recs, Math.abs(spec) - 1 - levels))
        if (c !== 0) return spec < 0 ? -c : c
      }
      const c = compareValues(a.key, b.key)
      return spec != null && Math.abs(spec) - 1 === level && spec < 0 ? -c : c
    })
    return gs
  }
  if (relationship === 1) {
    // Table: every distinct full key is one row, sorted by the key columns in order.
    const map = new Map()
    for (const r of records) {
      const ks = keyIndexOf(r).slice(0, levels), id = ks.map(keyOf).join('|')
      if (!map.has(id)) map.set(id, { keys: ks, recs: [] })
      map.get(id).recs.push(r)
    }
    const gs = [...map.values()].sort((a, b) => {
      for (const s of sortSpec.length ? sortSpec : Array.from({ length: levels }, (_, i) => i + 1)) {
        const i = Math.abs(s) - 1
        const c = i < levels ? compareValues(a.keys[i], b.keys[i]) : compareValues(sortValue(a.recs, i - levels), sortValue(b.recs, i - levels))
        if (c !== 0) return s < 0 ? -c : c
      }
      return 0
    })
    for (const g of gs) slots.push({ keys: g.keys, recs: g.recs, kind: 'detail', parent: records })
  } else {
    const walk = (recs, level, prefix) => {
      for (const g of groupsAt(recs, level)) {
        const keys = [...prefix, g.key]
        if (level === levels - 1) { slots.push({ keys, recs: g.recs, kind: 'detail', parent: recs }); continue }
        const sub = { keys: [...keys, ...Array(levels - keys.length).fill(null).map(BLANK)], recs: g.recs, kind: 'subtotal', parent: recs }
        if (subtotals && before) slots.push(sub)
        walk(g.recs, level + 1, keys)
        if (subtotals && !before) slots.push(sub)
      }
    }
    if (levels === 0) slots.push({ keys: [], recs: records, kind: 'detail', parent: records })
    else walk(records, 0, [])
  }
  if (depth !== 0 && levels > 0) {
    const grand = { keys: [StringValueObject.create('Total'), ...Array(levels - 1).fill(null).map(BLANK)], recs: records, kind: 'grand', parent: records }
    if (before) slots.unshift(grand); else slots.push(grand)
  }
  return slots
}

class PivotFunction extends BaseFunction {
  constructor(name) {
    super(name)
    this._getFns = () => PivotFunction.getFns()
    this.minParams = name === 'GROUPBY' ? 3 : 4
    this.maxParams = name === 'GROUPBY' ? 8 : 11
    this.needsReferenceObject = true
    this.needsAstChildren = true
  }
  calculateAst(children, getVariant) {
    const pivot = this.name === 'PIVOTBY'
    // A skipped argument (e.g. the 4th in GROUPBY(a,b,SUM,,0)) arrives as a bare null value: treat it as omitted.
    const arg = (i) => {
      if (children[i] == null) return undefined
      const v = getVariant(children[i])
      return v == null || (!v.isReferenceObject?.() && v.isNull?.()) ? undefined : v
    }
    const idx = pivot
      ? { row: 0, col: 1, val: 2, fn: 3, hdr: 4, rdepth: 5, rsort: 6, cdepth: 7, csort: 8, filter: 9, rel: 10 }
      : { row: 0, val: 1, fn: 2, hdr: 3, rdepth: 4, rsort: 5, filter: 6, relationship: 7 }
    const rowG = toGrid(arg(idx.row)), valG = toGrid(arg(idx.val))
    const colG = pivot ? toGrid(arg(idx.col)) : { rows: rowG?.rows ?? 0, cols: 0, at: NULL }
    for (const g of [rowG, valG, colG]) if (g == null) return err('#VALUE!'); else if (g.isError?.()) return g
    if (valG.rows !== rowG.rows || colG.rows !== rowG.rows) return err('#VALUE!')
    const aggs = aggregatorsFrom(children[idx.fn], getVariant, this._getFns())
    const aggErr = aggs.find((a) => a.error)
    if (aggErr) return aggErr.error
    const hdr = scalarNumber(arg(idx.hdr), null)
    const rdepthArg = scalarNumber(arg(idx.rdepth), null), cdepthArg = pivot ? scalarNumber(arg(idx.cdepth), null) : 0
    const rsort = numberList(arg(idx.rsort)), csort = pivot ? numberList(arg(idx.csort)) : []
    const relationship = pivot ? 0 : scalarNumber(arg(idx.relationship), 0)
    const rel = pivot ? scalarNumber(arg(idx.rel), 0) : 0
    for (const x of [hdr, rdepthArg, cdepthArg, rsort, csort, relationship, rel]) if (x?.isError?.()) return x
    if (hdr != null && ![0, 1, 2, 3].includes(hdr)) return err('#VALUE!')
    const filterG = arg(idx.filter) != null ? toGrid(arg(idx.filter)) : null
    if (filterG?.isError?.()) return filterG

    // field_headers: missing = automatic (headers if the first value is text and the second a number).
    let hasHeaders = hdr === 1 || hdr === 3, showHeaders = hdr === 2 || hdr === 3
    if (hdr == null) {
      const a = valG.at(0, 0), b = valG.rows > 1 ? valG.at(1, 0) : null
      hasHeaders = !!(a?.isString?.() && b?.isNumber?.())
      showHeaders = hasHeaders
    }
    const start = hasHeaders ? 1 : 0
    const headerText = (g, c, generated) => (hasHeaders ? g.at(0, c) : StringValueObject.create(generated))

    const records = []
    for (let r = start; r < rowG.rows; r++) {
      if (filterG) { const f = filterG.at(r - (filterG.rows === rowG.rows ? 0 : start), 0); if (f.isError() || !f.getValue()) continue }
      records.push({
        rk: Array.from({ length: rowG.cols }, (_, c) => rowG.at(r, c)),
        ck: Array.from({ length: colG.cols }, (_, c) => colG.at(r, c)),
        v: Array.from({ length: valG.cols }, (_, c) => valG.at(r, c)),
      })
    }
    const rLevels = rowG.cols, cLevels = colG.cols
    // Missing depth = automatic: grand totals, plus subtotals where there is more than one level.
    const rdepth = rdepthArg ?? (rLevels > 1 ? 2 : 1)
    const cdepth = pivot ? (cdepthArg ?? (cLevels > 1 ? 2 : 1)) : 0
    const outCols = []
    for (let vc = 0; vc < valG.cols; vc++) for (const a of aggs) outCols.push({ vc, a })

    const aggregate = (recs, vc, a, allRecs) => {
      if (recs.length === 0) return BLANK()
      const res = a.run(column(recs.map((x) => x.v[vc])), column(allRecs.map((x) => x.v[vc])))
      return scalarResult(res)
    }
    const sortValue = (recs, i) => { const oc = outCols[i] ?? outCols[0]; return aggregate(recs, oc.vc, oc.a, records) }
    const rowSlots = buildAxis(records, (x) => x.rk, rLevels, rdepth, relationship, rsort, sortValue)
    const colSlots = pivot ? buildAxis(records, (x) => x.ck, cLevels, cdepth, 0, csort, sortValue)
      : [{ keys: [], recs: records, kind: 'detail', parent: records }]

    // relative_to (PIVOTBY): 0 column totals, 1 row totals, 2 grand total, 3 parent column total, 4 parent row total.
    const relativeSet = (rs, cs) => {
      const inCol = (x) => cs.recs.includes(x), inRow = (x) => rs.recs.includes(x)
      switch (rel) {
        case 1: return rs.recs
        case 2: return records
        case 3: return rs.recs.length ? records.filter((x) => inRow(x) && cs.parent.includes(x)) : []
        case 4: return records.filter((x) => inCol(x) && rs.parent.includes(x))
        default: return cs.recs
      }
    }
    const out = []
    const multiFn = aggs.length > 1, multiVal = valG.cols > 1
    // Header rows: column keys (one row per column level) and a row naming value/function columns.
    if (pivot) for (let lv = 0; lv < cLevels; lv++) {
      const row = Array(rLevels).fill(null).map(BLANK)
      if (showHeaders && lv === cLevels - 1) for (let c = 0; c < rLevels; c++) row[c] = headerText(rowG, c, `Row Field ${c + 1}`)
      for (const cs of colSlots) for (let i = 0; i < outCols.length; i++) row.push(i === 0 ? (cs.keys[lv] ?? BLANK()) : BLANK())
      out.push(row)
    }
    if (multiFn || multiVal || (!pivot && showHeaders)) {
      const row = Array(rLevels).fill(null).map((_, c) => (!pivot && showHeaders ? headerText(rowG, c, `Row Field ${c + 1}`) : BLANK()))
      for (const _cs of colSlots) for (const oc of outCols) {
        const vh = showHeaders || hasHeaders ? headerText(valG, oc.vc, `Value ${oc.vc + 1}`) : null
        const label = multiFn ? (vh && multiVal ? `${vh.getValue()} - ${oc.a.name}` : oc.a.name) : vh ? vh.getValue() : `Value ${oc.vc + 1}`
        row.push(StringValueObject.create(String(label)))
      }
      out.push(row)
    }
    for (const rs of rowSlots) {
      const row = [...rs.keys]
      for (const cs of colSlots) {
        const recs = rs.recs.filter((x) => cs.recs.includes(x))
        for (const oc of outCols) row.push(aggregate(recs, oc.vc, oc.a, relativeSet(rs, cs)))
      }
      out.push(row)
    }
    if (out.length === 0) return err('#CALC!')
    const width = Math.max(...out.map((r) => r.length))
    for (const r of out) while (r.length < width) r.push(BLANK())
    return ArrayValueObject.create({ calculateValueList: out, rowCount: out.length, columnCount: width,
      unitId: this.unitId, sheetId: this.subUnitId, row: this.row, column: this.column })
  }
}

// [Class, name] pairs for the engine's `formula.function` config, which registers them after the
// built-ins so GROUPBY replaces Univer's version.
export function pivotFunctions(getApi) {
  PivotFunction.getFns = () => getApi()._injector.get(IFunctionService)
  return {
    functions: [[PivotFunction, 'GROUPBY'], [PivotFunction, 'PIVOTBY']],
    descriptions: [{ functionName: 'PIVOTBY', functionType: FunctionType.Lookup, abstract: 'PIVOTBY', functionParameter: [],
      description: 'Creates a summary of your data via a formula, grouping along rows and columns.' }],
  }
}
