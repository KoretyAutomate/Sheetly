// PivotTables in an .xlsx: each report is saved as Excel saves one, a PivotCache (definition and records, the data
// the report was built from) plus a PivotTable part on its sheet; the cells keep the report's values too. The cache
// asks Excel to refresh on open, so Excel lays the report out itself. Reading takes back the reports Sheetly can show
// (worksheet sources, compact layouts of plain fields; anything else stays values).
import { FUNCS, cacheItems, isErrorValue, resolveItem, valueName } from './pivot-model.js'
import { colName, parseRange, unescXml } from './xlsx-objects.js'

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
// (everything here is an attribute: line breaks and tabs are written as references, or XML readers make them spaces)
const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  .replace(/\n/g, '&#10;').replace(/\r/g, '&#13;').replace(/\t/g, '&#9;')
const unesc = unescXml
const ref = (g) => `${colName(g.startColumn)}${g.startRow + 1}:${colName(g.endColumn)}${g.endRow + 1}`
const isBlank = (x) => x == null || x === ''

// A cache value as Excel writes it: missing, number, logical, error or text.
const valueXml = (x) => (isBlank(x) ? '<m/>' : typeof x === 'number' ? `<n v="${x}"/>` : typeof x === 'boolean' ? `<b v="${x ? 1 : 0}"/>`
  : isErrorValue(x) ? `<e v="${esc(x.toUpperCase())}"/>` : `<s v="${esc(x)}"/>`)

// What a field holds, as the sharedItems flags describe it ("semi-mixed" means it has text or blanks).
function sharedItemsXml(values, items) {
  const set = values.filter((x) => !isBlank(x))
  const blank = set.length < values.length
  const str = set.some((x) => typeof x === 'string' && !isErrorValue(x)), nums = set.filter((x) => typeof x === 'number')
  const kinds = [str, nums.length > 0, set.some((x) => typeof x === 'boolean'), set.some(isErrorValue)].filter(Boolean).length
  const a = []
  if (!str && !blank) a.push('containsSemiMixedTypes="0"')
  if (!str) a.push('containsString="0"')
  if (blank) a.push('containsBlank="1"')
  if (kinds > 1) a.push('containsMixedTypes="1"')
  if (nums.length) {
    a.push('containsNumber="1"')
    if (nums.every(Number.isInteger)) a.push('containsInteger="1"')
    a.push(`minValue="${nums.reduce((m, x) => Math.min(m, x))}"`, `maxValue="${nums.reduce((m, x) => Math.max(m, x))}"`)
  }
  if (!items) return `<sharedItems${a.map((x) => ` ${x}`).join('')}/>`
  return `<sharedItems${a.map((x) => ` ${x}`).join('')} count="${items.length}">${items.map(valueXml).join('')}</sharedItems>`
}

/**
 * The parts of one report. spec: { name, cache: {fields, records}, rows, cols, values, filters, source: {sheetName,
 * range} | {table}, location: range, firstDataRow, firstDataCol }.
 */
export function pivotXml(spec, cacheId) {
  const { fields, records } = spec.cache
  const axis = new Set([...spec.rows, ...spec.cols, ...spec.filters.map((f) => f.field)])
  // Fields on the report, and fields with anything but numbers, list their items; plain number fields are inline.
  const listed = fields.map((f, i) => axis.has(f) || records.some((r) => !isBlank(r[i]) && typeof r[i] !== 'number'))
  const items = fields.map((f, i) => (listed[i] ? cacheItems(spec.cache, f) : null))
  const source = spec.source.table ? `<worksheetSource name="${esc(spec.source.table)}"/>` : `<worksheetSource ref="${ref(spec.source.range)}" sheet="${esc(spec.source.sheetName)}"/>`
  const cache = `${HEAD}<pivotCacheDefinition xmlns="${MAIN}" xmlns:r="${R}" r:id="rId1" refreshOnLoad="1" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="${records.length}">`
    + `<cacheSource type="worksheet">${source}</cacheSource><cacheFields count="${fields.length}">`
    + fields.map((f, i) => `<cacheField name="${esc(f)}" numFmtId="0">${sharedItemsXml(records.map((r) => r[i]), items[i]?.items)}</cacheField>`).join('')
    + '</cacheFields></pivotCacheDefinition>'
  const recs = `${HEAD}<pivotCacheRecords xmlns="${MAIN}" xmlns:r="${R}" count="${records.length}">`
    + records.map((r) => `<r>${r.map((x, i) => (items[i] ? `<x v="${items[i].indexOf(x)}"/>` : valueXml(x))).join('')}</r>`).join('')
    + '</pivotCacheRecords>'
  const at = (f) => fields.indexOf(f)
  // A chosen filter item as the report matches it (the number 2024 for "2024"), as an index into the field's items.
  const itemIndex = (f, x) => { const it = resolveItem(spec.cache, records, f, x); return it == null ? null : items[at(f)].indexOf(it) ?? null }
  const filterOf = new Map(spec.filters.map((f) => [f.field, f]))
  const valueFields = new Set(spec.values.map((v) => v.field))
  const pivotFields = fields.map((f, i) => {
    const kind = spec.rows.includes(f) ? 'axisRow' : spec.cols.includes(f) ? 'axisCol' : filterOf.has(f) ? 'axisPage' : null
    const data = valueFields.has(f) ? ' dataField="1"' : ''
    if (!kind) return `<pivotField${data} showAll="0"/>`
    // A report filter showing some items hides the others.
    const chosen = filterOf.get(f)?.items
    const shown = chosen ? new Set(chosen.map((x) => itemIndex(f, x)).filter((k) => k != null)) : null
    const multi = chosen && chosen.length > 1 ? ' multipleItemSelectionAllowed="1"' : ''
    const list = items[i].items.map((_, k) => `<item x="${k}"${shown && chosen.length > 1 && !shown.has(k) ? ' h="1"' : ''}/>`).join('')
    return `<pivotField axis="${kind}"${data}${multi} showAll="0"><items count="${items[i].items.length + 1}">${list}<item t="default"/></items></pivotField>`
  }).join('')
  const fieldList = (tag, list) => (list.length ? `<${tag} count="${list.length}">${list.map((x) => `<field x="${x}"/>`).join('')}</${tag}>` : '')
  // Several value fields make a "Values" column field (-2), innermost.
  const colFields = [...spec.cols.map(at), ...(spec.values.length > 1 ? [-2] : [])]
  const pages = spec.filters.map((f) => {
    const one = f.items?.length === 1 ? itemIndex(f.field, f.items[0]) : null
    return `<pageField fld="${at(f.field)}"${one != null ? ` item="${one}"` : ''} hier="-1"/>`
  })
  const dataFields = spec.values.map((v) => `<dataField name="${esc(valueName(v))}" fld="${at(v.field)}"${v.fn && v.fn !== 'sum' ? ` subtotal="${v.fn}"` : ''} baseField="0" baseItem="0"/>`)
  const pageCount = spec.filters.length ? ` rowPageCount="${spec.filters.length}" colPageCount="1"` : ''
  const table = `${HEAD}<pivotTableDefinition xmlns="${MAIN}" name="${esc(spec.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0"`
    + ' applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">'
    + `<location ref="${ref(spec.location)}" firstHeaderRow="1" firstDataRow="${spec.firstDataRow}" firstDataCol="${spec.firstDataCol}"${pageCount}/>`
    + `<pivotFields count="${fields.length}">${pivotFields}</pivotFields>`
    + fieldList('rowFields', spec.rows.map(at)) + fieldList('colFields', colFields)
    + (pages.length ? `<pageFields count="${pages.length}">${pages.join('')}</pageFields>` : '')
    + (dataFields.length ? `<dataFields count="${dataFields.length}">${dataFields.join('')}</dataFields>` : '')
    + '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>'
  return { cache, records: recs, table }
}

// ---- reading
const attrs = (s) => Object.fromEntries([...String(s ?? '').matchAll(/([\w:]+)="([^"]*)"/g)].map((m) => [m[1], unesc(m[2])]))
const DAY = 86400000
// One cache value; dates (d) become serial numbers of the workbook's date system.
function valueIn(tag, a, epoch) {
  if (tag === 'm') return null
  if (tag === 'n') return Number(a.v)
  if (tag === 'b') return a.v === '1' || a.v === 'true'
  if (tag === 'd') { const t = Date.parse(/[zZ]|[+-]\d\d:?\d\d$/.test(a.v) ? a.v : `${a.v}Z`); return Number.isNaN(t) ? null : (t - epoch) / DAY }
  return a.v ?? ''
}
const VALUE_TAGS = /<(m|n|b|e|s|d|x)\b([^>]*?)\/?>(?:<\/\1>)?/g

/**
 * A report Sheetly can show, from its parts' XML, or null: { name, location, filterCount, rows, cols, values, filters,
 * cache, source: { ref, sheet } | { table } }. Sources other than a worksheet range or table, grouped or calculated
 * fields, values shown as (% of, running total ...), items hidden on rows or columns, report filters laid out across
 * and caches saved without their records are not supported.
 */
export function readPivot(tableXml, cacheXml, recordsXml, { date1904 = false } = {}) {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)
  const src = /<cacheSource\b([^>]*)>([\s\S]*?)<\/cacheSource>/.exec(cacheXml)
  if (!src || (attrs(src[1]).type ?? 'worksheet') !== 'worksheet') return null
  // Without its records (saveData="0") the cache is not the data the report shows.
  if (recordsXml == null || attrs(/<pivotCacheDefinition\b([^>]*)>/.exec(cacheXml)?.[1]).saveData === '0') return null
  const ws = attrs(/<worksheetSource\b([^>]*)\/?>/.exec(src[2])?.[1])
  const source = ws.name ? { table: ws.name } : ws.ref && parseRange(ws.ref) ? { range: parseRange(ws.ref), sheet: ws.sheet ?? null } : null
  if (!source) return null
  const fieldsXml = [...cacheXml.matchAll(/<cacheField\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cacheField>)/g)]
  if (fieldsXml.some((m) => attrs(m[1]).formula != null || /<fieldGroup\b/.test(m[2] ?? ''))) return null
  const fields = fieldsXml.map((m) => attrs(m[1]).name ?? '')
  const shared = fieldsXml.map((m) => [...(/<sharedItems\b[^>]*>([\s\S]*?)<\/sharedItems>/.exec(m[2] ?? '')?.[1] ?? '').matchAll(VALUE_TAGS)].map((x) => valueIn(x[1], attrs(x[2]), epoch)))
  const records = [...(recordsXml ?? '').matchAll(/<r>([\s\S]*?)<\/r>|<r\/>/g)].map((m) => [...(m[1] ?? '').matchAll(VALUE_TAGS)]
    .map((x, i) => (x[1] === 'x' ? shared[i]?.[Number(attrs(x[2]).v)] ?? null : valueIn(x[1], attrs(x[2]), epoch))))
  if (records.some((r) => r.length !== fields.length)) return null
  const t = attrs(/<pivotTableDefinition\b([^>]*)>/.exec(tableXml)?.[1])
  const loc = attrs(/<location\b([^>]*)\/?>/.exec(tableXml)?.[1])
  const location = parseRange(loc.ref ?? '')
  if (!location) return null
  // Report filters one per row above the report (Excel's default "Down, Then Over"), else not read.
  if (t.pageOverThenDown === '1' || Number(t.pageWrap ?? 0) > 0 || Number(loc.colPageCount ?? 1) > 1) return null
  // Each pivot field's items: the shared item each shows, and whether it is hidden.
  const pf = [...(/<pivotFields\b[^>]*>([\s\S]*?)<\/pivotFields>/.exec(tableXml)?.[1] ?? '').matchAll(/<pivotField\b([^>]*?)(?:\/>|>([\s\S]*?)<\/pivotField>)/g)]
    .map((m) => ({ a: attrs(m[1]), items: [...(m[2] ?? '').matchAll(/<item\b([^>]*?)\/?>/g)].map((x) => attrs(x[1])).filter((x) => x.x != null && (!x.t || x.t === 'data')) }))
  const list = (tag, inner) => [...(new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(tableXml)?.[1] ?? '').matchAll(new RegExp(`<${inner}\\b([^>]*?)/?>`, 'g'))].map((m) => attrs(m[1]))
  const rows = list('rowFields', 'field').map((f) => Number(f.x)).filter((x) => x >= 0)
  const cols = list('colFields', 'field').map((f) => Number(f.x)).filter((x) => x >= 0)
  if ([...rows, ...cols].some((x) => !fields[x] || pf[x]?.items.some((i) => i.h === '1' || i.h === 'true'))) return null
  const itemOf = (fi, k) => { const x = pf[fi]?.items[k]?.x; const v = x != null ? shared[fi]?.[Number(x)] : undefined; return v === undefined ? undefined : isBlank(v) ? '(blank)' : v }
  const filters = []
  for (const p of list('pageFields', 'pageField')) {
    const fi = Number(p.fld)
    if (!fields[fi]) return null
    let items = null
    if (p.item != null) items = [itemOf(fi, Number(p.item))]
    else if (pf[fi]?.items.some((i) => i.h === '1' || i.h === 'true')) items = pf[fi].items.map((i, k) => (i.h === '1' || i.h === 'true' ? undefined : itemOf(fi, k)))
    if (items?.some((x) => x === undefined && p.item != null)) return null
    filters.push({ field: fields[fi], items: items ? items.filter((x) => x !== undefined) : null })
  }
  const values = []
  for (const d of list('dataFields', 'dataField')) {
    const fi = Number(d.fld), fn = d.subtotal ?? 'sum'
    if (!fields[fi] || !FUNCS.includes(fn) || (d.showDataAs && d.showDataAs !== 'normal')) return null
    const v = { field: fields[fi], fn }
    // A caption Excel made (with the trailing space it adds when it would equal a field name) is the default one.
    if (d.name && d.name.trimEnd() !== valueName(v)) v.name = d.name
    values.push(v)
  }
  return { name: t.name ?? 'PivotTable1', location, filterCount: filters.length, rows: rows.map((x) => fields[x]), cols: cols.map((x) => fields[x]), values, filters,
    cache: { fields, records }, source }
}
