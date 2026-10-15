// Excel functions Univer's engine lacks, implemented as native engine functions so they see
// value objects (errors, omitted arguments, references) exactly as built-ins do.
import {
  ArrayValueObject, BaseFunction, BooleanValueObject, ErrorValueObject, FunctionType, IFunctionService,
  NullValueObject, NumberValueObject, StringValueObject,
} from '@univerjs/engine-formula'

const err = (e) => ErrorValueObject.create(e)
const first = (v) => {
  if (v == null) return v
  if (v.isReferenceObject?.()) return v.getFirstCell()
  if (v.isArray?.()) return v.get(0, 0) ?? NullValueObject.create()
  return v
}
// Excel's numeric coercion for a scalar argument: numbers, numeric text, booleans; blank -> 0.
const num = (v) => {
  const x = first(v)
  if (x == null || x.isNull()) return 0
  if (x.isError()) return x
  if (x.isNumber()) return +x.getValue()
  if (x.isBoolean()) return x.getValue() ? 1 : 0
  const n = Number(String(x.getValue()).trim())
  return String(x.getValue()).trim() !== '' && Number.isFinite(n) ? n : err('#VALUE!')
}
const isErr = (x) => typeof x === 'object' && x?.isError?.()

class Fn extends BaseFunction {
  constructor(name, min, max, impl, flags = {}) {
    super(name)
    this.minParams = min
    this.maxParams = max
    this._impl = impl
    Object.assign(this, flags)
  }
  calculate(...args) { return this._impl.apply(this, args) }
}

export function installExcelFunctions(api, env = {}) {
  const fns = api._injector.get(IFunctionService)
  const call = (name, ...args) => fns.getExecutor(name).calculate(...args)

  const defs = [
    // ISO.CEILING is defined by Excel as identical to CEILING.PRECISE.
    ['ISO.CEILING', FunctionType.Math, 1, 2, (n, s) => call('CEILING.PRECISE', n, s ?? NumberValueObject.create(1)),
      'Rounds a number up to the nearest integer or multiple of significance, regardless of sign.'],

    // Without furigana data (Sheetly stores none) Excel returns the cell's own text.
    ['PHONETIC', FunctionType.Text, 1, 1, function (ref) {
      const x = first(ref)
      if (x.isError()) return x
      return StringValueObject.create(x.isNull() ? '' : String(x.getValue()))
    }, 'Extracts the phonetic (furigana) characters from a text string.', { needsReferenceObject: true }],

    // Sheetly has no PivotTables, and Excel returns #REF! when pivot_table is not one.
    ['GETPIVOTDATA', FunctionType.Lookup, 2, 254, () => err('#REF!'),
      'Returns data stored in a PivotTable report.'],

    ['AMORDEGRC', FunctionType.Financial, 6, 7, function (cost, purchased, firstPeriod, salvage, period, rate, basis) {
      const a = [cost, purchased, firstPeriod, salvage, period, rate, basis ?? NumberValueObject.create(0)].map(num)
      const bad = a.find(isErr)
      if (bad) return bad
      let [c, d0, d1, s, per, r, b] = a
      per = Math.trunc(per); b = Math.trunc(b)
      if (r <= 0 || per < 0 || s < 0 || s > c || d1 < d0 || ![0, 1, 3, 4].includes(b)) return err('#NUM!')
      const life = 1 / r
      if ((life > 0 && life < 1) || (life > 1 && life < 2) || (life > 2 && life < 3) || (life > 4 && life < 5)) return err('#NUM!')
      const coeff = life < 3 ? 1 : life < 5 ? 1.5 : life <= 6 ? 2 : 2.5
      r *= coeff
      const yf = call('YEARFRAC', NumberValueObject.create(d0), NumberValueObject.create(d1), NumberValueObject.create(b))
      if (yf.isError()) return yf
      let rate1 = Math.round(+yf.getValue() * r * c)
      c -= rate1
      let rest = c - s
      for (let n = 0; n < per; n++) {
        rate1 = Math.round(r * c)
        rest -= rate1
        if (rest < 0) return NumberValueObject.create(per - n <= 1 ? Math.round(c * 0.5) : 0)
        c -= rate1
      }
      return NumberValueObject.create(rate1)
    }, 'Returns the depreciation for each accounting period by using a depreciation coefficient.'],

    ['INFO', FunctionType.Information, 1, 1, function (type) {
      const t = first(type)
      if (t.isError()) return t
      const key = String(t.getValue()).toLowerCase()
      const wb = api.getActiveWorkbook(), ws = wb?.getActiveSheet()
      switch (key) {
        case 'directory': return StringValueObject.create(env.directory?.() ?? '/')
        case 'numfile': return NumberValueObject.create(wb ? wb.getSheets().length : 0)
        case 'origin': {
          const st = ws?.getScrollState?.() ?? {}
          const col = colName(st.sheetViewStartColumn ?? 0), row = (st.sheetViewStartRow ?? 0) + 1
          return StringValueObject.create(`$A:$${col}$${row}`)
        }
        case 'osversion': return StringValueObject.create(env.osversion?.() ?? 'Linux')
        case 'recalc': return StringValueObject.create('Automatic')
        case 'release': return StringValueObject.create('16.0')
        // Excel documents only "pcdos" (Windows) and "mac"; Sheetly reports the Windows value.
        case 'system': return StringValueObject.create('pcdos')
        case 'memavail': case 'memused': case 'totmem': return err('#N/A')
        default: return err('#VALUE!')
      }
    }, 'Returns information about the current operating environment.'],

    // An omitted LAMBDA argument reaches the function as a bare null value; a reference to an
    // empty cell arrives as a reference object, which is not "omitted".
    ['ISOMITTED', FunctionType.Information, 1, 1, (v) => BooleanValueObject.create(v == null || (!v.isReferenceObject?.() && v.isNull())),
      'Checks whether the value in a LAMBDA is missing.', { needsReferenceObject: true }],

    ['TRIMRANGE', FunctionType.Lookup, 1, 3, function (ref, rowsArg, colsArg) {
      let arr = ref
      if (arr.isReferenceObject?.()) arr = arr.toArrayValueObject()
      if (arr.isError()) return arr
      if (!arr.isArray()) return arr
      const tr = rowsArg == null || first(rowsArg).isNull() ? 3 : num(rowsArg)
      const tc = colsArg == null || first(colsArg).isNull() ? 3 : num(colsArg)
      if (isErr(tr)) return tr
      if (isErr(tc)) return tc
      if (![0, 1, 2, 3].includes(tr) || ![0, 1, 2, 3].includes(tc)) return err('#VALUE!')
      const rows = arr.getRowCount(), cols = arr.getColumnCount()
      const blank = (r, c) => { const x = arr.get(r, c); return x == null || x.isNull() }
      const rowBlank = (r) => { for (let c = 0; c < cols; c++) if (!blank(r, c)) return false; return true }
      const colBlank = (c) => { for (let r = 0; r < rows; r++) if (!blank(r, c)) return false; return true }
      let r0 = 0, r1 = rows - 1, c0 = 0, c1 = cols - 1
      if (tr & 1) while (r0 <= r1 && rowBlank(r0)) r0++
      if (tr & 2) while (r1 >= r0 && rowBlank(r1)) r1--
      if (tc & 1) while (c0 <= c1 && colBlank(c0)) c0++
      if (tc & 2) while (c1 >= c0 && colBlank(c1)) c1--
      if (r0 > r1 || c0 > c1) return err('#REF!')
      const out = []
      for (let r = r0; r <= r1; r++) { const row = []; for (let c = c0; c <= c1; c++) row.push(arr.get(r, c) ?? NullValueObject.create()); out.push(row) }
      return ArrayValueObject.create({ calculateValueList: out, rowCount: out.length, columnCount: out[0].length,
        unitId: this.unitId, sheetId: this.subUnitId, row: this.row, column: this.column })
    }, 'Scans in from the edges of a range or array until it finds a non-blank cell or value, then excludes those blank rows or columns.', { needsReferenceObject: true }],

    // Linked data types (Stocks, Geography) do not exist in Sheetly; Excel returns #FIELD! for
    // a value that is not a data type.
    ['FIELDVALUE', FunctionType.Lookup, 2, 2, (v) => (first(v)?.isError() ? first(v) : err('#FIELD!')),
      'Returns a field from a linked data type.'],
  ]
  // Cube functions need an OLAP / Data Model connection. Sheetly has none, so every connection
  // name is invalid, for which Excel returns #NAME?.
  for (const n of ['CUBEKPIMEMBER', 'CUBEMEMBER', 'CUBEMEMBERPROPERTY', 'CUBERANKEDMEMBER', 'CUBESET', 'CUBESETCOUNT', 'CUBEVALUE'])
    defs.push([n, FunctionType.Cube ?? FunctionType.Lookup, 1, 255, () => err('#NAME?'), `Cube function ${n} (requires a Data Model connection).`])

  for (const [name, type, min, max, impl, description, flags] of defs) {
    fns.registerExecutors(new Fn(name, min, max, impl, flags))
    fns.registerDescriptions({ functionName: name, functionType: type, description, abstract: description,
      functionParameter: [] })
  }
}

function colName(c) {
  let s = ''
  for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}
