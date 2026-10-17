// FRange.getCalculatedValue(): a cell's value as formulas see it — the stored value, or for a cell inside a dynamic
// array's spill, the calculated result there — never the formatted text that Univer's getValue() returns for a
// formatted number ("$100"). Sheetly reads cell values through this.
import { FormulaDataModel } from '@univerjs/engine-formula'
import { FRange } from '@univerjs/sheets/facade'

/** The calculated spill value at a cell of a worksheet (null when none). */
export function spillValue(injector, unitId, sheetId, r, c) {
  const data = injector.get(FormulaDataModel).getArrayFormulaCellData?.()
  return data?.[unitId]?.[sheetId]?.[r]?.[c]?.v ?? null
}

class SheetlyCalculatedValue extends FRange {
  getCalculatedValue() {
    const v = this.getRawValue()
    if (v !== null && v !== undefined) return v
    return spillValue(this._injector, this._workbook.getUnitId(), this._worksheet.getSheetId(), this._range.startRow, this._range.startColumn)
  }
}
FRange.extend(SheetlyCalculatedValue)
