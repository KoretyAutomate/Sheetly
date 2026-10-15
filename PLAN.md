# Sheetly — Excel-compatible spreadsheet for Linux (open source)

Goal: a Linux desktop spreadsheet whose shortcuts and formulas behave like MS Excel.
Not a port of Excel (its source is not open). Built on Univer (Apache-2.0) + Electron + ExcelJS.

## Decisions
- Delivery: Electron desktop app (aarch64 box: no Rust/webkit2gtk; Electron ships arm64).
- Engine: @univerjs/presets + preset-sheets-core only (Apache-2.0). Advanced preset not used (licence unclear).
- .xlsx I/O: ExcelJS <-> Univer snapshot (Univer's own xlsx I/O is paid).
- Honest limit: "exactly the same" is a target, tracked by tests, not a claim.

## Phases
1. Scaffold: Electron shell + Univer sheet renders; File > Open/Save .xlsx.
2. Formula conformance: table of Excel documented examples -> assert results; list gaps vs Excel function list.
3. Shortcut conformance: map Excel shortcut table -> verify each; fill gaps (Ctrl+Arrow, F2, F4, Ctrl+;, Alt+=, ...).
4. Packaging: AppImage/.deb (arm64 + x64), .desktop file, xlsx association.

## Verification
Launch app, drive via Playwright-for-Electron, assert formula results and shortcut effects; save to test_results/.

## Status 2026-10-15
Formulas: all 517 functions in test/excel-functions.txt are registered (coverage 517/517); value suite 398/398.
- Engine fixes kept as patches/ (patch-package, applied on postinstall): IM* complex results returned as text,
  VALUETOTEXT/DOLLAR text results, SORT no longer mutates the cached range array, eta lambdas
  (BYROW(r,SUM) etc.), Excel error values #CONNECT!/#BLOCKED!/#UNKNOWN!/#FIELD!/#BUSY! + ERROR.TYPE 9-13,
  formula autocomplete: Enter commits as typed (Tab accepts), as in Excel.
- Sheetly-native functions (src/excel-functions.js, src/groupby.js): ISO.CEILING, PHONETIC, GETPIVOTDATA,
  AMORDEGRC, INFO, ISOMITTED, TRIMRANGE, FIELDVALUE, CUBE*, WEBSERVICE, TRANSLATE, DETECTLANGUAGE,
  STOCKHISTORY, GROUPBY (replaces Univer's), PIVOTBY.
- Workbook locale is en-US (Univer default zh-CN made dates/times render in Chinese).
- Every sheet has Excel's grid: 1,048,576 rows x 16,384 columns.

Shortcuts: test/shortcuts_full.mjs drives real key presses and asserts each effect: 139/139 (see test_results/).
Flash Fill learner: test/flashfill.test.mjs 11/11.

Known gaps (not aligned yet):
- Row/column grouping & outline (Alt+Shift+Right/Left, Ctrl+8): Univer has no outline feature.
- Ribbon key tips (Alt / F10), F6 pane cycling, Ctrl+F1 ribbon toggle, Ctrl+Q Quick Analysis, Shift+F8 add-to-selection.
- TRANSLATE / DETECTLANGUAGE / STOCKHISTORY return #CONNECT! (they need Microsoft online services);
  CUBE* return #NAME? (no Data Model connection); GETPIVOTDATA returns #REF! (no PivotTables).
- Charts: clustered column only (F11 / Alt+F1), not saved to .xlsx. PivotTables not available.
- .xlsx styles beyond bold/italic/number format/widths/merges are not round-tripped.
- Expected values in the formula suite were written from Excel's documentation, not produced by Excel.
