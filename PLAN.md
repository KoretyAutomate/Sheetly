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

## Next: outline (grouping) + ribbon KeyTips (2026-10-16)
Outline — src/outline.js
- Storage: level per row/column in Univer's row/col data `custom.outlineLevel` (1..7), collapsed flag on the
  summary row/col `custom.outlineCollapsed`. Lives in rowData, so insert/delete rows shift it, wb.save() keeps it.
- Writes via mutations only (sheet.mutation.set-row-data / set-col-data, set-row-hidden / visible,
  set-col-hidden / visible), one undo step per user action (IUndoRedoService.pushUndoRedo).
- Excel semantics: summary rows below / columns right; group = maximal run with level >= k; +/- button on the
  summary row; collapse hides the run; expand re-shows it except nested collapsed groups; level buttons 1..n+1
  collapse every group of level >= k; ungroup decrements and leaves hidden rows hidden.
- Keys: Alt+Shift+Right group, Alt+Shift+Left ungroup (whole rows -> rows, whole columns -> columns, otherwise
  Rows/Columns dialog, default Rows), Ctrl+8 show/hide outline symbols. Also Hide/Show Detail, Clear Outline.
- Rendering: header extensions (registerSheetRowHeaderExtension / ColumnHeader): widen the header by a gutter
  (setRowHeaderWidth / setColumnHeaderHeight), draw brackets and +/- in the gutter, redraw the numbers/letters
  in the remaining part (default header text made transparent while a gutter is shown). Level buttons are a DOM
  box over the fixed corner. Gutter clicks: capture-phase pointerdown on the canvas, row/col from the skeleton.
- .xlsx: row/column outlineLevel, hidden, collapsed round-trip via ExcelJS.
KeyTips — src/keytips.js
- Alt (press+release) or F10 shows tab KeyTips; Alt+letter goes straight into the tab; letters walk the tree;
  multi-letter tips filter as typed; Esc goes back one level; Alt/F10/click exits. Not active while editing.
- Panel: an overlay ribbon strip (tabs, then the current level's commands with their KeyTip badges); clickable.
- Tree data in one table of [keys, label, action | children]; sequences taken from Excel 365 (sources
  cross-checked), plus the Office 2003 access keys Excel still accepts (Alt E S, Alt O C A, ...).
- Dialog access keys (letters select options in Paste Special etc.) so Alt+E+S+V / Alt+H+V+S work end to end.
- Electron: renderer preventDefault()s Alt keyup so the window menu does not take focus.
Tests: shortcuts_full.mjs cases for group/ungroup/collapse/expand/levels/Ctrl+8/undo/insert-row shift and for
KeyTip sequences (end effect asserted, e.g. Alt H O I autofits, Alt H V V pastes values); xlsx round-trip test.
Delegation: KeyTip research ran as a Claude subagent; LocalCrew not used (broken since 2026-10-15, user said skip).
Review (plan agent) adopted: inserted rows join the group (command interceptor), fixed base header size, active
cell leaves collapsed rows, Univer unhide arrows suppressed, multi-selection / 8-level refusals, exceljs patched for
the collapsed flag and cell-less rows, KeyTip listener registered before Univer's, no forceDisable nesting.
Not done: summary rows above / columns left (outlinePr) are not read; Ctrl+8 state is not saved; Auto Outline;
Alt+Shift+= / Alt+Shift+- (listed on Microsoft's outline page, not verified) are not bound.
KeyTip letters with single-source support (to check against a live Excel): H 2, H 3 U, H A T/M/B, H P, H O A/D/T/E,
H U A/C/M/I/F, H F I *, H S F/C/Y, H F D R/G/S/U/M/N, H F Q *, H V N, N I, N C1 (opens a clustered column chart
directly), A C, A F F, legacy E D/E L/E A */E I */O C W/O R */D G */D L/D E/V Z. Sheetly lists only commands it has.

Status 2026-10-16: outline + KeyTips done. Shortcuts 176/176, outline 27/27, formulas 398/398, coverage 517/517,
flash fill 11/11. Also fixed: a column-width change froze the app for >10 s (Univer queued auto-height for all
1,048,576 rows; patches/@univerjs+sheets-ui clips the task to the last row with cells or row data).

Known gaps (not aligned yet):
- F6 pane cycling, Ctrl+F1 ribbon toggle, Ctrl+Q Quick Analysis, Shift+F8 add-to-selection.
- KeyTips cover only commands Sheetly has; single-source letters listed above still need checking in real Excel.
- TRANSLATE / DETECTLANGUAGE / STOCKHISTORY return #CONNECT! (they need Microsoft online services);
  CUBE* return #NAME? (no Data Model connection); GETPIVOTDATA returns #REF! (no PivotTables).
- Charts: clustered column only (F11 / Alt+F1), not saved to .xlsx. PivotTables not available.
- .xlsx styles beyond bold/italic/number format/widths/merges are not round-tripped.
- Expected values in the formula suite were written from Excel's documentation, not produced by Excel.

## Next: remaining gaps (2026-10-16, round 2)
A. .xlsx fidelity (src/xlsx.js) — resolve style ids (snapshot `styles` map; cell.s may be a string id) both ways:
   font bold/italic/underline/strike/size/name/color, fill (solid fgColor), borders (t/b/l/r style + color),
   alignment (horizontal, vertical, wrapText, textRotation incl. vertical text 255), number format.
   Row heights (pt <-> px * 4/3, customHeight only), frozen panes (views[0] state 'frozen' <-> config.freeze),
   tab colour, hidden sheets, shared formulas (cell.formula gives the translated text), dates read as Excel
   serials (not ISO strings), formula error results (#N/A etc.), defined names (workbook-level), notes.
   Test: test/xlsx.test.mjs in Node (ExcelJS file written by Sheetly's writer, read back by ExcelJS, and a
   workbook built by ExcelJS read through xlsxToSnapshot) + one in-app round-trip of a UI-styled cell.
B. Shortcuts
   - F6 / Shift+F6 cycle focus: worksheet -> ribbon (toolbar) -> sheet tabs bar -> worksheet (Excel: worksheet,
     ribbon, task pane, status bar/zoom). Window capture listener so it works wherever focus is; Esc from the
     ribbon/tabs returns to the grid as in Excel. Not while a .sx-overlay dialog is open.
   - Ctrl+F1 collapse/expand the ribbon (toolbar hidden, tabs kept); state kept in localStorage.
   - Ctrl+Q Quick Analysis: popup at the selection's bottom-right with tabs Formatting (Data Bars, Color Scale,
     Icon Set, Greater Than, Top 10%, Clear Format), Charts (Clustered Column), Totals (Sum, Average, Count,
     % Total, Running Total, rows below / columns right), Tables (Table). Left/Right switch tabs, Tab/arrows move
     between items, Enter runs, Esc closes. Conditional formats via Univer's CF facade.
   - Shift+F8 Add to Selection mode: status bar "Add to Selection"; keeps the selections present when entered
     (and those committed by mouse clicks) while arrows/Shift+arrows/clicks make a new area; Shift+F8 or Esc ends.
C. Outline
   - Settings dialog (Data > Outline launcher, KeyTip Alt A J? -> use the Group menu "Settings"): "Summary rows
     below detail", "Summary columns to right of detail", buttons Create (Auto Outline), OK, Cancel.
   - Summary direction per sheet in worksheet `custom.sheetlyOutline = { summaryBelow, summaryRight,
     hideSymbols }`; the summary of a group is g.end+1 (below/right) or g.start-1 (above/left) everywhere
     (collapse bit, +/- box, bracket tick, active-cell move, Show/Hide Detail, xlsx markCollapsed).
   - Ctrl+8 state saved in the same custom field and in .xlsx sheetView showOutlineSymbols (ExcelJS patch).
   - xlsx: sheetPr/outlinePr summaryBelow/summaryRight round-trip (ExcelJS supports outlineProperties).
   - Auto Outline (Data > Group > Auto Outline, Alt A G A): rows whose formulas reference a vertical range in
     the same column directly adjacent (above when summary below) become summaries; nested ranges nest levels;
     same for columns with horizontal ranges. Replaces an existing outline after "Modify existing outline?"; no
     such formulas -> "Cannot create an outline."
Not in this round: charts in .xlsx (ExcelJS cannot write charts), online functions, PivotTables,
Alt+Shift+= / Alt+Shift+- (unverified).
Review (plan agent, 2026-10-16) adopted: style ids resolved (cell/row/column styles, merged col<row<cell on write),
Univer's clockwise text angle (angle up = -45; KeyTips Alt H FQ and Format Cells degrees fixed too), colour forms
#rgb / rgb() / theme+tint / indexed, blank styled cells kept, filled-down {si} formulas resolved from the live
workbook on save, _xlfn./_xlws. stripped on load and added on save, exact error list, very hidden sheets, scrolled
freeze (topLeftCell), notes get a size, format 14 -> m/d/yyyy, auto-fit rows stay automatic (ExcelJS patched for
customHeight), defined names written/read raw (sheet scope, constants, formulas; ExcelJS patched), collapse bits
move when the summary direction changes, Quick Analysis text options + Sparklines tab, Auto Outline over
subtotals. Also found by tests: a plain arrow with a range selected moved from the range edge (Univer); it now
moves from the active cell as in Excel.

Status 2026-10-16 (round 2): done. xlsx 56/56, round2 43/43, outline 27/27, formulas 398/398, coverage 517/517.
Not done from this round (deliberate or unverified):
- F6 on the ribbon does not show KeyTip badges (Excel does, per the reviewer; unverified).
- Quick Analysis: no live preview; Charts tab is fixed (Clustered Column), not recommendations; Sparklines and
  PivotTable unavailable; Totals > Count uses COUNTA (unverified against Excel).
- Outline Settings has no "Automatic styles" / "Apply Styles"; Auto Outline scans the whole sheet, not the
  selection / current region.
- .xlsx: workbook default font (Calibri / Aptos) and Excel's default row height / column width not applied;
  date1904 workbooks read as 1900 dates (values shown as the same dates); LAMBDA/LET parameter names are written
  without _xlpm.; spills saved as anchor formulas; conditional formats, data validation, tables, hyperlinks and
  threaded comments are not written.
- Opening a file given on the command line (file association) is not supported.
