# Sheetly

An open-source desktop spreadsheet for Linux whose **formulas and keyboard shortcuts follow Microsoft Excel**.
It is built on [Univer](https://github.com/dream-num/univer) (Apache-2.0), [ExcelJS](https://github.com/exceljs/exceljs)
(MIT) and [Electron](https://www.electronjs.org/). It contains no Microsoft code: Excel is not open source.
Sheetly is an independent, compatible application, and its compatibility is checked by automated tests rather
than claimed.

> Sheetly is not affiliated with or endorsed by Microsoft. "Microsoft" and "Excel" are trademarks of Microsoft
> Corporation and are used here only to describe compatibility.

## What works

- **Formulas.** All 517 functions in Excel's function list are registered, including dynamic arrays, LAMBDA
  helpers, GROUPBY and PIVOTBY, and the newer text, regex and statistics functions. 398 worked examples taken from
  Excel's documentation give the same results.
- **Shortcuts.** Excel's keyboard map, tested with real key presses (176 cases):
  - navigation and selection, including End mode, F8 Extend Selection and Shift+F8 Add to Selection
  - editing, number formats, Paste Special, Go To Special and Flash Fill (Ctrl+E)
  - Ctrl+Q Quick Analysis, F6 / Shift+F6 pane cycling and Ctrl+F1 to collapse the ribbon
- **Ribbon KeyTips.** Press and release Alt (or press F10), then type sequences such as Alt H O I, Alt H V V or
  Alt W F F. Excel's Office 2003 access keys also work, e.g. Alt E S V for Paste Special > Values.
- **Outline (grouping).**
  - Alt+Shift+Right / Left groups and ungroups rows or columns.
  - +/- buttons, level buttons and Ctrl+8 show and hide the outline symbols.
  - Hide / Show Detail, Auto Outline, and summary rows above or below the detail (Settings).
- **.xlsx files.**
  - Values, formulas and their cached results; dates, errors and newer functions are written the way Excel expects.
  - Cell, row and column styles, with theme and indexed colours resolved.
  - Column widths, row heights, merged cells, frozen panes, tab colours and hidden sheets.
  - Defined names, notes and the outline.

## Install and run

Requires Node.js 20 or later and npm. It is tested on Linux arm64 and should also run on x64.

```sh
git clone https://github.com/KoretyAutomate/Sheetly.git
cd Sheetly
npm install                  # also applies the fixes in patches/ (patch-package)
./sheetly.sh                 # builds on first run, then starts the app
./sheetly.sh --install-desktop   # optional: add Sheetly to your desktop's application menu
```

`sheetly.sh` starts Electron with `--no-sandbox`, because many distributions do not ship Electron's setuid
sandbox helper. Remove that flag if your system supports the Chromium sandbox.

## Tests

The app is driven through Playwright for Electron, and each suite writes its results to `test_results/`.

```sh
npx vite build
node test/formulas_full.mjs        # formula results vs Excel's documented examples
node test/shortcuts_full.mjs       # every shortcut and KeyTip sequence (about 14 minutes)
node test/outline.test.mjs         # grouping with real mouse clicks
node test/round2.test.mjs          # F6, Ctrl+F1, Ctrl+Q, Shift+F8, outline settings, Auto Outline
node test/xlsx.test.mjs            # .xlsx read / write fidelity (Node only, no app)
node test/flashfill.test.mjs
python3 test/coverage.py test_results/coverage.txt   # function coverage vs Excel's function list
```

## Layout

| Path | What it holds |
|---|---|
| `src/main.js` | Starts Univer and wires everything together. |
| `src/excel-shortcuts.js` | Excel shortcuts that Univer lacks or binds differently. |
| `src/keytips.js`, `src/keytips-tree.js` | Ribbon KeyTips and the Office 2003 access keys. |
| `src/outline.js`, `src/outline-model.js` | Row and column grouping. |
| `src/quick-analysis.js`, `src/panes.js` | Ctrl+Q; F6 and Ctrl+F1. |
| `src/excel-functions.js`, `src/groupby.js` | Excel functions that are not in Univer. |
| `src/xlsx.js` | Conversion between .xlsx files and Univer's workbook snapshots. |
| `patches/` | Fixes to Univer and ExcelJS, re-applied by `npm install`. |
| `PLAN.md` | Design notes, status and the list of known gaps. |

## Known gaps

- TRANSLATE, DETECTLANGUAGE and STOCKHISTORY return `#CONNECT!`, because they need Microsoft's online services.
- CUBE functions return `#NAME?` (there is no Data Model), and GETPIVOTDATA returns `#REF!` (there are no PivotTables).
- Charts are clustered column only and are not saved to .xlsx.
- These are not written to .xlsx: conditional formats, data validation, tables and hyperlinks. Spilled dynamic
  arrays are saved as their anchor formula only.
- KeyTips exist only for commands Sheetly has. Some KeyTip letters come from a single source and still need
  checking against a live Excel; `PLAN.md` lists them.
- Expected formula results come from Excel's documentation, not from running Excel.

## License

Apache License 2.0; see [LICENSE](LICENSE). The bundled dependencies keep their own licences: Univer (Apache-2.0),
ExcelJS (MIT), Chart.js (MIT) and Electron (MIT).
