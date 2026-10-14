# Sheetly
Open-source, Excel-style spreadsheet for Linux. Electron + [Univer](https://github.com/dream-num/univer) (Apache-2.0) + ExcelJS.
Not Excel's code (Excel is not open source) — an Excel-compatible app, verified by tests.

    npm install && npx vite build && ./sheetly.sh     # run
    node test/conformance.mjs                          # formula table vs Excel-documented results
    node test/shortcuts.mjs                            # Excel shortcuts + .xlsx round trip

Shortcuts Excel has that Univer lacked are added in `src/excel-shortcuts.js`.
Known gaps: see PLAN.md / test_results.
