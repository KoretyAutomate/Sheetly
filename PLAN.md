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

## Status 2026-10-14
Phase 1-3 done and tested: formulas 83/83, shortcuts+xlsx 12/12 (test_results/).
Open gaps: Ctrl+1 Format Cells dialog, Ctrl+-/Ctrl++ delete/insert cells, Ctrl+F/H find-replace, Ctrl+K,
Ctrl+Shift+L filter, Ctrl+T tables, charts/pivots (Univer advanced preset not used), xlsx styles beyond
bold/italic/numfmt/widths/merges (fills, borders, fonts colours), open-from-file-manager, AppImage packaging (Phase 4).
Formula coverage beyond the 83-case table is unmeasured.
