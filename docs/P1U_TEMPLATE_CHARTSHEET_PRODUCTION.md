# P1U-A — Template & Chartsheet Production Semantics

P1U-A focuses narrowly on `.xltx` / `.xltm` and first-class Chartsheets. It builds on P1T's mixed Worksheet/Chartsheet ownership model and fills the workbook-state and chart-sheet metadata gaps needed for production template workflows.

## Delivered

### Mixed workbook state

- `WorkbookSheetVisibility::{Visible, Hidden, VeryHidden}` applies uniformly to Worksheets and Chartsheets.
- `Workbook::setWorkbookSheetVisibility()` refuses to hide the final visible tab.
- `Workbook::activeWorkbookSheetIndex()`, `setActiveWorkbookSheetIndex()` and `setActiveWorkbookSheet()` work across mixed sheet kinds.
- Mixed-sheet reorder/removal keeps the active index bound to the same logical tab where possible and repairs it to a surviving visible tab otherwise.
- Workbook XML writes and reads `workbookView@activeTab`, `firstSheet` and `sheet@state` without changing the legacy worksheet-only meaning of `sheetCount()` / `sheetNames()`.
- Model validation reports hidden active tabs, out-of-range active tabs and workbooks with no visible tab.

### XLTX / XLTM package identity

- `Workbook::setTemplate(true)` remains independent from VBA identity.
- Template + no VBA emits the `.xltx` main workbook content type.
- Template + VBA emits the `.xltm` macro-enabled template content type.
- Load recovers template identity from `[Content_Types].xml`.
- Active Chartsheets, hidden/veryHidden mixed tabs and Chartsheet metadata survive template save/load.

### Chartsheet metadata model

`Chartsheet` now exposes independent sheet-level state in addition to its chart:

- `ChartsheetProperties`: code name, tab color, published flag.
- `ChartsheetView`: workbook-view id, selected-tab flag, zoom scale, zoom-to-fit.
- `ChartsheetProtection`: content/object protection and legacy password hash/password helper.
- `PageMargins`.
- `PageSetup`: orientation, paper size, scale, fit-to-page geometry, first page number, black-and-white/draft flags.
- `HeaderFooter`, including odd/even and newly modeled first-page header/footer strings.

The serializer emits these elements in Chartsheet schema order and the loader restores them with strict numeric parsing.

### Preservation-safe imported Chartsheets

Imported Chartsheets have two independent mutation domains:

1. sheet metadata (`sheetPr`, views, protection, page setup/header/footer);
2. chart/drawing subtree.

A metadata-only edit patches the original Chartsheet part while preserving imported DrawingML and ChartML bytes. Mutable `Chartsheet::chart()` access opts into chart regeneration. Unmodeled top-level Chartsheet extension nodes such as `customSheetViews`, `drawingHF`, `picture`, `webPublishItems` and `extLst` are carried forward when the modeled part is patched.

P1U also hardens repeated saves on the same `Workbook` object:

- metadata-only patches do not revert to original source XML on the second save;
- regenerated imported charts do not fall back to the old preserved chart subtree on later saves.

### Streaming and C ABI

The P1T streaming mixed-sheet APIs continue to expose worksheet/chartsheet names and original tab order. P1U adds C ABI access for:

- tab visibility get/set;
- active tab get/set;
- template identity get/set.

The C API smoke test exercises these additions together with a first-class Chartsheet.

## Interoperability gates

The P1U fixture corpus includes an openpyxl-generated advanced `.xltx` Chartsheet with:

- active chart-only tab;
- code name, published state and tab color;
- zoom/view state;
- Chartsheet protection;
- page margins and setup;
- odd and first-page headers/footers;
- an owned chart.

Independent openpyxl 3.1.5 validation accepts XL++ generated `.xltx` and `.xltm` packages, sees the active Chartsheet and modeled metadata, preserves chart ownership after metadata-only patching, and sees the regenerated chart title after repeated saves.

The package validator reports zero relationship/content-type/owner/orphan errors for generated `.xltx`, patched imported `.xltx`, and `.xltm` template artifacts.

## Verification

- Main regression: **198/198 suites, 3,632/3,632 checks PASS**.
- P1P/P1Q/P1R/P1S/P1T/P1U dedicated tests: PASS.
- C API smoke: PASS.
- Standalone public-header target: PASS.
- Installed `find_package(XLPP CONFIG)` consumer using template/active/hidden Chartsheet state: PASS.
- Source `add_subdirectory()` consumer using the same state: PASS.
- Focused Clang 17 ASan + UBSan + leak detection on `XLPP_P1U_TemplateChartsheetTests`: PASS.
- Independent openpyxl 3.1.5 host gate: PASS.
- `xlpp-package-validator`: PASS on generated/patched XLTX and XLTM artifacts.

## Compatibility / ABI

P1U is additive at the source-API level, but `Workbook`, `WorkbookSheetDescriptor`, `Chartsheet` and `HeaderFooter` object layouts changed. Native C++ consumers and bindings must rebuild with the P1U headers/library. The opaque C API remains source-compatible after rebuilding.

## Next template/chartsheet targets

1. Preserve and expose additional workbook-view/window geometry while keeping mixed active/visibility state canonical.
2. Add Chartsheet print options, drawing-header/footer picture ownership and richer custom-sheet-view semantics.
3. Add more Excel Desktop / LibreOffice `.xltx/.xltm` corpus files and differential package/host checks.
4. Pull the P1U mixed-sheet/template state through runtime-tested Python and C# wrappers.
5. Extract Chartsheet/package writer logic from `Workbook.cpp`; P1U leaves the TU around 581 KiB / 9.6k lines.
