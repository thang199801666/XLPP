# P1V-A Template / Chartsheet Depth & I/O Decomposition

P1V-A continues the `.xltx` / `.xltm` and Chartsheet production track from P1U. The batch deliberately combines deeper chart-only sheet semantics with a small architectural extraction so new Chartsheet work does not continue inflating the main `Workbook.cpp` translation unit.

## Delivered

### Chartsheet I/O boundary

Chartsheet XML parsing/serialization and chart-only drawing relationship helpers now live in `src/XLPP/Workbook/WorkbookChartsheetIO.cpp`. The main `Workbook.cpp` drops from roughly 9,607 to 9,442 lines. Public APIs are unchanged by the extraction.

### Advanced page setup

`PageSetup` now models paper height/width, page order, printer-default behavior, comment/error rendering, horizontal/vertical DPI, copies and printer-settings relationship identity. The fields round-trip through both worksheet and Chartsheet page-setup XML.

### Modern protection descriptors

`ChartsheetProtection` retains `algorithmName`, `hashValue`, `saltValue` and `spinCount` in addition to existing content/object/legacy-password state. The API preserves descriptors rather than pretending to derive them from a plaintext password.

### Custom Chartsheet views

`CustomChartsheetView` models GUID, scale, visible/hidden/veryHidden state, zoom-to-fit and optional nested page margins, page setup and header/footer. Imported raw `customSheetViews` XML remains byte-preserved until mutable semantic access occurs. Duplicate/missing GUIDs are rejected or diagnosed.

The P1V regression suite found and fixed a real state-loss defect in the first implementation: repeatedly requesting a mutable nested page/header object recreated the optional payload and erased previously assigned values. Accessors now materialize once and preserve state.

## Interoperability gates

Generated advanced XLTX/custom-view/worksheet artifacts are consumed independently by openpyxl 3.1.5. The host sees the expected paper dimensions/order, printer settings, DPI/copies, modern protection descriptor and nested custom-view settings. The XL++ package validator reports no relationship, duplicate-ID, dangling, orphan, content-type or owner-reference errors for the generated artifacts.

## Verification

- Main regression: 198/198 suites, 3,632/3,632 checks PASS.
- P1P/P1Q/P1R/P1S/P1T/P1U/P1V dedicated regressions: PASS.
- C API smoke: PASS.
- Standalone public headers: PASS.
- Installed `find_package(XLPP CONFIG)` consumer: PASS.
- Source `add_subdirectory()` consumer: PASS.
- Focused Clang 17 ASan+UBSan+leak detection on P1V: PASS.
- Independent openpyxl 3.1.5 host check: PASS.
- Package graph validation: PASS with zero reported errors.

## Compatibility

P1V is source-additive for ordinary callers, but `PageSetup`, `ChartsheetProtection` and Chartsheet-related value layouts gain state. Native C++ consumers and bindings should be rebuilt together with the core; do not mix binaries compiled against P1U headers with P1V libraries. Project version remains `1.1.2` because P1V-A is a development milestone.

## Next

Continue extracting Chartsheet lifecycle/package-closure writer code, then model printer-settings and drawing-header/footer picture ownership and add Excel Desktop/LibreOffice template corpora.
