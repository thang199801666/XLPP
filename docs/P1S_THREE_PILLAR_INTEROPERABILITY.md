# P1S-A — XLSX / Charts / Pivot interoperability push

P1S-A deliberately concentrates on the three areas used as the next competitive gates against openpyxl: basic workbook read/write compatibility, chart generation, and PivotTable physical-cache fidelity. The milestone keeps the P1R transactional/input-hardening work and adds independent host-consumer validation rather than relying only on XL++ self round-trips.

## 1. Basic XLSX read/write

### Template workbook identity

`Workbook` now models template identity explicitly:

```cpp
xlpp::Workbook wb;
wb.setTemplate(true);
wb.addWorksheet("Template");
wb.save("model.xltx");
```

`Workbook::isTemplate()` reports the identity loaded from `[Content_Types].xml`. Save chooses the workbook main content type from the template/VBA combination:

- normal + no VBA: XLSX workbook main type;
- normal + VBA: XLSM macro-enabled main type;
- template + no VBA: XLTX template main type;
- template + VBA: XLTM macro-enabled template main type.

Both `.xltx` and generated `.xltm` are covered by the dedicated P1S regression and by an independent openpyxl 3.1.5 load check. Template identity survives XL++ load/save round-trip.

This change closes a concrete workbook-format gap while retaining P1R transactional load, strict numeric/package validation, preservation of unknown OPC parts, streaming I/O, encryption, ZIP64 and resource limits.

## 2. Chart generation

P1S tightens chart generation at the schema level instead of treating all series as category/value pairs.

### Scatter

Generated scatter charts now emit `xVal` / `yVal` numeric references and use a value X axis. `Chart::scatterStyle()` / `setScatterStyle()` expose the standard scatter-style modes used by the generator and loader.

### Bubble

`ChartSeries` now has a separate bubble-size reference/cache. Generated Bubble charts emit `xVal`, `yVal` and `bubbleSize` rather than falling through the generic `cat`/`val` path.

### Combined plots and secondary axes

`Chart` adds a generated multi-plot model:

```cpp
xlpp::Chart chart(xlpp::Chart::Type::Bar);
chart.addPlot(xlpp::Chart::Type::Bar, xlpp::Chart::Grouping::Clustered, false);
chart.addSeriesToPlot(0, primarySeries);
chart.addPlot(xlpp::Chart::Type::Line, xlpp::Chart::Grouping::Standard, true);
chart.addSeriesToPlot(1, secondarySeries);
```

The writer emits independent plot groups sharing the X axis and can create a secondary Y axis. The current generated multi-plot path covers Bar, Line, Area, Scatter and Bubble families, with an explicit guard against mixing incompatible categorical-axis and XY-axis families in one generated chart.

The P1S host check opens the generated file with openpyxl 3.1.5. It recognizes the standalone charts as `ScatterChart` and `BubbleChart`, and the combined chart as a Bar chart containing two plot objects (`BarChart` + `LineChart`) with primary axis IDs `10/100` and secondary plot axis IDs `10/200`.

Chartsheets are intentionally not claimed in P1S-A. They require first-class sheet-kind/order ownership in the workbook object model and are carried forward as a high-priority chart target rather than being implemented as a preservation shortcut.

## 3. PivotTables

### Typed physical cache values

Pivot cache records previously stored string payloads and could therefore lose the physical distinction between values such as the string `"00123"` and the number `123`. P1S adds `PivotCacheValueKind`:

- Missing
- Number
- String
- Boolean
- Error
- DateTime

`PivotCache` now supports typed record setters/accessors. The save path writes the matching SpreadsheetML Pivot cache value elements (`m`, `n`, `s`, `b`, `e`, `d`), and the load path restores the physical kinds. Shared-item uniqueness is type-aware rather than text-only.

Typed metadata is preserved across supported cache-field structural mutations. Calling the legacy mutable `records()` accessor intentionally invalidates exact type metadata because arbitrary string mutation does not communicate a physical Pivot value kind; callers that require exact types should use the typed APIs.

P1S also corrected Pivot shared-item metadata emitted by the generator so the generated cache definition is accepted by openpyxl 3.1.5. The independent host check loads the generated workbook and discovers one PivotTable.

## Verification

Final P1S verification in the development environment:

- main regression: **198/198 suites, 3,632/3,632 checks PASS**;
- `XLPP_P1P_LazyFormulaTests`: PASS;
- `XLPP_P1Q_CoreHardeningTests`: PASS;
- `XLPP_P1R_TransactionalLimitsTests`: PASS;
- `XLPP_P1S_ThreePillarTests`: PASS;
- C API smoke: PASS;
- standalone public-header target: PASS;
- installed `find_package(XLPP CONFIG)` consumer: PASS;
- source `add_subdirectory()` consumer: PASS when built with the verification `-O0 -DNDEBUG -g0` profile; the aggregate Release package-consumer harness remains compiler-time dominated by the large workbook translation unit;
- focused Clang 17 ASan + UBSan + leak-detection run of `XLPP_P1S_ThreePillarTests`: PASS;
- independent openpyxl **3.1.5** host check: XLTX/XLTM template identity recognized; generated Scatter/Bubble/Bar+Line-secondary chart structures recognized; generated typed PivotTable loaded and discovered.
- reproducible optional host-check script: `tests/interop/p1s_openpyxl_host_check.py`.

## ABI note

P1S adds state to `Workbook`, `Chart`, `ChartSeries` and `PivotCache`, so the native C++ object layouts change. Source usage is additive, but **P1S-A is a native C++ ABI break** relative to P1R-A. Rebuild XL++ and native consumers/bindings together. Opaque C API callers remain source-compatible after the wrapper/core are rebuilt.

## Next three-pillar targets

The next pass should continue depth, not return to unrelated feature breadth:

1. Basic XLSX: first-class workbook sheet-kind/order model, chartsheet package integration, broader XLTX/XLTM external corpus, stricter workbook-part compatibility corpus and differential Excel/openpyxl/LibreOffice round-trip gates.
2. Charts: chartsheets, generated Stock/Surface/3-D/combo corpus gates, richer combined-plot axis ownership, chart data-label/error-bar/trendline generation parity, and host validation of every generated family.
3. PivotTables: stronger shared-item mutation/compaction with index repair, page/data/filter edge cases, grouping/date bins, PivotChart ownership, external Excel-produced corpus, and eventually slicer/timeline foundations where they can be preserved safely.
