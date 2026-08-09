# XL++ Feature Test Coverage

Updated: **2026-08-09**

## Current automated result

| Metric | Result |
|---|---:|
| Current unit test suites | **177 / 177 PASS** |
| P0Z-I native regression | **177 / 177 PASS; 3,225 / 3,225 checks** |
| Assertions/checks | **3,225 / 3,225 PASS** |
| Public headers compiled independently | **56 / 56** |
| Example sources compiled | **43 / 43** |
| `.cpp` line coverage | **92.86%** |
| `.cpp` branch coverage | **73.31%** |
| `.cpp` function coverage | **99.09%** |

The line/branch/function percentages above are the last full LLVM coverage measurement from the P0Z-H baseline; v1.12.0 re-runs functional, sanitizer, fuzz and corpus gates but does not claim a newly sampled line-coverage percentage. Coverage was measured with Clang/LLVM 17 using `-fprofile-instr-generate -fcoverage-mapping`.

## Reader-focused coverage

The reader tests are no longer limited to write-then-read round trips. Three fixtures construct OOXML packages manually, without calling the XL++ workbook writer, and then load them through `Workbook::load()`:

1. **External cell and style fixture** — shared/inline strings, numbers, booleans, errors, formulas and cached values, rich-text runs, dates, fonts, fills, borders and alignment.
2. **External worksheet feature fixture** — row/column dimensions, merge ranges, split panes, zoom, protection, filters/sorting, conditional formatting/DXF, validation, external/internal hyperlinks, comments, page setup, margins, print options and header/footer.
3. **External workbook metadata fixture** — 1904 date system, core/custom properties, workbook protection, calculation properties, defined names, print area and print titles.

Existing tests still retain round-trip coverage because it is useful for detecting serializer/loader incompatibilities, but these external fixtures exercise the loader independently from the serializer.


## P0Z-I Phase 28–37 refinement coverage

P0Z-I adds regression for dependency-driven dirty recalculation, typed calculation modes, advanced AutoFilter round-trip, preservation-first External Data/Data Model inspection, C ABI v2 capability negotiation, package consumers, enterprise corpus delta checks and preservation of an empty LibreOffice custom-properties part. Reliability verification covers all 177 suites under ASan+UBSan in bounded slices, and the Clang 17 load/validate/resave fuzzer completes 1,000 runs over 16 OpenPyXL/LibreOffice seed workbooks with no finding. The enterprise corpus foundation completes 16/16 `noop`/`edit-a1`/`add-sheet`/`copy-sheet` scenarios with zero unexpected removed parts.

## VBA source/project coverage

The P0Z-H VBA regression covers:

- Standard module create/read/update/remove and CRLF normalization;
- real CFB/OVBA `vbaProject.bin` construction and reload;
- Class modules plus `MODULEREADONLY` / `MODULEPRIVATE`;
- `ThisWorkbook` and worksheet Document modules with event procedures;
- stable worksheet VBA code names across save/load, copy and deletion/reordering;
- copied worksheet event source under a new code name;
- project Name, Description, Help File, Help Context ID and conditional-compilation constants;
- project metadata preservation across subsequent module/topology rebuilds;
- raw project byte access/export;
- external project metadata inspection while destructive source/metadata rewrite remains blocked;
- signature-part detection, source-editability state and signed-project rewrite rejection;
- invalid module/code-name rejection;
- `.xlsm` content types, workbook VBA relationship and worksheet `sheetPr/@codeName`.

Generated projects are source-only. XL++ does not parse/execute VBA, generate UserForm designer/FRX resources, author arbitrary type-library references, generate project passwords/locks, or create/re-sign digital signatures. Existing unknown external VBA projects remain a preservation surface rather than being rebuilt.

## Feature families covered

- Cell references, every cell value type, errors, formulas, dates, rich text and optional metadata
- Cell/range overloads, indexed ranges, reversed bounds, clear operations and stored-reference mutation
- Rows, columns, dimensions, merges, structural edits, iteration and worksheet reference stability
- Font, fill, border, alignment, number format, protection, named styles and styled empty cells
- AutoFilter operators, mutation, clear operations, sorting and round-trip
- Formula, CellIs, data bar, color scale and icon-set conditional formatting
- Data-validation types, operators, prompts/errors, mutation, clear and package schema
- Tables, defined names, print area/titles, document properties and custom properties
- Comments/notes, hyperlinks, rich-text cells, images, charts and pivot model/package generation
- Sheet view, split/frozen panes, tab color, zoom, right-to-left and outline symbols
- Page setup, margins, print options, header/footer and worksheet/workbook protection
- VBA binary attach/preserve/export plus Standard/Class/Document source authoring, stable code names, project metadata and signed-project safety
- Shared strings, streaming reader/writer, iterator operations and move operations
- ZIP/ZIP64, compression modes, progress/cancellation, malformed input and package preservation
- XML pull/scanner helpers, SIMD scanning, memory-mapped files and thread-pool execution
- Strict/transitional OOXML, date systems, copy/round-trip and unknown binary part preservation

## Remaining gaps

- Microsoft Excel Desktop macro execution and Trust Center behavior cannot be automated on the Linux verification host.
- Windows-specific file mapping/filesystem paths require a Windows CI runner.
- VBA UserForms/designer/FRX authoring, custom references, password/lock authoring and digital-signature generation are not generated from text; signed/external projects are preserve-only.
- Charts, pivots and drawings still rely more heavily on round-trip/package tests than on broad third-party fixture corpora.
- Defensive malformed-input branches require continued fuzzing to improve branch coverage materially.


## Chart preservation/editing coverage

The preservation corpus includes independent OpenPyXL and LibreOffice chart fixtures plus generated XL++ charts covering common 2-D charts, combined/secondary-axis charts, labels/trendlines/error bars, per-point formatting, chart layout/axes/legend, auxiliary objects, stock charts, 3-D/surface charts, projected pie/doughnut/radar, style/theme resources and series caches. P0U added automatic cache synchronization from worksheet ranges and theme-transform resolution.


## P0Z-G hardening coverage

P0Z-G adds release-gated coverage for durable path-save behavior, dependency-aware worksheet cloning (formula/Table/Chart/Pivot/local DefinedName), effective Pivot and Chart model diagnostics, strict architecture boundaries, full ASan+UBSan execution and a Clang/libFuzzer load/validate/resave harness. The feature-family coverage below remains inherited from P0Z-F.

P0Z-F expands generation and semantic reload across the current Excel Insert-chart families. Classic ChartML coverage includes column/bar, line, pie/doughnut, area, XY scatter, bubble, stock, surface, radar and their supported 3-D/projected variants. Modern Excel chart families use ChartEx parts and relationships for Histogram, Pareto, Box & Whisker, Waterfall, Funnel, Treemap, Sunburst and Filled Map. Combo plots can select series ranges and primary/secondary axes, while Scatter and Bubble now serialize native `xVal`/`yVal` payloads and Bubble size data rather than categorical stand-ins.

## PivotTable coverage

P0Z-F adds worksheet-source PivotTable authoring and semantic reload for row/column/page/data fields, multiple data fields with the synthetic Values field, Cartesian row/column item tuples, hidden items, field behavior, subtotals, styles/display options, cache lifecycle metadata, typed cache records, numeric/date grouping and Pivot filters including Top-N. `Show Values As` covers the legacy modes and the Office 2010+ x14 extension modes such as percent-of-parent, percent-of-running-total and ranking. Generated cache/definition XML is regression-tested through save, package inspection and reload.

Still outside the complete Pivot ecosystem are calculated fields/items, PivotCharts, OLAP/Data Model caches, slicers and timelines. Microsoft Excel Desktop recovery-log validation remains pending for both charts and pivots.
