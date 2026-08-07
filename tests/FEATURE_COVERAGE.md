# XL++ Feature Test Coverage

Updated: **2026-08-07**

## Current automated result

| Metric | Result |
|---|---:|
| Current unit test suites | **154 / 154 PASS** |
| Clean Release verification | **154 / 154 PASS** |
| Assertions/checks | **2,708 / 2,708 PASS** |
| Public headers compiled independently | **56 / 56** |
| Example sources compiled | **43 / 43** |
| `.cpp` line coverage | **92.86%** |
| `.cpp` branch coverage | **73.31%** |
| `.cpp` function coverage | **99.09%** |

Coverage was measured with Clang/LLVM 17 using `-fprofile-instr-generate -fcoverage-mapping`. The three unexecuted `.cpp` functions are two internal little-endian stream fallbacks in `ZipArchiveReader.cpp` and one internal worksheet-serialization lambda. No newly added VBA text API remains unexecuted.

## Reader-focused coverage

The reader tests are no longer limited to write-then-read round trips. Three fixtures construct OOXML packages manually, without calling the XL++ workbook writer, and then load them through `Workbook::load()`:

1. **External cell and style fixture** — shared/inline strings, numbers, booleans, errors, formulas and cached values, rich-text runs, dates, fonts, fills, borders and alignment.
2. **External worksheet feature fixture** — row/column dimensions, merge ranges, split panes, zoom, protection, filters/sorting, conditional formatting/DXF, validation, external/internal hyperlinks, comments, page setup, margins, print options and header/footer.
3. **External workbook metadata fixture** — 1904 date system, core/custom properties, workbook protection, calculation properties, defined names, print area and print titles.

Existing tests still retain round-trip coverage because it is useful for detecting serializer/loader incompatibilities, but these external fixtures exercise the loader independently from the serializer.

## VBA text coverage

The new VBA source-text test covers:

- creating a standard module from a C++ string;
- CRLF source normalization;
- building a real CFB `vbaProject.bin` with OVBA project/module streams;
- `.xlsm` content types, workbook VBA relationship and host code names;
- worksheet document modules (`Sheet1`, `Sheet2`, ...) plus `ThisWorkbook`;
- rebuilding document modules when worksheets are added after VBA creation or after reload;
- loading the `.xlsm` and extracting module source again;
- case-insensitive lookup;
- replacing, adding and removing modules after load;
- removing the final standard module and returning to a non-macro package;
- invalid module-name rejection.

The generated project supports standard source modules, one document module per worksheet, and the `ThisWorkbook` document module required by the host workbook. XL++ does not parse VBA syntax, execute macros, create UserForms, generate designer streams, customize references, or produce digital signatures.

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
- Existing VBA binary add/remove plus source-text module build/read/update/remove lifecycle
- Shared strings, streaming reader/writer, iterator operations and move operations
- ZIP/ZIP64, compression modes, progress/cancellation, malformed input and package preservation
- XML pull/scanner helpers, SIMD scanning, memory-mapped files and thread-pool execution
- Strict/transitional OOXML, date systems, copy/round-trip and unknown binary part preservation

## Remaining gaps

- Microsoft Excel Desktop macro execution and Trust Center behavior cannot be automated on the Linux verification host.
- Windows-specific file mapping/filesystem paths require a Windows CI runner.
- VBA UserForms, class/document module authoring, custom references, password protection and signatures are not generated from text.
- Charts, pivots and drawings still rely more heavily on round-trip/package tests than on broad third-party fixture corpora.
- Defensive malformed-input branches require continued fuzzing to improve branch coverage materially.


## Chart preservation/editing coverage

The preservation corpus now includes independent OpenPyXL and LibreOffice chart fixtures plus generated XL++ charts covering common 2-D charts, combined/secondary-axis charts, labels/trendlines/error bars, per-point formatting, chart layout/axes/legend, auxiliary objects, stock charts, 3-D/surface charts, projected pie/doughnut/radar, style/theme resources and series caches. P0U adds automatic cache synchronization from worksheet ranges and theme-transform resolution. Microsoft Excel Desktop recovery-log validation remains pending.
