# XL++ Core-to-100% Roadmap

## Current focus: production-grade read/write core

1. **Milestone 15 — Streaming foundation** — complete
   - Append-only streaming writer
   - Row callback reader
   - Chunked zlib ZIP output
   - File-backed package entries

2. **Milestone 16 — Direct streaming ZIP reader** — complete
   - Central-directory index without loading the full archive
   - Incremental inflate of worksheet entries
   - Pull iterator and callback APIs
   - Shared-string lazy access and cache policy

3. **Milestone 17 — Fast XML scanner** — complete
   - Non-allocating token scanner based on `string_view`
   - Direct numeric conversion with `from_chars`
   - Reduced temporary strings and attribute maps
   - Benchmarks against the current parser

4. **Milestone 18 — Shared strings and date/time core** — complete
   - Streaming shared-string writer
   - Deduplication modes: disabled, hash, bounded LRU
   - Excel serial date/time conversion
   - ISO date cells and workbook date epoch

5. **Milestone 19 — Parallel package pipeline** — complete
   - Parallel worksheet serialization
   - Optional parallel deflate per package entry
   - Deterministic package ordering
   - Configurable compression level and strategy

6. **Milestone 20 — Core compatibility completion** — complete
   - Preserve unknown package parts and relationships
   - Strict/transitional namespace handling
   - Malformed-file diagnostics and recovery options
   - ZIP64 and large-entry support
   - Cancellation, progress callbacks, and limits

## Core 100% definition

“Core 100%” means complete and robust handling of workbook, worksheet, row, cell, value, formula metadata, styles, dimensions, relationships, shared strings, dates, errors, package preservation, and both DOM and streaming I/O. It does not mean 100% parity with every optional openpyxl chart, pivot, macro, or drawing feature.


### Completed development note — P0Q

- First-class StockChart read/generate foundation (`stockChart`).
- Stock high-low/up-down auxiliary interoperability through the shared plot model.
- New-chart serialization of P0P auxiliary objects via `Chart::primaryPlot()`.
- Data-table `txPr` default text-style read/write support.
- OpenPyXL and LibreOffice host validation for imported and generated stock charts.


### Completed development note — P0R

- Namespace-tolerant read model for `bar3DChart`, `line3DChart`, `area3DChart`, `pie3DChart`, `surfaceChart`, and `surface3DChart`.
- `view3D` plus floor/side-wall/back-wall thickness and fill/line metadata.
- Stable-ID selective edits for 3D view and wall formatting without rewriting untouched chart siblings.
- Generation foundation for Bar3D and Surface3D with `catAx`/`valAx`/`serAx`, gap depth/shape and wireframe support.
- OpenPyXL host validation for all six imported types and generated Bar3D/Surface3D; LibreOffice host normalization of Surface/Surface3D to Bar3D is documented.

### Completed development note — P0S

- First-class Pie-of-Pie and Bar-of-Pie read/generate support through native `ofPieChart`/`ofPieType`.
- Projected-pie split model for gap width, split criteria/position, custom second-plot points, second-plot size and series-line formatting.
- Doughnut/Pie first-slice angle, Doughnut hole-size and Radar native style read/generate/selective-edit support.
- Independent OpenPyXL corpus covers Pie-of-Pie, Bar-of-Pie, Doughnut, Radar and an image sibling with byte-preservation checks.
- Direct OpenPyXL and LibreOffice host validation completed; Calc normalization of projected-pie split parameters and doughnut hole size is documented.


### Completed development note — P0T

- Workbook theme palette parsing for chart scheme colors and base-color resolution.
- Chart style ID plus chart-style/color-style relationship resource discovery and preservation.
- `strCache` / `numCache` read model for series title/category/value references, including numeric format codes and indexed points.
- Selective cache/style edits on imported charts plus first-class cache generation for new charts.
- OpenPyXL direct host validation completed; LibreOffice cache recomputation and style-resource removal on re-save are documented as host normalization.


### Completed development note — P0U

- Workbook-level `synchronizeChartCaches()` rebuilds title/category/value caches from current worksheet A1 ranges.
- Local and quoted cross-sheet references are supported, including escaped apostrophes in sheet names; external, union, structured and 2-D references are skipped with diagnostics.
- Sparse cache indexes preserve blank source cells while `ptCount` continues to represent the referenced range length.
- `ChartSeriesCache` now exposes duplicate-index, ordering, sparse-state and public consistency validation helpers.
- Theme parsing now includes major/minor Latin font metadata and format-scheme counts, plus sequential final RGB/alpha resolution for DrawingML color transforms.
- OpenPyXL validates direct synchronized caches; LibreOffice cache/materialization normalization on re-save is documented.
