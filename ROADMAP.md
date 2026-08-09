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


### Completed development note — P0V-A

- Dependency snapshots are registered per chart/series/title-category-value reference after successful synchronization.
- `ChartCacheSyncOptions::changedReferencesOnly` skips cache rebuild/patch work when the exact source range fingerprint is unchanged.
- Source fingerprints include stored cell values, formulas, number formats and workbook date epoch, so chart-visible source changes invalidate deterministically while unrelated edits do not.
- `ChartCacheSyncReport` now exposes checked/unchanged/registered/changed dependency counters.
- `Workbook::resetChartCacheDependencyTracking()` and `trackedChartCacheDependencyCount()` expose lifecycle/introspection controls.
- `SaveOptions::synchronizeChartCaches` adds opt-in pre-save synchronization on a private workbook copy; caller state stays unchanged.
- Regression baseline: 155/155 suites, 2,731/2,731 checks passing; public-header, C API and package consumers build successfully.

### Completed development note — P0W

- In-process formula calculation with recursive/cross-sheet references, defined names, array constants, cycle detection, cached-result updates and calculate-before-save.
- Common aggregate/statistical, logical/error, math/trigonometric, text, date/time, criteria/multi-criteria, lookup and financial function families.
- Office Agile AES-256-CBC/SHA-512 password-to-open encryption with DataSpaces, HMAC integrity and CFB miniFAT/FAT/DIFAT writer/reader.
- Standard AES-128/192/256 + SHA-1 encrypted OOXML reader and password verifier.
- Password rotation/removal, CSPRNG/constant-time verification, malformed-container bounds/cycle hardening and >9 MiB DIFAT regression.
- Independent LibreOffice interoperability in both directions.
- Regression baseline: 160/160 suites, 2,838/2,838 checks passing.

### Completed development note — P0X / v1.2.0

- General-purpose XLSX editing-core weighted gate: **90.7/100**.
- Workbook dependency graph for cell/range, defined-name, table, external and volatile dependencies.
- Transactional reference-safe insert/delete rows/columns with rollback and stable live handles.
- Formula core matrix expanded from 21/26 to **26/26 capability families**, including dynamic arrays, structured references, iterative calculation and external resolvers.
- Imported PivotTables now support semantic edit/regeneration with common layout/style/cache/data/report-filter options while untouched pivots remain preservation-first.
- Standard AES-128/192/256 + SHA-1 encryption writer added; password-encryption matrix is **19/21 (90.5%)**.
- C ABI smoke test is now registered with the project test option; Python/C#/C parity expanded.
- Buildable source sample and Linux/Windows/macOS CI definition added.
- Native regression baseline at closeout: see `PACKAGE_STATUS.md` and `docs/P0X_90_PERCENT_ENGINE.md`.

### Completed development note — P0Z-I / v1.12.0

- Formula dispatch gains an internal function-family registry and dependency-driven dirty recalculation from explicit changed-cell roots.
- Calculation mode has a typed API for automatic, automatic-except-data-tables and manual behavior.
- AutoFilter round-trips Top10, dynamic, color, icon and date-group filter forms.
- External workbook links, connections/query tables, Power Query-related parts and Data Model/OLAP state gain preservation-first inspection APIs.
- C ABI advances to version 2 with additive capability negotiation; Python/.NET packaging workflows and .NET `SafeHandle` ownership are formalized.
- Native benchmark budget/guard, strict/sanitizer/fuzz CI and an enterprise corpus runner become first-class release infrastructure.
- Enterprise corpus result: **16/16 scenarios PASS, 0 unexpected removed parts**; it found and forced a fix for empty `docProps/custom.xml` preservation.
- Native regression baseline: **177/177 suites, 3,225/3,225 checks PASS**; full suite covered under ASan+UBSan and 1,000 seeded libFuzzer runs complete without a finding.
- See `docs/P0ZI_PHASE28_TO_37_REFINEMENT.md` for completed scope and explicit carry-over.

### Completed development note — P0Z-H / v1.11.0

VBA authoring now covers Standard/Class/Document source modules, workbook and worksheet event code, stable host code names, module read-only/private flags, project metadata and conditional compilation constants, raw project export, and explicit signed/external-project mutation safety. UserForms/FRX/designer authoring, custom reference authoring, project-password generation and digital-signature generation remain separate roadmap items.

### Completed development note — P0Z-G / v1.10.0

- Source-level strict-warning cleanup: full core builds with 0 warnings under the release strict-warning profile.
- Durable atomic path save with file flush and POSIX parent-directory fsync / Windows write-through semantics.
- Dependency-aware worksheet copy across formulas, charts, pivots, workbook-global object identifiers and local DefinedName scopes.
- Effective-state Pivot validation and deeper Chart reference/family validation.
- Additional architecture dependency guards plus Clang/libFuzzer load/validate/resave harness.
- Full native ASan + UBSan regression clean; seeded libFuzzer smoke clean.
- Native regression baseline: **174/174 suites, 3,164/3,164 checks PASS**.

### Next development target — post-v1.12 compatibility depth

- Run the existing Microsoft Excel Desktop fixture/repair-log automation on a Windows host with Excel installed and promote failures to release blockers.
- Expand the enterprise corpus from the current seed foundation toward hundreds of anonymized production workbooks.
- Continue Formula Engine 2.0 with LAMBDA-family semantics and DAG scheduling only after compatibility gates remain stable.
- Deepen Pivot ecosystem support: calculated fields/items, PivotCharts, slicers/timelines and OLAP/Data Model boundaries.
- Deepen VBA UserForm/FRX/reference handling and DrawingML/SmartArt preservation/editing.
- Execute and publish Python wheel/.NET package matrices through the new release workflows.

