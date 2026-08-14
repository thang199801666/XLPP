# XL++ Full Project Package — P1X-A Chartsheet Writer Decomposition & Relationship-Collision Hardening

## P1X-A verification — 2026-08-14

- Main regression: **198/198 suites PASS, 3,632/3,632 checks PASS**.
- Dedicated P1P/P1Q/P1R/P1S/P1T/P1U/P1V/P1W/P1X regression executables: PASS.
- Moved the complete Chartsheet package-write loop into `WorkbookChartsheetPackage.cpp/.h`; `Workbook.cpp` drops from about 9,547 to about 9,440 lines.
- Extracted writer retains imported/generated Chartsheet ownership, drawing/chart closure replacement, printer-settings emission, preservation sibling merge and repeated-save behavior.
- Fixed a real relationship-ID collision bug: generated relationship repair no longer globally rewrites every matching `r:id` in Chartsheet XML. Drawing collisions patch only `<drawing>` and printer-settings collisions patch only `<pageSetup>`, protecting preserved `legacyDrawingHF`/unknown owner nodes.
- Added `XLPP_P1X_ChartsheetWriterTests` with a deliberate preserved-VML/generated-drawing `rId1` collision plus repeated-save graph validation.
- C API smoke and standalone public headers: PASS.
- Installed `find_package(XLPP CONFIG)` consumer: PASS; source `add_subdirectory()` consumer: PASS with the verification optimization profile.
- Focused Clang ASan+UBSan+leak detection on P1X writer regression: PASS.
- P1X adds no new public API and does not change the P1P-era `Cell`/`Style` layout.

Detailed notes: `docs/P1X_CHARTSHEET_WRITER_DECOMPOSITION.md`.

---

# XL++ Full Project Package — P1W-A Template / Chartsheet Auxiliary Ownership Hardening

## P1W-A verification — 2026-08-14

- Main regression: **198/198 suites PASS, 3,632/3,632 checks PASS**.
- Dedicated P1P/P1Q/P1R/P1S/P1T/P1U/P1V/P1W regression executables: PASS.
- Added first-class opaque Chartsheet `printerSettings` payload ownership, including load/save/replace/clear and exact binary round-trip with embedded NUL bytes.
- Generated/replaced printer settings receive a package part/content type/relationship; imported `pageSetup@r:id` ownership is validated on load rather than treated as an uninterpreted string.
- Fixed preservation of the standard Chartsheet header/footer-picture owner node from the incorrect `drawingHF` spelling to `legacyDrawingHF`.
- Imported chart regeneration now replaces only the old drawing/chart closure while preserving sibling VML header/footer, printer-settings and unknown relationships/parts.
- `RelationshipGraph` and `xlpp-package-validator` validate `legacyDrawingHF -> vmlDrawing` and `pageSetup -> printerSettings` ownership and inventory Header/footer drawings and Printer settings separately.
- C ABI adds binary printer-settings set/size/copy/clear functions; `xlpp_capi_smoke` round-trips the payload successfully.
- Independent openpyxl 3.1.5 host validation of final printer-settings/header-footer artifacts: PASS.
- Standalone public headers, installed `find_package(XLPP CONFIG)` consumer and source `add_subdirectory()` consumer: PASS.
- Focused Clang 17 ASan+UBSan+leak detection on `XLPP_P1W_ChartsheetPackageTests`: PASS.
- `Workbook.cpp` is roughly **579 KiB / 9,547 lines**. P1X should extract generic relationship merge/closure and Chartsheet package-writer logic before another broad feature expansion.

Detailed notes: `docs/P1W_TEMPLATE_CHARTSHEET_PACKAGE_OWNERSHIP.md`.

---

# XL++ Full Project Package — P1V-A Template / Chartsheet Depth & I/O Decomposition

## P1V-A verification — 2026-08-14

- Main regression: **198/198 suites PASS, 3,632/3,632 checks PASS**.
- Dedicated P1P/P1Q/P1R/P1S/P1T/P1U/P1V regression executables: PASS.
- Chartsheet XML parsing/serialization and chart-only drawing relationship helpers moved from `Workbook.cpp` into `WorkbookChartsheetIO.cpp`; the remaining monolith falls from roughly **9,607 to 9,442 lines**.
- Chartsheet `PageSetup` now models paper height/width, page order, printer-default behavior, cell-comment/error rendering, horizontal/vertical DPI, copies and printer-settings relationship identity; the common worksheet path round-trips the same advanced settings.
- Chartsheet protection retains modern password metadata (`algorithmName`, `hashValue`, `saltValue`, `spinCount`) in addition to legacy protection flags/password state.
- Added semantic `CustomChartsheetView` support (`guid`, scale, visibility state, zoom-to-fit plus nested margins/page setup/header-footer) with preservation-first raw XML for untouched imported custom views.
- Duplicate/empty custom-view GUIDs are rejected/diagnosed; partial modern-protection tuples and unresolved generated printer-settings relationships are surfaced by model validation.
- Regression caught and fixed a lazy-accessor state-loss bug where repeated mutable access to nested custom-view page/header objects could reset previously edited fields.
- Independent openpyxl 3.1.5 host validation accepts the generated XLTX advanced page setup, modern protection, custom Chartsheet view and extended worksheet page setup.
- `xlpp-package-validator` accepts the generated XLTX/custom-view/worksheet artifacts with zero relationship, duplicate-ID, dangling, orphan, content-type or owner-reference errors.
- C API smoke, standalone public headers, installed `find_package(XLPP CONFIG)` consumer and source `add_subdirectory()` consumer: PASS. The aggregate nested package-consumer harness remains compile-time dominated, so its two constituent consumers were also verified separately.
- Focused Clang 17 ASan+UBSan+leak detection on `XLPP_P1V_ChartsheetDepthTests`: PASS.

Detailed notes: `docs/P1V_TEMPLATE_CHARTSHEET_DEPTH.md`.

---

# XL++ Full Project Package — P1U-A Template & Chartsheet Production Semantics

## P1U-A verification — 2026-08-14

- Main regression: **198/198 suites PASS, 3,632/3,632 checks PASS**.
- Dedicated P1P/P1Q/P1R/P1S/P1T/P1U regression executables: PASS.
- Mixed Worksheet/Chartsheet order now carries `Visible` / `Hidden` / `VeryHidden` state plus an active workbook-tab identity; last-visible-tab and active-hidden invariants are enforced.
- `.xltx` and `.xltm` retain template identity together with active/hidden Chartsheet state and correct template/macro-template content types.
- Chartsheets model code name/tab color/published state, zoom/view, content/object protection, page margins/setup and odd/even/first-page header/footer text.
- Metadata-only edits of imported Chartsheets preserve owned chart/drawing bytes; repeated-save regression protects both metadata-only patches and regenerated imported chart mutations from reverting to the source subtree.
- Independent openpyxl 3.1.5 host validation: generated XLTX, XLTM+VBA, active Chartsheet, modeled metadata, preservation-backed patch and repeated regenerated chart mutation PASS.
- `xlpp-package-validator`: generated XLTX, patched imported XLTX and XLTM PASS with zero relationship, duplicate-ID, dangling, orphan, content-type or owner-reference errors.
- C API smoke exercises template identity, active mixed tab and tab visibility additions: PASS.
- Standalone public headers: PASS.
- Installed `find_package(XLPP CONFIG)` consumer: PASS.
- Source `add_subdirectory()` consumer: PASS.
- Focused Clang 17 ASan+UBSan+leak detection on P1U regression: PASS.
- `Workbook.cpp` is about **581,063 bytes / 9,607 lines**; Chartsheet/package writer decomposition remains a build-scalability target.

Detailed notes: `docs/P1U_TEMPLATE_CHARTSHEET_PRODUCTION.md`.

---

# XL++ Full Project Package — P1T-A Three-Pillar Sheet Model / Chartsheet / Pivot Identity

## P1T-A verification — 2026-08-14

- Main regression: **198/198 suites PASS**, **3,632/3,632 checks PASS**.
- Dedicated P1P/P1Q/P1R/P1S/P1T regression executables: PASS.
- P1T adds first-class Chartsheets and a mixed workbook-tab model while preserving legacy worksheet-only `sheetCount()`/`sheetNames()` semantics.
- Generated Chartsheet package ownership, imported chartsheet byte-preservation, lifecycle cleanup, mixed-order streaming APIs and `LoadOptions::maxChartsheets` are covered by P1T regression.
- Pivot field item semantics can bind to typed `(kind,value)` identities and are re-resolved to the current physical shared-item `x` index after cache reorder/compaction.
- C API smoke creates, reorders, saves and reloads a Chartsheet containing a generated Bar+Line secondary-axis combo.
- Independent openpyxl 3.1.5 host validation: mixed worksheet/chartsheet order, Chartsheet combo plot families/axes, preserved imported Chartsheet and generated Pivot discovery PASS.
- `xlpp-package-validator`: generated Chartsheet/Pivot artifacts PASS with 0 relationship-syntax, duplicate-ID, dangling, orphan, content-type or owner-reference errors; inventory reports 1 Chartsheet for the chart-only package.
- Focused Clang 17 ASan+UBSan + leak detection on `XLPP_P1T_ThreePillarTests`: PASS.
- Installed `find_package(XLPP CONFIG)` consumer and source `add_subdirectory()` consumer using first-class Chartsheet APIs: PASS.
- `Workbook.cpp` is roughly **561 KiB / 9.3k lines** after P1T; extracting sheet/package/chart writer logic is a P1U build-scalability priority.

Detailed notes: `docs/P1T_THREE_PILLAR_SHEETMODEL_CHARTSHEET_PIVOT_IDENTITY.md`.


This directory is the complete XL++ source tree through P1T-A. P1T-A builds on P1S-A and continues the three competitive gates: Basic XLSX read/write, Charts and PivotTables.

## P1S-A verification — 2026-08-14

- Main regression: **198/198 suites PASS**, **3,632/3,632 checks PASS**.
- Dedicated P1P/P1Q/P1R/P1S regression executables: PASS.
- C API smoke and standalone public-header target: PASS.
- Installed `find_package(XLPP CONFIG)` consumer: PASS.
- Source `add_subdirectory()` consumer: PASS under the low-optimization verification profile; the aggregate package harness remains dominated by Release compile time of the remaining large workbook TU.
- Focused Clang 17 ASan+UBSan + leak detection for the P1S three-pillar test: PASS.
- Independent host validation with the locally installed **openpyxl 3.1.5**: generated XLTX and XLTM are recognized as templates; generated Scatter and Bubble charts are parsed as their native chart classes; the Bar+Line combination is parsed as two plots with primary/secondary axes; the generated typed PivotTable is discovered successfully.
- Basic read/write: first-class template identity plus correct XLTX/XLTM workbook content types and round-trip.
- Charts: schema-correct Scatter `xVal/yVal`, Bubble `xVal/yVal/bubbleSize`, scatter style, generated multi-plot charts and secondary Y-axis support.
- PivotTables: physical cache-value kinds (missing/number/string/Boolean/error/date-time) round-trip through cache definitions/records; type-aware shared items and corrected shared-item metadata.

Detailed notes: `docs/P1S_THREE_PILLAR_INTEROPERABILITY.md`.

### P1S-A ABI compatibility

P1S adds fields to `Workbook`, `Chart`, `ChartSeries` and `PivotCache`. Public source usage is additive, but native C++ layouts change. **Rebuild XL++ and every native C++ consumer/binding together.** Opaque C API source usage remains compatible after rebuilding the core/wrapper.

---

This directory is the complete XL++ source tree through P1R-A. It builds on P1Q-A and hardens the materialized Workbook load boundary: transactional commit, bounded stream/model materialization, exact numeric parsing, lexical XML scanning, owner-aware OPC relationships and duplicate content-type rejection. P1R-A does not change `Cell`/`Style` layout.

## P1R-A verification — 2026-08-14

- Main regression: **198/198 suites PASS**, **3,632/3,632 checks PASS**.
- Dedicated P1P, P1Q and P1R regression executables: PASS.
- Focused Clang ASan+UBSan + leak detection on P1Q/P1R malformed-input/transactional/package paths: PASS.
- Standalone public-header target: PASS.
- C API rebuild + `xlpp_capi_smoke`: PASS.
- Installed `find_package(XLPP CONFIG)` and source `add_subdirectory()` consumers: PASS.
- Same-machine 200k-cell Debug load: peak RSS about **121 MiB -> 101 MiB (~16-17% lower)**; final wall time remained within roughly **1-2%** of P1Q.
- `Workbook.cpp` remains about **527 KiB / 8.8k lines**; `WorkbookIO.cpp` and `WorkbookPackageReader.cpp` are now separate TUs.

Detailed notes: `docs/P1R_CORE_TRANSACTIONAL_PACKAGE_HARDENING.md`.

## P1Q-A verification — 2026-08-14

- Main regression: **198/198 suites PASS**, **3,632/3,632 checks PASS**.
- Dedicated `XLPP_P1P_LazyFormulaTests`: PASS.
- Dedicated `XLPP_P1Q_CoreHardeningTests`: PASS.
- Focused ASan+UBSan + leak detection on XML/ZIP-streaming/relationship/CFB hardening paths: PASS.
- Standalone public-header target, including the new `StreamingReadOptions`: PASS.
- C API rebuild + `xlpp_capi_smoke`: PASS.
- Staged install plus an external `find_package(XLPP CONFIG)` consumer using `StreamingReadOptions`: PASS.
- Streaming XML now has a 64 MiB default per-buffer/element guard (caller may set 0 to opt out) and rejects truncated matching elements instead of silently returning EOF.
- Streaming OPC relationship resolution normalizes `.`/`..`, rejects package-root escape, URI schemes/backslashes for internal targets, duplicate rIds, dangling worksheet relationships and missing worksheet parts.
- Streaming ZIP reader accepts caller limits for entry count, per-entry uncompressed size, total uncompressed size and package file size.
- CFB reader now rejects short FAT/miniFAT chains relative to declared stream size, overflowed sector addressing, cyclic/invalid directory references and duplicate case-folded paths.

Detailed notes: `docs/P1Q_CORE_HARDENING.md`.

This directory is the complete XL++ source tree through P1P-A. It builds on P1O-A and targets the largest remaining default-cell payloads: inline Style/FormulaMetadata storage, formula-metadata load density and mutable-subobject change tracking. Ordinary source usage is preserved, but `Cell` layout changes again and mutable lazy access may allocate, so native C++ consumers must be rebuilt with this core.

## Included development batches

- P0–P0Z — Preservation, chart, formula dependency and style/theme foundation
- P1A–P1F — Pivot/VBA selective editing and preservation expansion
- P1G–P1I — Password-to-open encryption foundation and hardening
- P1J — Core structural safety, validation and sheet-lifecycle hardening
- P1K — Transactional core hardening, strict save gates and 3-D reference safety
- P1L-A — Atomic output commit, ZIP/streaming integrity hardening and I/O copy reduction
- P1M-A — Worksheet hot-path, geometry-cache, SST ownership and streaming parser hardening
- **P1N-A — Compact cell/style payloads, lazy optional metadata and allocation-free change-count queries**
- **P1O-A — Workbook subsystem decomposition, compact rare cell optionals, hinted bulk append and peak-RSS benchmark telemetry**
- **P1P-A — Lazy Style/FormulaMetadata storage, formula-load density, mutation tracking and cell-density benchmark/regression isolation**
- **P1Q-A — Malformed-input rejection, streaming resource limits, OPC target normalization and strict CFB validation**
- **P1R-A — Transactional materialized load, model resource limits and package relationship/content-type hardening**
- **P1S-A — XLSX templates, schema-correct XY chart generation, combined secondary-axis charts and typed Pivot caches**

## P1P-A verification — 2026-08-14

- Main regression relinked against the final core: **198/198 suites PASS**, **3,632/3,632 checks PASS**.
- Dedicated `XLPP_P1P_LazyFormulaTests`: PASS.
- C API rebuild + `xlpp_capi_smoke`: PASS.
- Standalone public-header target including `CompactValue.h`: PASS.
- Staged install + external `find_package(XLPP CONFIG)` consumer: PASS.
- Focused Clang ASan+UBSan 50,000-cell lazy-value/deep-copy stress: PASS with leak detection/halt-on-error.
- Full ASan+UBSan project builds were attempted, but the remaining monolithic `Workbook.cpp` instrumented compile exceeded this environment's execution window; P1P-A does not claim a full sanitizer-suite pass.
- `Cell` local size probe: **352 B -> 152 B** (~56.8% lower vs P1O-A; ~85.8% lower vs the pre-P1N layout).
- Same-profile GCC `-O1` / 1M numeric cells: peak RSS **633,303,040 -> 441,380,864 B** (~30.3% lower), bulk build **1,717.172 -> 1,065.867 ms** (~37.9% faster), five tracked-change scans **456.608 -> 244.661 ms** (~46.4% faster), Store-save **3,848.857 -> 3,778.813 ms** (~1.8% faster / effectively flat), output bytes identical.
- P1P 200K density probe: default cells ~89.0 MB RSS; fully styled cells ~170.5 MB RSS. The library intentionally retains independent mutable Style value semantics rather than introducing benchmark-driven COW sharing.

Detailed notes: `docs/P1P_LAZY_MODEL_DENSITY.md`.

### P1P-A ABI compatibility

P1P-A changes `Cell` layout from 352 to 152 bytes and mutable lazy style/formula-metadata access can allocate. **Rebuild XL++ and every native C++ consumer/binding together.** Do not mix P1O-A-compiled C++ objects with P1P-A headers/libraries. Opaque-handle C API callers remain source compatible after rebuilding the core/wrapper.

Project version remains `1.1.2`; P1P-A is a development milestone.

## P1O-A verification — 2026-08-14

- Full unit regression: **197/197 suites PASS**
- Assertions/checks: **3,613/3,613 PASS**
- C API rebuild + `xlpp_capi_smoke`: PASS
- Standalone public-header target (including `CompactOptional.h`): PASS
- Installed `find_package(XLPP)` consumer after staged install: PASS
- `Cell` local size probe: **424 B -> 352 B** (~17.0% smaller vs P1N-A; ~67.2% smaller vs the pre-P1N layout)
- `Workbook.cpp`: **724,828 B / 12,282 lines -> 525,433 B / 8,773 lines**
- Same-profile single-TU compile: **~28.46 s / 969,944 KB -> ~22.51 s / 807,508 KB compiler RSS**
- 200K-cell same-profile median bulk build: **~1,267.0 ms -> ~1,109.2 ms** (~12.5% faster)
- 200K-cell external peak RSS: **~145,952 KB -> ~129,792 KB** (~11.1% lower)
- 1M-cell external peak RSS: **~697,604 KB -> ~619,332 KB** (~11.2% lower)
- Empty-sheet extents cache is valid from construction; optimized 1M-cell geometry-query block: **~79.3 ms -> ~0.029 ms** (~2,700x in this workload)
- Clean GCC `-O1 -j2` build peak RSS: **~1,203,644 KB -> ~821,964 KB** (~31.7% lower); clean wall time rose ~13.4% from extra TU front-end overhead, recorded as a P1P build-graph target

Full ASan+UBSan unit regression with leak detection/halt-on-error: **197/197 suites, 3,613/3,613 checks PASS**. Detailed notes: `docs/P1O_CORE_DECOMPOSITION_DENSITY.md`.

### P1O-A ABI compatibility

P1O-A preserves public method signatures/source usage but changes `Cell` object layout. **Rebuild XL++ and every native C++ consumer/binding together.** Do not mix P1N-A-compiled C++ objects with P1O-A headers/libraries. The C API remains opaque-handle/source compatible when rebuilt with the new core.

Project version remains `1.1.2`; P1O-A is a development milestone.

## P1N-A verification — 2026-08-14

- Full unit regression: **196/196 suites PASS**
- Assertions/checks: **3,600/3,600 PASS**
- Focused 50,000-cell ASan+UBSan compact/copy stress: PASS
- Standalone public-header compile check: PASS
- C API rebuild + `xlpp_capi_smoke`: PASS
- `Cell` local size probe: **1,072 B -> 424 B** (~60.4% smaller)
- `Style` local size probe: **616 B -> 184 B** (~70.1% smaller)

Local identical GCC `-O0` P1M-A -> P1N-A benchmark medians (20,000 x 10 / 200,000 cells):

- bulk build: **~1,816.6 ms -> ~1,244.7 ms** (~31.5% faster)
- Store save: **~2,010.6 ms -> ~1,423.0 ms** (~29.2% faster)
- peak RSS: **~270.8 MB -> ~146.0 MB** (~46.1% lower)
- five repeated tracked-change queries: **~997 ms -> ~91 ms** (~10.9x faster) in the focused numeric workload

These are local same-profile development-regression measurements, not universal cross-library performance claims. `sizeof` results are compiler/ABI dependent.

Detailed notes: `docs/P1N_CORE_MEMORY_DENSITY.md`.

### ABI compatibility

P1N-A preserves public method signatures/source usage but intentionally changes C++ value-type layouts (`Cell`, `Style`, `Hyperlink`, `Comment`, and related nested types). **A full rebuild of XL++ and every native consumer/binding is required.** Do not link P1N-A binaries with object files compiled against P1M-A headers.

The source package excludes generated build directories and compiled artifacts. Project version remains `1.1.2`; P1N-A is a development milestone.

## P0W status

- Formula-precedent propagation for simple A1 formulas: implemented for incremental chart-cache dependency detection.
- Stale formula-backed chart caches are preserved and `calcOnSave/fullCalcOnLoad` can be requested instead of inventing calculated values.
- Recursive local/cross-sheet precedent traversal is bounded by `maxFormulaDependencyDepth`.
- Imported chart color-style palettes can be selectively applied to series after theme resolution.
- Baseline: 156/156 unit suites, 2,768 checks passing in the P0W working tree.

## P0X status

- Formula dependency propagation now accepts rectangular 2-D A1 ranges, not only chart-cache-compatible 1-D ranges.
- Whole-column (`A:A`, `$A:$C`) and whole-row (`1:1`, `$2:$4`) precedents are supported for dependency tracking without materializing million-cell chart caches.
- Workbook/local defined names that resolve to A1 references can participate in formula dependency propagation; nested names are resolved with cycle protection.
- Chart cache references can directly use defined names when they resolve to a one-dimensional A1 range.
- Dependency scanning now walks materialized worksheet cells for mutation checks/formula traversal, keeping sparse large-range checks bounded by actual sheet density.
- `ChartCacheSyncReport` exposes formula-reference and defined-name resolution/skip counters plus informational dependency diagnostics.
- Cyclic or dynamic/unresolvable defined-name graphs are preserved and diagnosed rather than causing synchronization failure.
- New P0X regression suite covers 2-D ranges, whole rows/columns, nested defined names, direct named-range chart caches, and cyclic-name diagnostics.

### P0X verification (2026-08-12)

Clean Ninja/GCC verification in the development sandbox:

- Configure: PASS
- Full source-tree build: PASS
- Standalone public-header target: PASS
- C API target: PASS
- Package validator target: PASS
- Unit tests: **157/157 suites PASS**
- Checks: **2,787/2,787 PASS**
- Installed `find_package(XLPP)` consumer: PASS
- `add_subdirectory()` consumer: PASS
- Generated chart package-validator smoke test: PASS (0 graph/content-type/owner-reference errors)

Project version remains `1.1.2`; this package advances the development milestone without changing the public release-version policy.
## P0Y status

- Common Excel structured table references (`Table[Column]`, `#Data`, `#All`, `#Headers`, `#Totals`, `#This Row`, column ranges and row-scoped `[@Column]`) can participate in formula dependency tracking.
- One-dimensional structured references can be materialized directly into chart caches; two-dimensional structured ranges are tracked as precedents without being flattened into chart caches.
- Workbook/local names using bounded reference-form `OFFSET` and `INDEX` are resolved when their geometry is statically known; expressions that require formula calculation remain safely unresolved with diagnostics.
- `ChartCacheSyncReport` exposes structured-reference and dynamic-defined-name visited/resolved/skipped counters.
- Theme `fmtScheme` fill/line/effect/background-fill entries are materialized, with fill entries preserving exact XML child order for correct matrix indexing.
- `Workbook::applyChartThemeStyleMatrix()` applies selected theme fill/line matrix entries to imported series and resolves `phClr` through chart color-style/theme fallbacks.
- Effect-style inspection materializes outer shadow, glow and soft-edge geometry/color metadata without claiming full Office chart-style rule targeting.
- P0Y regression coverage includes custom reordered theme fills to guard order-sensitive style indices.

### P0Y verification (2026-08-12)

- Unit tests: **158/158 suites PASS**
- Checks: **2,840/2,840 PASS**
- Clean full source-tree build, public-header checks, C API, package validator, installed `find_package` consumer and `add_subdirectory()` consumer: PASS.
- Generated chart package-validator smoke test: PASS with 0 relationship-syntax, duplicate-ID, dangling-relationship, orphaned-part, content-type, or owner-reference errors.

Project version remains `1.1.2`; P0Y is a development milestone and does not change the public release-version policy.



## P0Z status

- Structured-table references support escaped special header characters, including literal `]`, `#`, `@` and apostrophes, while worksheet-quote parsing remains independent from table-token escaping.
- Contiguous multi-item selectors (for example `#Headers` + `#Data`) resolve as rectangular dependencies; non-contiguous table areas are retained as diagnostics until an explicit multi-area dependency model is introduced.
- Bounded reference-form dynamic names accept `INDEX(...) : INDEX(...)` range endpoints when each endpoint reduces to one cell on the same worksheet.
- `ChartStyleResources` exposes parsed Office chart-style target rules and marker layout, including line/fill/effect references, line-width scale, font reference metadata and explicit `spPr` overrides.
- Chart color-style entries and color-transform lists retain source XML order, preventing `styleClr=auto` and transform pipelines from being silently re-indexed/reordered.
- `Workbook::applyChartStyleRules()` materializes supported chart-style fill/line rules onto imported chart/plot/series/marker/axis/auxiliary objects through the preservation-safe selective mutation layer.
- Theme effect materialization adds inner-shadow, reflection and blur geometry/flags. Target-specific effect serialization remains intentionally incomplete and is reported separately from resolved references.
- New regression coverage protects special structured-reference escaping, `INDEX:INDEX` range geometry, mixed color-style order, ordered `styleClr` transforms, `spPr`/`phClr` override behavior, marker layout, save/reload style persistence and expanded theme effects.

### P0Z verification (2026-08-12)

Clean out-of-tree verification in the development sandbox (`Release` configuration with verification flags `-O0 -DNDEBUG -g0`):

- Configure: PASS
- Full source-tree build: PASS
- Standalone public-header target: PASS
- C API target: PASS
- Package validator target: PASS
- Unit tests: **159/159 suites PASS**
- Checks: **2,901/2,901 PASS**
- Installed `find_package(XLPP)` consumer: PASS
- `add_subdirectory()` consumer: PASS
- Generated chart package-validator smoke test: PASS with 0 relationship-syntax, duplicate-ID, dangling-relationship, orphaned-part, content-type, or owner-reference errors.

The aggregate `PackageConsumerTests` CTest driver exceeded the sandbox command timeout while compiling its nested Release consumers; the same two consumers were therefore configured, built and run separately with the verification optimization flags above, and both passed. This is recorded as a harness/compiler-time limitation rather than a library failure.

Project version remains `1.1.2`; P0Z is a development milestone and does not change the public release-version policy.


## P1A status

- Existing worksheet pivots can be discovered through either `<pivotTableParts>` or relationship-only ownership and loaded into `PivotTable`/`PivotCache` models.
- Common cache records and row/column/page/data field metadata are readable/editable; untouched imported pivot parts remain byte-preserved.
- Append-only Pivot generation coexists with imported Pivot OOXML; mutable imported Pivot regeneration suppresses replaced pivot-table roots to keep the OPC object graph valid.
- `PivotCache` adds source-name/refresh/save/enable/missing-item controls; `PivotField` adds compact/outline/default-subtotal/hidden-item state; `PivotFieldReference` adds captions, show-data-as, base field/item and number-format metadata; page-field/layout/style/grand-total settings are modeled.
- VBA authoring now distinguishes Standard/Class/Document modules, preserves host document source on workbook-driven rebuilds, and round-trips module doc-string/read-only/private metadata.
- `VbaProjectInfo` supports project name/description/help/context/constants/project ID plus registered type-library reference name/LibId pairs.
- Generated-project detection is stable across customized project names/GUIDs, allowing later worksheet insertion to keep `SheetN` modules synchronized.
- Advanced Pivot OLAP/grouping/slicer/timeline/calculated-member semantics and VBA UserForms/FRX/signatures/protected-project authoring remain explicit future work.

### P1A verification (2026-08-12)

- Clean configure: PASS
- Static core, examples, standalone public-header checks, C API and package-validator targets: PASS
- Unit tests: **161/161 suites PASS**
- Checks: **2,965/2,965 PASS**
- Installed `find_package(XLPP)` consumer: PASS
- `add_subdirectory()` consumer: PASS
- Combined generated Pivot + source-authored VBA `.xlsm` save/reload smoke: PASS
- Package validator for that smoke workbook: PASS; reachable inventory contains 1 Pivot table / 1 Pivot cache and reports 0 relationship, orphan, content-type or owner-reference errors.

The aggregate nested package-consumer CTest and one-shot all-target build can exceed the sandbox command limit because they recompile the large core/test translation units. The same targets/consumer paths were run separately and passed; this is a harness/compiler-time limitation rather than a library error.

Project version remains `1.1.2`; P1A is a development milestone and does not change the public release-version policy.

## P1B Pivot/VBA development verification — 2026-08-12

P1B extends the Pivot and VBA subsystems while retaining the existing preservation-first package model.

Delivered Pivot coverage:
- opt-in physical PivotCache sharing with compatibility validation;
- calculated cache fields;
- numeric and date/time `fieldGroup`/`rangePr` grouping;
- explicit field subtotal functions;
- imported Pivot mutation/save/reload coverage for grouping and calculated fields.

Delivered VBA coverage:
- stable worksheet `sheetPr@codeName` and document-module identity;
- PROJECT system kind/LCID/invoke-LCID/code-page and module help-context metadata;
- registered type-library, external-project and ActiveX `REFERENCECONTROL` binary reference round-trip.

Final P1B verification:
- **164/164 test suites PASS**
- **3,016/3,016 checks PASS**
- Full source-tree build: PASS with no compiler warning/error in the P1B verification build.
- Standalone public-header target: PASS.
- C API target: PASS.
- Package validator target and examples: PASS.
- Installed `find_package(XLPP)` consumer: PASS.
- `add_subdirectory()` consumer: PASS.
- Combined P1B Pivot + VBA `.xlsm` smoke: save/reload PASS; validator inventory reports 2 PivotTables / 1 shared PivotCache and 0 relationship, orphan, content-type, or owner-reference errors.
- Aggregate `PackageConsumerTests` reached the sandbox command limit while compiling its nested Release consumers; both consumer paths were then run separately with verification optimization flags and passed.

UserForms/FRX/designer storage authoring is explicitly not marked complete; `REFERENCECONTROL` is the prerequisite reference layer only.


## P1C status — PivotChart/filter/selective-cache + VBA Designer/UserForm raw storage

- `Chart` exposes a Pivot source model and `linkPivotTable()`, and chart read/write recognizes DrawingML `c:pivotSource` name/format ID metadata.
- `PivotTable` exposes root `chartFormat`, `chartFormats/chartFormat` entries and lossless raw `pivotArea` subtrees so PivotChart-specific selectors can round-trip without truncation.
- Pivot filter modeling now includes field/type/id/evaluation-order, measure/member-property selectors, text values and a lossless nested `autoFilter` subtree.
- Pivot serialization follows schema-sensitive ordering for `chartFormats`, `pivotTableStyleInfo` and `filters`.
- `Workbook::updateImportedPivotCacheOptions()` selectively patches refresh/cache-retention attributes on an imported physical shared PivotCache while leaving sibling PivotTable XML byte-for-byte unchanged and keeping all loaded models sharing that cache coherent.
- VBA adds `VbaModuleType::Designer`, first-class designer-module source, Package/BaseClass metadata and a recursive raw `VbaDesignerStorage` stream model.
- UserForm/designer root storages are written/read as real CFB storage trees outside the `VBA` storage; nested binary streams such as `f`, `o`, `vbFrame` and arbitrary descendants are preserved/editable byte-for-byte.
- Removing a Designer module also removes its matching root Designer Storage while retaining the project when other user modules remain.
- P1C remains deliberately raw below the Designer Storage boundary: semantic MS-OFORMS control parsing, FRX property editing, arbitrary form/control authoring, signatures and protected-project mutation are not claimed yet.

### P1C verification (2026-08-12)

- Unit tests: **166/166 suites PASS**
- Checks: **3,072/3,072 PASS**
- Full source-tree build, public-header checks, C API, examples and package validator: PASS with no compiler warning/error.
- Installed `find_package(XLPP)` consumer: PASS.
- `add_subdirectory()` consumer: PASS.
- Combined P1C `.xlsm` smoke: PASS with 2 PivotTables sharing 1 physical cache, PivotChart linkage, Pivot filter metadata and a VBA Designer/UserForm storage.
- Package validator: 0 relationship-syntax, duplicate-ID, dangling-relationship, orphaned-part, content-type or owner-reference errors.

### Next Pivot/VBA depth target — P1D

- Pivot: expand selective imported-cache mutation beyond root options, add richer pivot-item/filter semantics and reference-safe PivotChart ownership/update helpers, then validate against more Excel/LibreOffice-produced PivotChart corpora.
- VBA: parse the first useful semantic layer of MS-OFORMS/UserForm properties while continuing to preserve unknown designer streams byte-for-byte; add import fixtures from real Excel-created UserForms and tighten designer storage/reference ownership validation.
- After the native contracts stabilize, pull the new Pivot/VBA APIs through C, C# and Python bindings.

## P1D status — selective Pivot cache-field editing + semantic UserForm Form layer

P1D deepens the same Pivot/VBA focus without weakening the preservation-first package model.

Delivered Pivot coverage:
- selective imported physical `cacheField` mutation for name/caption/formula/number-format/database-field metadata;
- synchronization of patched field metadata across all in-memory PivotTables sharing the same physical cache;
- common Pivot field-item semantics (`x`, `t`, `n`, hidden, show-details, formula, missing) plus raw-XML preservation;
- PivotChart ownership/link validation for missing/ambiguous PivotTable sources and chart-format identity mismatches;
- regression proving selective shared-cache field edits leave two sibling `pivotTable*.xml` parts byte-for-byte unchanged.

Delivered VBA/UserForm coverage:
- semantic inspection of MS-OFORMS Designer `f` streams at the Form-property layer;
- form version/property-mask, colors, IDs/flags, border/scroll/cycle/effect, caption, zoom, picture-layout metadata, draw buffer, displayed/logical size and scroll position;
- safe patching of already-present form properties, including variable-length caption rewrite and compressed-to-UTF-16 conversion;
- preservation of trailing FormStreamData/SiteData and all unrelated Designer/object/nested streams;
- Designer/UserForm ownership validation for duplicate/missing storage/module pairs and malformed Form streams.

### P1D verification (2026-08-12)

- Unit tests: **168/168 suites PASS**
- Checks: **3,122/3,122 PASS**
- Static core: PASS
- Standalone public-header target: PASS
- C API target: PASS
- Package validator target: PASS
- Example writer targets (`XLPP_WriteChart`, `XLPP_WriteImage`, `XLPP_WriteValues`): PASS
- Installed `find_package(XLPP)` consumer: PASS
- `add_subdirectory()` consumer: PASS
- Combined P1D `.xlsm` smoke: PASS after save → reload → selective shared-cache field edit → Unicode UserForm property patch → save → reload.
- Smoke inventory: 1 worksheet, 1 drawing, 1 chart, 2 PivotTables and 1 shared PivotCache.
- Package validator reports 0 relationship-syntax, duplicate-ID, dangling-relationship, orphaned-part, content-type or owner-reference errors.

`ctest` completed `XLPP_UnitTests` successfully; the aggregate `PackageConsumerTests` driver reached the sandbox command limit while compiling nested Release consumers. The installed `find_package` and `add_subdirectory` paths were therefore configured, built and run separately with verification optimization flags and both passed.

**Boundary:** P1D adds semantic editing for the UserForm **Form-level** stream only. Child `FormSiteData`, `OleSiteConcreteControl`, object-stream control records, arbitrary control synthesis, signatures and protected-project mutation are not yet claimed. Unknown payloads continue to be preserved rather than guessed.

Project version remains `1.1.2`; P1D is a development milestone and does not alter the public release-version policy.


## P1E Pivot/VBA development verification — 2026-08-12

P1E deepens preservation-safe editing rather than broadening unrelated workbook features.

Delivered Pivot coverage:
- selective imported Pivot field-item mutation;
- selective imported Pivot-filter mutation, including nested AutoFilter criteria;
- selective physical PivotCache record-value mutation with shared-cache model synchronization.

Delivered VBA/UserForm coverage:
- semantic FormSiteData/OleSite control-site inspection;
- safe site metadata editing for already-materialized properties;
- lossless per-control object-stream slicing from Designer stream `o`.

Final clean-room verification:
- **170/170 test suites PASS**
- **3,177/3,177 checks PASS**
- Static core: PASS.
- Standalone public-header target: PASS.
- C API target: PASS.
- Package validator target: PASS.
- Example writer targets: PASS.
- Installed `find_package(XLPP)` consumer: PASS.
- `add_subdirectory()` consumer: PASS.
- Combined P1E `.xlsm` smoke: PASS after load → selective Pivot item/filter/cache-record edits → Unicode UserForm control-site edit → save → reload.
- Smoke inventory: 1 worksheet, 1 drawing, 1 chart, 2 PivotTables and 1 shared PivotCache.
- Package validator: 0 relationship-syntax, duplicate-ID, dangling-relationship, orphaned-part, content-type or owner-reference errors.
- Clean verification build log: no compiler warning/error.

P1E intentionally does not claim semantic authoring of every MS-OFORMS object stream. Unsupported control bytes remain preserved and are exposed as lossless slices for subsequent control-specific parsers.

Project version remains `1.1.2`; P1E is a development milestone.

## P1F status — Pivot data/page fields + first semantic UserForm object controls

P1F extends the preservation-first Pivot/VBA track into two additional layers.

Delivered Pivot coverage:
- selective imported `dataField` mutation for field index, caption/name, subtotal function, show-data-as mode, base field/item and number-format ID;
- selective imported `pageField` mutation for field index, selected item, hierarchy and display name;
- regression proving these edits preserve the physical PivotCache definition and record parts while the edited PivotTable round-trips correctly.

Delivered VBA/UserForm coverage:
- built-in MSForms control-kind classification from cached class indexes;
- object-slice semantic inspection for `CommandButton` and `Label`;
- safe mutation of already-materialized common object properties, including variable-length caption rewrite, colors, size and selected Label/CommandButton style fields;
- automatic control `cbControl` and site `ObjectStreamSize` repair when an edited object slice changes length;
- lossless preservation of opaque trailing object data, unrelated controls and nested Designer streams;
- type detection/common-header validation for additional built-ins such as TextBox/MorphData without pretending their family-specific layouts are already supported.

### P1F verification (2026-08-12)

- Unit tests: **172/172 suites PASS**
- Checks: **3,207/3,207 PASS**
- Static core: PASS.
- Standalone public-header target: PASS.
- C API target: PASS.
- Package validator and example writers (`XLPP_WriteChart`, `XLPP_WriteImage`, `XLPP_WriteValues`): PASS.
- Installed `find_package(XLPP)` consumer: PASS.
- `add_subdirectory()` consumer: PASS.
- Clean-room build from the release source tree: PASS; verification logs contain no compiler warning/error.
- Generated chart package smoke: PASS with 0 relationship-syntax, duplicate-ID, dangling-relationship, orphaned-part, content-type or owner-reference errors.

**Boundary:** P1F semantically edits `CommandButton` and `Label` object streams only. TextBox/MorphData, list controls, toggle/check/option controls, picture/font StreamData, arbitrary control synthesis, protected VBA projects and signatures remain future depth work.

Project version remains `1.1.2`; P1F is a development milestone and does not change the public release-version policy.



## P1G status — password-to-open encryption

P1G adds real encrypted Office-package support. It is intentionally separate from worksheet/workbook protection.

### Writer
- ECMA-376 Agile `EncryptionInfo` 4.4 outer CFB package.
- AES-256-CBC + SHA-512 profile.
- Password KDF default: 100,000 spins.
- 16-byte salt/block geometry and 4096-byte `EncryptedPackage` segment encryption.
- Password verifier, encrypted intermediate package key and `DataIntegrity` HMAC.
- Generic CFB writer upgraded with DIFAT support so large encrypted workbooks are not limited to the old compact VBA-project FAT size.

### Reader
- Agile AES-256/SHA-512 password packages written by XL++/compatible Office producers.
- Standard CryptoAPI AES-128/192/256 + SHA-1 password packages, including the LibreOffice 3.2 fixture in `tests/fixtures/encryption`.
- Unicode UTF-8 API passwords converted to UTF-16LE for Office password derivation.
- Wrong-password rejection and Agile HMAC tamper detection.

### Platform/build integration
- Windows: built-in CNG/BCrypt crypto backend; direct `XL++.sln` source project includes `OfficeCrypto.cpp`.
- Non-Windows: OpenSSL Crypto via CMake/package export.
- Basic password APIs are exposed through native C++, C API, C#, and Python binding source.
- C#/Python binding regression source is included; the current Linux sandbox did not have the .NET SDK or pybind11 toolchain available, so runtime verification for those two binding layers remains CI/toolchain work.
- C API encryption smoke test is registered whenever `XLPP_BUILD_TESTS=ON`.

### P1G verification (2026-08-12)
- Unit suites: **175/175 PASS**.
- Checks: **3,229/3,229 PASS**.
- C API password-encryption smoke: **PASS**.
- Standalone public-header target: **PASS**.
- Package validator/examples: **PASS**.
- Installed `find_package(XLPP)` consumer: **PASS**.
- `add_subdirectory()` consumer: **PASS**.
- Independent interoperability: LibreOffice headless opens the final XL++ Agile AES-256/SHA-512 file and reads the expected cell values; XL++ opens a LibreOffice-produced Standard AES-128/SHA-1 password-encrypted fixture.
- Large encrypted raw package test: >8 MiB payload round-trips through 4096-byte Agile segmentation and CFB DIFAT.

**Boundary:** P1G does not yet write Standard Encryption, support RC4, certificate key-encryptors, extensible encryption, IRM, or every alternate Agile cipher/hash profile. Public version remains `1.1.2`.


## P1H status — encryption profile matrix, compatibility writer and hardening

Delivered:
- Agile AES-128/192/256-CBC reader/writer with SHA-1/SHA-256/SHA-384/SHA-512; default remains AES-256/SHA-512.
- Standard CryptoAPI AES-128/192/256 + SHA-1 writer, complementing the P1G reader.
- Password-free `PackageEncryptionInfo` inspection API.
- `LoadOptions` resource guards for Agile spin count and decrypted inner-package size.
- Algorithm-generic Windows BCrypt/OpenSSL crypto path, constant-time comparison independent of OpenSSL, and best-effort key-buffer zeroization.
- C API extended writer/profile functions; C#/Python binding source exposes the new profile/options surface.
- 12-profile Agile matrix, Standard 128/192/256 writer regression, public non-default Agile round-trip and hostile-input guard regression.

Verification baseline: **178/178 suites, 3,336/3,336 checks PASS**. Independent LibreOffice interoperability is confirmed for the default Agile AES-256/SHA-512 output and Standard AES-128 compatibility output. Non-default Agile combinations are spec-driven/self-tested; this LibreOffice version did not accept several non-default Agile combinations, so XL++ does not claim external LibreOffice interoperability for every matrix member.

Boundary: certificate key encryptors, RC4, Extensible Encryption/IRM, password-change without decrypt/re-encrypt, and Excel Desktop CI remain future work.


## P1I status — in-memory encrypted package path, multi-key inspection and policy hardening

Delivered:
- `ZipArchive::saveToBytes()` / `ZipArchive::open(bytes)` memory boundary, retaining ZIP64 and existing read limits.
- Encrypted file load/save no longer writes the decrypted/plain inner OOXML workbook to a temporary `.xlsx`; inner package bytes flow directly between ZIP serialization/parsing and the CFB encryption layer.
- Agile key-encryptor parsing is URI-aware and requires exactly one password key-encryptor instead of assuming the first `<encryptedKey>` belongs to the password path.
- `PackageEncryptionInfo` now reports total/password key-encryptor counts plus decoded certificate key-encryptor metadata. Certificate-chain validation and certificate/private-key decryption are intentionally not claimed.
- New load policies: `maxEncryptionInfoBytes`, `allowStandardEncryption`, and `requireAgileDataIntegrity`, in addition to P1H spin-count and decrypted-size guards.
- C API extended load policy surface and key-encryptor counts; C#/Python binding source updated for P1I inspection/policy fields.
- Negative regression covers duplicate password encryptors, oversized `EncryptionInfo`, missing Agile `DataIntegrity` under strict policy, and Standard-encryption rejection policy.

Verification baseline: **180/180 suites, 3,352/3,352 checks PASS**. Static core, standalone headers, C API smoke, package validator/examples, installed `find_package`, and source `add_subdirectory` consumers pass.

Boundary: P1I does not decrypt with X.509/private keys or emit certificate key-encryptors. It also does not add RC4, Extensible Encryption or IRM.


## P1J completed — core structural safety and semantic validation

- Added workbook-aware insert/delete row/column APIs with A1/range/whole-row/whole-column reference transformation and stable `StructuralEditReport` diagnostics.
- Structural edits now coordinate formulas/formula metadata, hyperlinks, merged/frozen/dimension ranges, AutoFilter/sort state, CF/DV, tables, print settings, charts/caches, Pivot sources/cache schema/records and drawing anchors.
- Added Pivot structural repair for field/record schema and row/column/page/data/filter/group index remapping.
- Added Excel grid and worksheet-name invariants, including XFD/1,048,576 limits and case-insensitive sheet lookup/uniqueness.
- Added reference-safe `Workbook::renameWorksheet()` / `removeWorksheet()` with preservation ownership cleanup for imported descendants.
- Added `Workbook::validateModelIntegrity()` and opt-in `SaveOptions::validateModelBeforeSave`.
- Hardened defined-name creation/lookup for case-insensitive Excel semantics, scope-aware duplicate checks and local-scope bounds.
- C API now contains sheet lifecycle exceptions at the ABI boundary; Python/C# source exposes safe rename.
- Regression baseline: **188/188 suites, 3,491/3,491 checks PASS**.
- P1I → P1J release diff: **32 changed/new files**, with no removed source files.

## P1K status — transactional core hardening

- Structural edits are transactional by default. `StructuralEditOptions::rollbackOnFailure` snapshots worksheet content, defined names and calculation/cache state and restores them in-place if a mutation throws or is cooperatively cancelled. Existing `Worksheet&` identities remain valid after rollback.
- `StructuralEditOptions::cancel` provides cooperative checkpoints for large edits; cancellation raises `StructuralEditCancelled`.
- `StructuralEditOptions::validateResult` rejects and rolls back edits that introduce a new model-integrity error relative to the pre-edit model.
- `StructuralEditReport` now reports post-edit model error/warning counts and references deliberately preserved because their structural semantics are unsupported.
- `SaveOptions` adds `rejectModelWarningsBeforeSave` and `validatePackageBeforeWrite` so production callers can combine semantic model validation with relationship/content-type/owner-graph validation before bytes are written or encrypted.
- Worksheet mutation invariants reject overlapping merged ranges, overlapping table geometry, empty/duplicate table names and case-insensitive table-name collisions; table lookup is case-insensitive.
- Model validation now checks merge/table overlap, table cross-sheet ranges, AutoFilter ranges, conditional-formatting/data-validation `sqref` geometry and the P1J semantic invariants.
- Formula/reference hardening recognizes 3-D worksheet qualifiers such as `Start:End!A1`. Sheet rename/remove safely rewrites/invalidates 3-D endpoints; row/column structural edits preserve 3-D references and emit diagnostics rather than guessing coordinate semantics.
- Negative-reference corpus regression ensures malformed, out-of-grid, R1C1-like, incomplete quoted and external-workbook tokens are preserved rather than misparsed as local A1 references.
- Regression baseline: **194/194 suites, 3,561/3,561 checks PASS**.


## P1L-A status — core I/O stability and performance

- Plain ZIP saves and encrypted Office output now use sibling temporary files plus atomic replacement, preventing cancellation/write failures from truncating a previously valid destination.
- Materialized ZIP open parses directly from `MappedFile`/input spans; the former full-archive copy is removed.
- `saveToBytes()` writes directly into the destination byte vector through a stream buffer instead of `ostringstream -> string -> vector`.
- ZIP parsing validates EOCD/ZIP64 geometry, central/local consistency, entry method/flags, duplicate names, payload bounds, exact inflate sizes and CRCs.
- Streaming ZIP reader is ZIP64-aware and validates headers/bounds once during directory parse, then jumps directly to payload data.
- `ZipEntrySource` has explicit move ownership for zlib state, including rebasing buffered `next_in` state when moved after a partial read.
- `MappedFile` supports empty files, overflow-safe slices and complete native-handle ownership transfer.
- CMake exposes `XLPP_ENABLE_ASAN` / `XLPP_ENABLE_UBSAN`; native performance regression builds no longer require `libxlsxwriter`.
- Remaining P1L priorities are formula/reference tokenization, wider mutation transactions, `Workbook.cpp` decomposition, dedicated parser fuzz targets, Windows/MSVC v145 sanitizer/CI coverage, and deeper chart/Pivot/VBA semantic validation.