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


### Completed development note — P0V

- Public chart-cache dependency inspection for owner/source sheet, chart/series, cache kind and resolved A1 range.
- Incremental `synchronizeChangedChartCaches()` skips cache references that do not intersect tracked worksheet cell changes, including mutations through retained `Cell&` references via per-cell mutation revisions.
- Formula source cells without current cached worksheet values preserve matching existing chart-cache points by default.
- Chart style/color-style resource inspection now exposes IDs, method and color entries with theme-based final RGB resolution.
- Theme font placeholders `+mj-lt` and `+mn-lt` resolve to major/minor Latin fonts from `theme1.xml`.
- OpenPyXL and LibreOffice host validation completed; Calc removal of Office chart-style/color-style extension resources on re-save remains documented host normalization.

## P0W completed — Formula dependency propagation and style application foundation

- Follow simple A1 formula precedents during changed-chart-cache synchronization.
- Preserve stale formula-backed caches instead of treating cached worksheet values as newly calculated results.
- Request host recalculation through workbook calculation flags.
- Resolve and selectively apply imported chart color-style palette colors through the theme.

Next: broaden formula dependency grammar/diagnostics and deepen chart-style matrix/effect application without turning XL++ into a full Excel calculation engine.

## P0X completed — Expanded formula dependency grammar and defined-name resolution

- Formula precedent tracking now handles rectangular 2-D A1 regions independently from the intentionally 1-D chart-cache materialization rules.
- Whole-column and whole-row precedents are recognized for incremental chart-cache invalidation while sparse-sheet traversal avoids scanning every coordinate.
- Workbook/local defined names that reduce to A1 references are resolved recursively, including nested names, with bounded cycle detection.
- One-dimensional defined names can now be used directly as chart-cache references and are surfaced as supported dependencies by `chartCacheDependencies()`.
- `ChartCacheSyncReport` now reports formula-reference/defined-name resolution and skip counts plus non-fatal dependency diagnostics.
- Regression baseline: 157/157 suites and 2,787/2,787 checks pass.

Next: add table structured-reference dependency resolution, deepen dynamic defined-name analysis (`INDEX`/`OFFSET`-style names without becoming a full calc engine), and continue chart-style matrix/effect materialization.


### Completed development note — P0Y

- Structured table references now participate in chart-cache dependency resolution, including common `#Data`, `#All`, `#Headers`, `#Totals`, `#This Row`, named-column and column-range selectors.
- Formula dependency tracking resolves row-scoped structured references such as `[@Sales]` against the containing table and formula row.
- Direct chart-cache references can use one-dimensional structured table references; two-dimensional table ranges remain dependency-only rather than being flattened into an invalid cache shape.
- Workbook/local defined names now support bounded reference-form `OFFSET` and `INDEX` expressions when their geometry is statically knowable. Calculation-dependent dynamic expressions remain diagnostic-only and continue to request host recalculation when appropriate.
- Theme `fmtScheme` fill, line, effect and background-fill matrices are materialized for inspection. Fill-style parsing preserves original DrawingML child order so zero-based matrix indices remain semantically correct for custom themes.
- `Workbook::applyChartThemeStyleMatrix()` selectively applies a chosen theme fill/line entry to imported chart series and materializes `phClr` through chart-color-style/theme fallback resolution.
- DrawingML effect-style inspection now materializes outer shadow, glow and soft-edge metadata without pretending to implement the complete Office chart-style targeting engine.
- P0Y regression coverage includes structured/table dependency propagation, bounded/unbounded dynamic names, theme placeholder materialization, non-empty effect styles and order-sensitive fill-style matrices.

### P0Z completed — Escaped structured references, INDEX endpoint ranges and chart-style rule materialization

- Structured references now decode Excel apostrophe escapes for `]`, `#`, `@`, `'` and related special-header grammar without confusing those escapes with quoted worksheet names.
- Contiguous structured item combinations such as `#Headers + #Data` resolve to one region; non-contiguous selections remain explicitly unsupported instead of being flattened into false geometry.
- Reference-form `INDEX(...) : INDEX(...)` endpoint ranges are resolved when both endpoints are statically bounded single cells on the same worksheet. Calculation-dependent endpoint expressions remain diagnostic-only.
- Office chart-style parts now expose parsed per-target rules (`lnRef`, `fillRef`, `effectRef`, `fontRef`, `lineWidthScale`, `spPr` overrides and marker layout) rather than only resource IDs.
- `Workbook::applyChartStyleRules()` maps supported chart-style targets onto existing selective formatting APIs for chart/plot area, series/markers, axes/gridlines, auxiliary lines/bars, trendlines/error bars, data tables and 3-D floor/walls when present.
- Chart color-style entry order and style-color transform order are preserved, and `styleClr=auto` can feed `phClr` materialization during rule application.
- Theme effect materialization now includes inner shadow, reflection and blur in addition to outer shadow, glow and soft edge.
- Regression baseline: 159/159 suites and 2,901/2,901 checks pass.

### P1A completed — Pivot reader/editor deepening and VBA source-project expansion

Following the post-P0Z audit, P1A was deliberately retargeted from additional chart work to the lower-scoring Pivot/VBA subsystems.

- Imported worksheet PivotTables now load common `pivotTableDefinition`, `pivotCacheDefinition` and `pivotCacheRecords` content into the native object model, including relationship-only producer variants without `<pivotTableParts>`.
- Pivot cache source/options, fields and materialized records; row/column/page/data field metadata; hidden items; style/layout/grand-total settings; aggregation/show-data-as/base-item metadata are readable and regeneratable.
- Untouched imported pivots remain preservation-safe; adding a new pivot to a loaded workbook appends generated Pivot OOXML while retaining original pivot bytes and non-colliding cache IDs.
- Mutating an imported pivot opts into modeled regeneration; replaced legacy pivot-table root parts are retired from the package to avoid orphaned object-graph nodes while shared/possibly-shared preserved caches are handled conservatively.
- VBA source generation now supports Standard, Class and workbook/worksheet Document modules with correct type recovery from the PROJECT stream.
- Document-module source is retained when the generated project is rebuilt after worksheet insertion/removal.
- Module doc strings plus read-only/private metadata round-trip through the OVBA dir stream.
- Project name/description/help file/help context/constants/project ID round-trip, and generated-project ownership detection no longer depends on customizable name/GUID metadata.
- Registered type-library references can be added and read back as name/LibId pairs; the minimal stdole/Office baseline is retained and references are deduplicated.
- Regression baseline: 161/161 suites and 2,965/2,965 checks pass.

### Next refinement target — P1B (Pivot/VBA continuation)

- Introduce selective mutation for imported Pivot XML so common edits do not require regenerating unsupported advanced Pivot children.
- Model shared pivot-cache ownership explicitly and retire/replace stale caches only when reference analysis proves they are unshared.
- Add grouping, calculated items/fields, richer filter/item metadata and pivot-chart linkage before slicer/timeline work.
- Extend VBA reference parsing beyond registered type libraries (project/control/original references) and preserve/edit more project-level metadata.
- Start a dedicated UserForm/FRX/designer-stream subsystem while keeping signed/protected external projects byte-preserved unless explicitly replaced.
- Pull P1A Pivot/VBA APIs through the non-native bindings after the native contracts stabilize.

## P1B status — Pivot shared caches/grouping + VBA reference/host identity depth

- Generated PivotTables can opt into one physical shared PivotCache through `PivotCache::sharedCacheKey()`; incompatible cache definitions using the same identity are rejected instead of silently producing an inconsistent package.
- Imported caches retain a stable physical-cache identity, so shared-cache topology survives load/inspect flows and can be regenerated deliberately.
- Pivot cache fields support calculated-field formulas (`cacheField@formula`, `databaseField=0`) with rectangular cached-record geometry preserved when a field is added after records have loaded.
- Pivot field grouping now models `fieldGroup`/`rangePr`, including parent/base field, numeric bins, date/time grouping units, explicit/automatic bounds and group items. Convenience APIs validate numeric intervals and supported date grouping units.
- Row/column Pivot fields expose explicit subtotal functions in addition to default subtotal behavior; common imported subtotal attributes round-trip.
- Worksheet VBA `codeName` is modeled from `sheetPr`, assigned stably for source-generated VBA projects, and used as the identity of worksheet document modules. Removing/reordering sibling worksheets no longer rebinds source to a positional `SheetN` name.
- VBA project metadata now round-trips system kind, LCID, invoke LCID, code page and per-module help context.
- VBA reference modeling covers registered Automation type libraries, external VBA projects (`REFERENCEPROJECT`) and ActiveX/control libraries (`REFERENCECONTROL`) including extended/twiddled LibIds, original type-library GUID and cookie.
- `REFERENCECONTROL` is the interoperability foundation required by Office Forms/ActiveX designer projects; actual UserForms designer storages, FRX/control serialization and form mutation remain future work rather than being claimed as complete.

### P1B verification (2026-08-12)

- Unit tests: **164/164 suites PASS**
- Checks: **3,016/3,016 PASS**
- P1B regression covers two PivotTables sharing one physical cache, incompatible-share rejection, calculated fields, numeric/date grouping, imported-Pivot edit/save/reload, explicit subtotals, worksheet codeName stability, external-project references and ActiveX `REFERENCECONTROL` round-trip.
- Full source-tree build, public-header checks, C API, examples and package validator: PASS with no P1B compiler warning/error.
- Installed `find_package(XLPP)` consumer: PASS.
- `add_subdirectory()` consumer: PASS.
- Combined P1B Pivot + VBA `.xlsm` smoke: PASS; validator sees 2 PivotTables / 1 shared PivotCache with 0 graph/content-type/owner-reference errors.
- Aggregate nested `PackageConsumerTests` exceeded the sandbox command limit during nested Release compilation; both consumers passed when run separately with verification optimization flags.

Project version remains `1.1.2`; P1B is a development milestone and does not change the public release-version policy.

### Next Pivot/VBA depth target — P1C

- Pivot: PivotChart source/linkage model, richer filters/item metadata, shared-cache selective mutation helpers, grouping edge cases and more Excel/LibreOffice-produced cache corpora.
- VBA: preserve/inspect designer-related CFB storages, start a first-class UserForm/FRX raw-storage model, add control-reference/original-reference corpus fixtures, and keep worksheet/workbook document-module identity stable across structural edits.


## P1C completed — PivotChart/filter/selective-cache and VBA Designer/UserForm raw storage

- Added DrawingML PivotChart source read/write (`c:pivotSource`) plus PivotTable `chartFormat`/`chartFormats` modeling and lossless `pivotArea` preservation.
- Added high-level PivotFilter metadata with lossless nested `autoFilter` preservation and corrected schema-sensitive PivotTable child ordering.
- Added `Workbook::updateImportedPivotCacheOptions()` for targeted physical-cache option updates that preserve all sibling PivotTable XML byte-for-byte and synchronize every loaded model sharing the cache.
- Added VBA Designer modules with `PROJECT` Package/BaseClass metadata and module `VB_Base` identity.
- Added recursive root-level Designer Storage read/write so UserForm/registered-designer binary streams can be inspected, replaced or extended without flattening the CFB tree.
- Designer module removal cleans the matching storage, while unrelated project modules/references survive regeneration.
- Regression baseline: 166/166 suites and 3,072/3,072 checks pass; combined PivotChart + shared-cache + UserForm `.xlsm` smoke validates with zero package-graph/content-type/owner-reference errors.

### Next Pivot/VBA depth target — P1D

- Pivot: selective edits for additional imported cache/field metadata, deeper item/filter semantics, stronger PivotChart ownership/link validation and a broader external PivotChart corpus.
- VBA: begin semantic MS-OFORMS/UserForm property decoding/editing on top of the raw storage model, add real Excel UserForm fixtures, and harden designer/reference ownership and malformed-CFB diagnostics.
- Bindings: once the native P1C/P1D contracts settle, expose the Pivot/VBA additions consistently through C, C# and Python.

## P1D completed — Selective Pivot field mutation and semantic UserForm Form-stream editing

- Imported physical PivotCache fields can be patched selectively through `Workbook::updateImportedPivotCacheField()` without regenerating sibling PivotTable definitions that share the cache.
- Cache-field name, caption, formula, number-format ID and `databaseField` metadata are modeled and synchronized across every loaded PivotTable that references the same physical cache identity.
- Pivot field items retain common SpreadsheetML item semantics (`x`, `t`, `n`, `h`, `sd`, `f`, `m`) while preserving their raw XML for unsupported extensions.
- `Workbook::validatePivotChartLinks()` checks PivotChart source ownership, missing/ambiguous PivotTable references and chart-format identity coherence before save pipelines rely on those links.
- UserForm Designer `f` streams now have a first semantic MS-OFORMS layer: form version/property mask, common form properties, displayed/logical size, scroll position and caption can be inspected without decoding unrelated child-control payloads.
- `Workbook::updateVbaUserFormProperties()` safely patches already-materialized form properties, including caption growth and compressed-to-UTF-16 conversion, while preserving trailing FormStreamData/SiteData and sibling Designer streams byte-for-byte.
- `Workbook::validateVbaDesignerProject()` reports duplicate/missing Designer ownership, orphan/missing root storage and malformed UserForm Form streams.
- Regression baseline: **168/168 suites and 3,122/3,122 checks PASS**.

Next P1E: parse FormSiteData/site records and the first useful child-control/object-stream layer (TextBox/Label/CommandButton/CheckBox/ComboBox/ListBox), add property-level editing while retaining raw fallbacks, deepen selective Pivot filter/item/cache-record mutation and validate PivotChart ownership against more Excel/LibreOffice-produced corpora. After those native contracts settle, bring the P1A–P1E Pivot/VBA surface through C, C# and Python bindings.


## P1E status — selective Pivot item/filter/record editing + UserForm control-site layer

P1E continues the Pivot/VBA depth track without falling back to whole-subsystem regeneration.

Delivered Pivot coverage:
- selective mutation of imported `pivotField/items/item` metadata (`x`, `t`, `n`, hidden, show-details, formula and missing flags);
- selective imported `pivotFilter` attribute editing plus explicit nested `autoFilter` replacement/removal;
- selective physical `pivotCacheRecords` cell mutation for missing, number, string, Boolean, error, date/time and shared-item record values;
- synchronization of physical cache-record edits across every loaded PivotTable model that shares the same physical PivotCache;
- regression proving Pivot item/filter edits do not mutate the shared cache definition/records and that physical record edits do not regenerate sibling PivotTable definitions.

Delivered VBA/UserForm coverage:
- semantic parsing of the `FormSiteData` tail and its `OleSiteConcreteControl` records;
- site depth/type/version/property-mask plus Name, Tag, ID, help context, flags, ObjectStreamSize, tab index, CLSID cache index, group ID, position, tooltip, runtime-license key, ControlSource and RowSource metadata;
- safe in-place editing for already-materialized Name/Tag/help/flags/tab/group/position/tooltip/source properties, including variable-length compressed/UTF-16 strings;
- `ObjectStreamSize` mapping onto the Designer `o` stream so each site exposes its lossless object-data slice and offset;
- preservation of `o`, nested Designer/control streams and unsupported site/control payloads.

### P1E verification (2026-08-12)

- Unit tests: **170/170 suites PASS**
- Checks: **3,177/3,177 PASS**
- Core/static, public-header standalone checks, C API, package validator and example writer targets: PASS.
- Installed `find_package(XLPP)` consumer: PASS.
- `add_subdirectory()` consumer: PASS.
- Combined P1E `.xlsm` smoke: PASS with 2 PivotTables sharing 1 physical PivotCache, 1 PivotChart and 1 UserForm control site.
- Package validator reports 0 relationship-syntax, duplicate-ID, dangling-relationship, orphaned-part, content-type or owner-reference errors.
- Clean verification build log contains no compiler warning/error.

**Boundary:** P1E models the UserForm Form and child-site metadata layers but does not yet semantically rewrite the individual object-stream control classes. Unknown control bytes remain intentionally lossless.

### Next Pivot/VBA depth target — P1F

- Pivot: expand selective mutation to richer cache-record/shared-item management, data/page field metadata, grouping/filter edge cases and imported PivotChart corpora from Excel/LibreOffice.
- VBA/UserForms: parse the first object-specific MS-OFORMS control streams (starting with common Label/CommandButton/TextBox-style controls), ClassTable/CLSID mapping and multi-control object-stream validation while retaining raw fallback.
- Add real external Excel-created UserForm/PivotChart corpus fixtures and begin pulling stabilized Pivot/VBA contracts through C, C# and Python bindings.

## P1F completed — Pivot data/page-field editing and first object-specific UserForm controls

- Added preservation-safe imported Pivot `dataField` selective mutation for caption/name, aggregate, show-data-as, base field/item and number format.
- Added preservation-safe imported Pivot `pageField` selective mutation for field/item/hierarchy/name metadata.
- Added built-in MSForms cached-class classification to UserForm control sites.
- Added semantic object-stream inspection and safe editing for `CommandButton` and `Label`, including caption growth/UTF-16 conversion, color/size/style fields, `cbControl` repair and site `ObjectStreamSize` repair.
- Other control families retain raw object slices and validated type/header metadata until their family-specific parsers are implemented.
- Regression baseline: **172/172 suites, 3,207/3,207 checks PASS**.

### Superseded P1G planning note

- VBA/UserForms: implement MorphData/TextBox semantic parsing first, then CheckBox/OptionButton/ToggleButton and list/combo families; add TextProps/font/picture StreamData handling only where binary ownership can be preserved safely.
- Pivot: add selective shared-item mutation/append/remove with reference repair, deeper data/page-field edge cases, richer grouping/filter validation and imported PivotChart corpus tests.
- Engineering: start exposing the stabilized P1A–P1F Pivot/VBA APIs through C/C#/Python while continuing the Workbook.cpp/test-monolith decomposition work.



## P1G completed — Office password-to-open encryption

- Added true encrypted Office-package load/save using outer CFB `EncryptionInfo` + `EncryptedPackage` streams.
- New-file writer uses Agile AES-256-CBC/SHA-512 with password verifier, encrypted intermediate key, HMAC data integrity, 4096-byte package segmentation and a default 100,000-spin KDF.
- Reader supports that Agile profile plus Standard CryptoAPI AES-128/192/256 + SHA-1 password packages.
- Upgraded the internal CFB writer with DIFAT output so large encrypted packages are supported.
- Added Unicode password handling, encrypted stream APIs, password-file detection, wrong-password rejection and ciphertext-integrity regression.
- Windows uses CNG/BCrypt with no third-party crypto dependency; non-Windows uses OpenSSL Crypto.
- Added basic password-to-open exposure to C API, C# and Python binding source.
- Regression baseline: **175/175 suites, 3,229/3,229 checks PASS**.
- Independent interoperability: final XL++ Agile output opens in LibreOffice; LibreOffice Standard AES-128 fixture opens in XL++.

### Next encryption target — P1H

Keep the focus on password-to-open/security rather than returning immediately to Pivot/VBA:

1. Expand Agile reader compatibility to AES-128/192 and SHA-1/SHA-256/SHA-384 variants where permitted by the format.
2. Add a Standard AES writer compatibility mode for ecosystems that still emit/read Standard Encryption.
3. Introduce explicit encryption-profile inspection/reporting (`Agile`/`Standard`, cipher/hash/key bits/spin count) without requiring decryption.
4. Reduce sensitive-key lifetime with secure-memory cleanup and strengthen malformed `EncryptionInfo` resource limits.
5. Remove temporary inner ZIP files by adding an in-memory/streaming ZIP package boundary, especially for large encrypted workbooks.
6. Add a larger external corpus from Office/LibreOffice producers and Windows Excel Desktop open/recovery validation.
7. Add C#/Python runtime tests for the new password APIs in environments where their toolchains are available.


## P1H completed — encryption profile matrix, Standard writer and hardening

- Generalized Agile password encryption to AES-128/192/256-CBC with SHA-1/SHA-256/SHA-384/SHA-512 for read/write.
- Added Standard CryptoAPI AES-128/192/256 + SHA-1 compatibility writer.
- Added password-free encryption profile inspection.
- Added KDF spin-count and decrypted-package-size guards for untrusted encrypted files.
- Removed the P1G Windows/OpenSSL coupling in constant-time comparison and added platform secure-zero helpers for sensitive byte buffers.
- Extended C/C#/Python binding source for profile selection/inspection.
- Regression baseline: **178/178 suites, 3,336/3,336 checks PASS**.
- Independent interoperability: LibreOffice opens XL++ default Agile AES-256/SHA-512 and Standard AES-128 writer output; XL++ reads the LibreOffice Standard AES-128 fixture.

### Next encryption target — P1I

1. Add certificate-key-encryptor read/inspection foundation without weakening password-key handling.
2. Remove temporary inner-ZIP files by introducing an in-memory package serialization boundary for encrypted file/stream saves.
3. Add explicit encryption policy APIs (minimum spins, allowed formats/algorithms) for enterprise callers.
4. Expand external corpus with Microsoft Excel-generated Agile profiles and Windows Excel Desktop open/recovery CI.
5. Add structured error categories for wrong password, integrity failure, unsupported profile and resource-limit rejection.
6. Add dedicated fuzz targets for `EncryptionInfo`, CFB directory/FAT geometry and encrypted-package size fields.
7. Run C#/Python encryption runtime suites in CI toolchains and expose load resource guards through the C ABI if required.


## P1I completed — in-memory encrypted package path and enterprise policy hardening

- Added in-memory ZIP serialization/open so password-protected file load/save no longer materializes the plaintext inner workbook as a temporary `.xlsx`.
- Agile key-encryptor parsing now resolves the password descriptor by URI and rejects malformed descriptors that do not contain exactly one password key-encryptor.
- Added certificate key-encryptor metadata inspection (decoded certificate bytes plus encrypted-key/verifier sizes) while retaining password decryption when certificate entries coexist.
- Added `LoadOptions::maxEncryptionInfoBytes`, `allowStandardEncryption`, and `requireAgileDataIntegrity` for untrusted/enterprise inputs.
- Extended C API/C#/Python source for P1I policy/profile fields.
- Regression baseline: **180/180 suites, 3,352/3,352 checks PASS**.

### Next encryption target — P1J

1. Add structured native encryption exception/error categories (wrong password, integrity failure, unsupported profile, malformed descriptor, policy/resource rejection).
2. Add dedicated fuzz executables/corpora for `EncryptionInfo`, CFB FAT/DIFAT/directory geometry and `EncryptedPackage` size/segment handling.
3. Build certificate-key-encryptor decryption behind an explicit private-key/certificate API; do not silently consult OS certificate stores.
4. Add a certificate-key-encryptor writer only after independent Office/Excel interoperability fixtures are available.
5. Remove the generic stream API's outer temporary-file bridge so encrypted and plain stream workflows can be fully memory/stream native.
6. Add Microsoft Excel Desktop encrypted corpus/open/recovery CI and run C#/Python runtime encryption suites on their native toolchains.
7. Continue sensitive-memory lifetime reduction and document enterprise policy presets.


## P1J completed — core correctness/refactor milestone

The user redirected P1J from encryption-only expansion to overall core hardening. The milestone therefore prioritizes semantic correctness and maintainability over breadth.

Completed:

1. Workbook-aware structural row/column editing with cross-sheet reference repair and `#REF!` invalidation.
2. Dedicated `ReferenceTransformer`, `WorksheetStructuralEdit`, `WorkbookStructuralEdit` and `PivotStructuralEdit` translation units instead of adding the new logic to the existing monolith.
3. Pivot source/cache schema/record/index repair during structural edits.
4. Excel grid bounds and worksheet-name invariants.
5. Reference-safe workbook sheet rename/remove, including imported relationship-owner cleanup.
6. In-memory semantic model validator with stable issue codes.
7. Opt-in `validateModelBeforeSave` gate.
8. Scope-aware/case-insensitive defined-name invariants.
9. C ABI exception containment for sheet lifecycle operations and binding-safe rename exposure.
10. Expanded negative/round-trip regression across the above areas.

### Recommended P1K core target

1. Continue decomposing `Workbook.cpp` and the unit-test monolith into subsystem translation units without public ABI churn.
2. Replace remaining heuristic formula-reference rewriting with a reusable formula/reference AST for structural edits, shared formulas and unions/intersections.
3. Extend structural repair to every modeled anchor/reference-bearing object and add explicit diagnostics for opaque extensions that cannot be transformed losslessly.
4. Add sanitizer/fuzz CI for ZIP/CFB/XML/formula/reference parsers and model-mutator sequences.
5. Add transactional mutation/rollback for multi-object workbook operations so failures cannot leave a half-transformed model.
6. Extend semantic validation to optional strict warning policies and save-time package + model combined validation.
7. Bring structural/model-validation APIs through all bindings and run native Windows/MSVC v145 plus C#/Python CI.


## P1K completed — transactional core hardening and strict save validation

P1K continues the P1J shift from feature breadth to core correctness.

Completed:

1. Transactional workbook structural edits with rollback-on-failure/cancellation while preserving `Worksheet&` object identity.
2. Cooperative structural-edit cancellation checkpoints and post-edit semantic validation that rolls back only when a new model-integrity error is introduced.
3. Strict save policy that can promote model warnings to fatal and validate the fully assembled OPC relationship/content-type/owner graph before writing or encrypting bytes.
4. Stronger worksheet invariants for overlapping merged ranges, overlapping/duplicate tables and case-insensitive table lookup.
5. Expanded model validator coverage for merge/table overlap, AutoFilter, conditional-formatting/data-validation ranges and cross-sheet table geometry.
6. 3-D worksheet-reference recognition. Sheet rename/remove rewrites endpoints safely; structural row/column edits preserve unsupported 3-D coordinate semantics with explicit diagnostics.
7. Malformed/non-A1 reference negative corpus and structural insert/delete composition regression.
8. Release baseline: **194/194 suites, 3,561/3,561 checks PASS** with installed `find_package` and `add_subdirectory` consumers passing.

### Recommended P1L core target

1. Replace the scanner-only reference transformer with a reusable formula/reference AST/token stream that can represent union/intersection, names, structured references, dynamic arrays, 3-D references and external qualifiers without heuristic ambiguity.
2. Extend transaction support from structural edits to workbook-wide sheet lifecycle, Pivot/chart selective mutations and other multi-object mutations.
3. Split `Workbook.cpp` and the unit-test monolith further into reader/writer/preservation/chart/Pivot/VBA/formula test translation units to reduce compile-time risk.
4. Add ASan/UBSan builds and dedicated fuzz targets for ZIP/CFB/XML/reference/formula parsers plus mutation-sequence fuzzing with model validation after every operation.
5. Add Windows/MSVC v145 CI and runtime C#/Python binding suites; keep public API source-compatible while bringing strict-save/transaction options through all bindings.
6. Add a true save transaction/output-commit layer (temporary destination + atomic replace where supported) so filesystem failures cannot replace a valid existing workbook with a partial output.
7. Expand semantic validation to chart/Pivot/VBA ownership and opaque-extension diagnostics while keeping preservation-first behavior.


## P1L in progress — core I/O stability and performance hardening

Completed in P1L-A (2026-08-13):

1. Added a true filesystem output transaction for ZIP and encrypted Office output: sibling temporary write, flush/close validation, then atomic replace where the host OS supports it. Cancellation or write failure leaves the previous destination intact.
2. Removed the full archive-copy path from materialized ZIP open by parsing directly from memory-mapped or caller-owned byte spans.
3. Removed the `ostringstream -> string -> vector` copy chain from `saveToBytes()` with a vector-backed stream buffer and reduced per-entry compression staging copies.
4. Hardened materialized ZIP parsing with CRC, exact inflate size, duplicate-name, method/flag, ZIP64, EOCD geometry, central/local name/header, bounds and decompression-ratio checks.
5. Hardened the streaming ZIP reader with ZIP64-aware 64-bit geometry, validated direct payload offsets, exact output-size/CRC checks, empty mmap support and ownership-correct movable zlib state.
6. Added `XLPP_ENABLE_ASAN` and `XLPP_ENABLE_UBSAN` CMake profiles and focused sanitizer coverage for the changed I/O core.
7. Decoupled the native XL++ performance-regression target from optional `libxlsxwriter`, so benchmark CI remains buildable on machines without the comparison library.
8. Expanded regression to **195 suites / 3,574 checks**, all passing in the non-sanitized full run.

Measured identical Clang `-O2` before/after ZIP microbenchmark medians show ~8.9x faster materialized open and ~1.52x faster `saveToBytes()` for a ~25.2 MB stored-entry workload. These are local regression measurements, not universal cross-library claims.

Still open for later P1L batches: formula/reference AST/token stream; broader transactional mutation; further `Workbook.cpp`/test decomposition; dedicated persistent fuzz targets/corpora; Windows MSVC v145 + bindings runtime CI; chart/Pivot/VBA ownership validation; and large-scale external benchmark runs against libxlsxwriter/openpyxl/XlsxWriter/ClosedXML on controlled hardware.


## P1M-A completed — worksheet hot paths and streaming parser hardening

Completed on 2026-08-13:

1. Added a lazily populated worksheet-extents cache with incremental single-cell updates and structural-edit invalidation, turning repeated geometry queries into O(1) operations after the first scan.
2. Optimized bulk `Worksheet::append()` so sequential row ingestion no longer pays the random-access tracked-cell-set insertion cost for every cell while preserving mutation revision/change tracking.
3. Reworked the default sequential worksheet serializer to iterate the row-major cell map directly instead of materializing O(N-cells) pointer/span work arrays; parallel-row structures are now allocated only when explicitly requested.
4. Reduced save-time shared-string ownership to one stored copy per unique string and folded comment/external-link detection into the existing style/SST scan.
5. Replaced allocation-heavy streaming shared-string element copies with `string_view` scanner slices and conditional entity decoding.
6. Tightened streaming numeric/style/SST index parsing to reject malformed prefixes and out-of-range indexes deterministically.
7. Added a persistent `xlpp_core_hotpath_benchmark` plus regression coverage for extents invalidation, bulk tracking, large parallel-row output and malformed streaming values.
8. Full regression: **195/195 suites, 3,583/3,583 checks PASS**; focused combined ASan+UBSan smoke PASS.

Local same-profile measurements show ~16% lower 300K-cell bulk-build latency, ~15% lower Store-save latency, ~18.5x faster repeated geometry-query loops, ~9.6% faster string-heavy streaming reads and ~20% lower streaming peak RSS versus P1L-A. See `docs/P1M_CORE_PERFORMANCE.md` for workload details and caveats.

### Recommended next core target

1. Reduce per-cell memory footprint/allocator traffic (cell payload, strings/formulas/styles and sparse-map overhead) while retaining stable public handles and source compatibility.
2. Split the very large `Workbook.cpp` into package writer/reader/preservation/feature translation units so Release/O2 compiler time and memory remain scalable.
3. Add persistent fuzz targets/corpora for ZIP/CFB/XML/SST/formula/reference parsers and model mutation sequences.
4. Continue toward a reusable formula/reference token stream/AST for non-heuristic structural edits, including unions/intersections, names, structured references, dynamic arrays and 3-D/external qualifiers.
5. Track both latency and peak RSS in controlled million-cell/native-vs-external performance runs; do not optimize latency at the cost of correctness or unbounded memory.


## P1N-A completed — core memory density and change-tracking hardening

Completed on 2026-08-14:

1. Reduced local `sizeof(Cell)` from 1,072 to 424 bytes by compacting canonical/default style/formula strings, packing coordinates/style-index state and lazily storing rare hyperlink/comment payloads.
2. Reduced local `sizeof(Style)` from 616 to 184 bytes while preserving semantic equality/hash and public source-level APIs.
3. Preserved deep-copy value semantics for compact strings, styles, comments, hyperlinks and formula metadata; added direct regression for non-aliasing copies and reset-to-default behavior.
4. Replaced `trackedCellChangeCount()` temporary-set cloning with an ordered O(N+K), O(1)-auxiliary-memory merge walk.
5. Extended `xlpp_core_hotpath_benchmark` with object-size and tracked-change metrics.
6. Full regression: **196/196 suites, 3,600/3,600 checks PASS**; focused ASan+UBSan stress, public-header check and C API smoke PASS.
7. Source API signatures remain compatible, but the compact value-type layout is an intentional C++ ABI change; all native consumers must rebuild.

Local same-profile GCC `-O0` A/B on 200K cells measured ~46% lower peak RSS, ~31% faster bulk build and ~29% faster Store save than P1M-A. See `docs/P1N_CORE_MEMORY_DENSITY.md` for methodology and caveats.

### Recommended P1O core target

1. Split the ~724 KB / 12.3K-line `Workbook.cpp` into cohesive package-reader, package-writer, preservation, chart/Pivot/VBA and serializer translation units to restore optimized-build scalability.
2. Add controlled million-cell latency + peak-RSS benchmarks and evaluate sparse-container/node-overhead changes only with stable-handle/iterator semantics explicitly tested.
3. Add persistent fuzz targets/corpora for ZIP/CFB/XML/SST/formula/reference parsers and mutation sequences with semantic validation after every mutation.
4. Continue toward a reusable formula/reference token stream/AST for structural edits, names, unions/intersections, structured references, dynamic arrays, 3-D references and external qualifiers.
5. Rebuild and exercise Python/C#/C consumers in Windows/MSVC v145 CI after the intentional P1N C++ ABI-layout change.


## P1O-A completed — core decomposition and density scaling

Completed on 2026-08-14:

1. Split VBA/UserForm, Pivot mutation/validation, chart-cache/reference synchronization and chart-style mutation logic out of the monolithic `Workbook.cpp` into dedicated translation units.
2. Reduced `Workbook.cpp` from roughly 724,828 bytes / 12,282 lines to 525,433 bytes / 8,773 lines; a local same-profile single-TU compile changed from ~28.46 s / 969,944 KB compiler RSS to ~22.51 s / 807,508 KB.
3. Added internal `CompactOptional<T>` and used it for rare rich-text, named-style, hyperlink and comment state while preserving the public `const std::optional<T>&` accessors and deep-copy value semantics.
4. Reduced the local `sizeof(Cell)` from 424 to 352 bytes (~17% vs P1N-A; ~67% vs the pre-P1N layout).
5. Added an exact-end insertion hint to monotonic `Worksheet::append()` map insertion while retaining stable node/reference behavior.
6. Added cross-platform peak-RSS output to `xlpp_core_hotpath_benchmark`.
7. Made the canonical empty-sheet `A1:A1` extent cache valid at construction so `cell()`/`append()` maintain extents incrementally from the first insertion; the optimized 1M-cell geometry block drops from ~79 ms to ~0.03 ms in the local workload.
8. Full non-sanitized regression and full ASan+UBSan regression: **197/197 suites, 3,613/3,613 checks PASS**; C API smoke, standalone public-header target and installed `find_package(XLPP)` consumer PASS.
9. Local same-profile A/B shows lower bulk-build latency and peak RSS at both `-O0` and `-O1`. A clean `-O1 -j2` build reduces peak compiler/build RSS by ~31.7% but increases clean wall time by ~13.4% due to additional TU front-end overhead; recovering that clean-build time is explicitly carried into P1P.

### Recommended P1P core target

1. Continue splitting `Workbook.cpp`: package reader/writer, worksheet serializer and preservation/package-patching logic are the next high-value boundaries.
2. Decompose the large unit-test translation unit by subsystem to reduce compiler memory/time and make sanitizer builds more incremental.
3. Add persistent fuzz targets/corpora for ZIP, CFB, XML, SST, formula/reference parsing and mutation sequences with semantic validation after each operation.
4. Replace remaining heuristic formula/reference rewriting with a reusable token stream/AST supporting unions/intersections, names, structured references, dynamic arrays, 3-D references and external qualifiers.
5. Evaluate sparse-cell container/node overhead only with explicit stable-handle/iterator tests and controlled 1M+ cell latency/RSS gates.
6. Exercise MSVC v145 and Python/C#/C runtime bindings after the P1N/P1O C++ ABI layout changes.


## P1P-A completed — lazy model density and formula/tracking hardening

Completed on 2026-08-14:

1. Added internal deep-copy `CompactValue<T>` storage and moved `Cell` Style/FormulaMetadata to lazy allocation.
2. Reduced the local GCC/libstdc++ `sizeof(Cell)` from 352 B to 152 B (~56.8% vs P1O-A; ~85.8% vs the pre-P1N 1,072-byte layout).
3. Made normal `setFormula()` clear stale shared/array/dynamic metadata while structural reference rewrites preserve metadata through a dedicated text-update path.
4. Changed OOXML formula loading so ordinary formulas remain FormulaMetadata-allocation-free; only real non-default `<f>` metadata materializes storage.
5. Hardened mutable subobject mutation tracking for style/formula metadata/comments/hyperlinks/named style/rich text/raw style index.
6. Added `xlpp_cell_density_benchmark` and a standalone P1P formula-density regression translation unit so new tests do not further grow the legacy monolithic test TU.
7. Main regression: **198/198 suites, 3,632/3,632 checks PASS**; dedicated formula-density package test, C API smoke, standalone public headers, staged installed-package consumer and focused lazy-cell ASan+UBSan stress PASS.
8. A full sanitizer build was attempted but the remaining monolithic `Workbook.cpp` instrumentation compile exceeded this environment's execution window; no full-sanitizer PASS is claimed for P1P-A.

Same-machine GCC `-O1` one-million-cell A/B measured ~30.3% lower peak RSS, ~37.9% faster bulk build and ~46.4% faster tracked-change scans, with Store-save effectively flat/slightly faster. See `docs/P1P_LAZY_MODEL_DENSITY.md` for workload details and caveats.

### Recommended P1Q core target

1. Split package reader/writer, worksheet serializer and preservation/package patching from `Workbook.cpp` to restore incremental sanitizer/optimized build scalability.
2. Continue decomposing the unit-test monolith by subsystem.
3. Add persistent ZIP/CFB/XML/SST/formula/reference/mutation fuzz corpora and CI gates.
4. Introduce a reusable formula/reference token stream/AST for non-heuristic structural edits.
5. Evaluate sparse-cell node overhead and style interning only with stable-handle and mutation-isolation correctness gates.
6. Run MSVC v145 plus C#/Python/C runtime CI after the P1P ABI change.


## P1Q-A completed — core input hardening and streaming guards

Completed on 2026-08-14:

1. Hardened `XmlPullReader` with quote-aware opening-tag termination, truncated-element rejection, source-contract checking, size-overflow checking and a bounded default 64 MiB streaming buffer.
2. Added public `StreamingReadOptions` with ZIP entry-count/per-entry/total/file-size limits plus a configurable XML element-buffer limit; existing one-argument `StreamingWorkbookReader` remains available.
3. Normalized internal OPC worksheet relationship targets relative to `xl/workbook.xml`; `.`/`..` are resolved without permitting escape above package root. Internal URI schemes/backslashes, duplicate relationship IDs, dangling worksheet rIds and missing worksheet parts are rejected.
4. Extended `ZipArchiveReader` so streaming open enforces the same resource-limit concepts as materialized package open.
5. Hardened CFB/VBA compound-file reading: declared stream sizes must be backed by complete FAT/miniFAT chains, sector offset arithmetic is overflow-safe, and malformed/cyclic/duplicate directory paths are rejected.
6. Added a dedicated P1Q malformed-input regression TU covering quoted `>`, truncated/oversized streamed XML, relationship traversal/duplicates/dangling IDs, streaming ZIP limits and truncated CFB stream chains.
7. Main regression remains **198/198 suites, 3,632/3,632 checks PASS**; dedicated P1P/P1Q tests, focused ASan+UBSan, standalone public headers, C API smoke and installed-package consumer all PASS.
8. P1Q-A intentionally does not change `Cell`/`Style` object layout. The P1P ABI rebuild requirement still applies to consumers coming from earlier milestones.

### Recommended P1R core target

1. Split package reader/writer and preservation/worksheet serializer blocks out of the remaining ~526 KB `Workbook.cpp` so full sanitizer and optimized builds become incremental.
2. Add persistent libFuzzer/AFL-compatible targets and seed corpora for ZIP, CFB, XML, SST, formula/reference and structural-mutation inputs.
3. Add explicit XML nesting/depth/entity-expansion policies to the remaining in-memory OOXML scanners/parsers, with per-part diagnostics in lenient load mode.
4. Harden OPC content-type and relationship graph validation globally (not only streaming worksheet binding), including duplicate/ambiguous part-name and relationship semantics.
5. Continue formula/reference tokenizer/AST work so structural edits do not depend on heuristic textual rewriting.
6. Run MSVC v145, Python/C#/C bindings and larger hostile-input corpora as release gates.


## P1R-A completed — transactional materialized load and package hardening

Completed on 2026-08-14:

1. `Workbook::load()` now loads into a candidate and commits only after success, preserving the old workbook on package/XML/model failure.
2. Stream load/save plumbing moved to `WorkbookIO.cpp`; stream input is chunked, bounded by `maxFileBytes`, cancellation-aware and RAII-cleaned on every tested failure path.
3. `LoadOptions` now adds `maxWorksheets`, `maxCells`, `maxSharedStrings` and `maxDefinedNames` so compressed XML cannot create an unbounded materialized model.
4. Critical materialized OOXML numeric fields use exact `from_chars` parsing; trailing garbage, overflow/non-finite values and out-of-range SST/style references are rejected.
5. Materialized XML scanning is quote-aware and lexical-markup aware (comments/CDATA/PI), balances same-name nested elements, and the worksheet cell loop is zero-copy.
6. Added `WorkbookPackageReader.cpp` to centralize relationship/content-type parsing. Duplicate IDs/content-type declarations, wrong owner relationship types, package-root traversal, missing internal targets and general internal URI schemes are rejected while unknown unconsumed relationships remain preservable.
7. Main regression is **198/198 suites, 3,632/3,632 checks PASS**; P1P/P1Q/P1R dedicated suites, C API, standalone headers, installed/source consumers and focused ASan+UBSan all PASS.
8. P1R does not change the P1P `Cell`/`Style` layout.

### Recommended P1S core target

1. Extract the materialized worksheet reader plus style/SST/package metadata readers from the remaining ~527 KiB `Workbook.cpp`.
2. Extract worksheet/package writer and preservation patching so optimized/sanitizer rebuilds do not recompile the whole workbook core.
3. Add persistent libFuzzer/AFL-compatible targets and seed corpora for ZIP, CFB, XML, SST, formula/reference and structural-mutation parsers.
4. Add XML nesting/depth and per-part model-count policies with lenient-mode diagnostics.
5. Continue formula/reference tokenizer/AST work and differential structural-edit fuzzing.
6. Run MSVC v145 plus Python/C#/C runtime CI as release gates.


## P1S-A completed — three-pillar XLSX / Charts / Pivot interoperability

Completed on 2026-08-14:

1. Basic XLSX read/write now models workbook template identity and emits/detects the proper XLTX/XLTM workbook main content types, including generated VBA-backed XLTM templates.
2. Scatter/Bubble generation now uses the correct XY-family schema (`xVal`/`yVal`, plus Bubble `bubbleSize`) and round-trips scatter style and bubble-size references/caches.
3. Generated chart plots can be modeled independently through `addPlot()` / `addSeriesToPlot()`, including Bar+Line-style combined charts with a secondary Y axis and deterministic axis ownership.
4. Pivot cache records now retain explicit physical value kinds (missing, number, string, Boolean, error and date/time); save/load and type-aware shared-item construction no longer rely on text inference when typed records are supplied.
5. Generated Pivot shared-item metadata was corrected for interoperability; an independent openpyxl 3.1.5 consumer accepts the generated cache/Pivot package.
6. Added `XLPP_P1S_ThreePillarTests` as an isolated regression TU covering XLTX/XLTM, Scatter/Bubble/combo-secondary charts and typed Pivot caches.
7. Final main regression remains **198/198 suites, 3,632/3,632 checks PASS**. P1P/P1Q/P1R/P1S, C API, headers, installed/source consumers and focused Clang ASan+UBSan verification pass.
8. Independent openpyxl 3.1.5 host validation recognizes template identity, Scatter/Bubble chart classes, a two-plot Bar+Line secondary-axis combination and the generated PivotTable.

### Recommended P1T three-pillar target

1. Basic XLSX: introduce first-class sheet-kind/order ownership needed for chartsheets; broaden XLTX/XLTM and Excel/LibreOffice differential round-trip corpora.
2. Charts: add chartsheet generation/load/preservation, deepen generated combined-axis ownership and make every classic chart family pass an external-host generation corpus.
3. PivotTables: add typed shared-item compaction/index repair, deeper grouping/filter/page/data-field mutations, broader PivotChart corpus and preservation-safe slicer/timeline foundations.
4. Keep persistent cross-host regression as a release gate: XL++ self round-trip is necessary but no longer sufficient for these three subsystems.

## P1T-A completed — first-class Chartsheets and stable Pivot item identity

Completed on 2026-08-14:

1. Basic XLSX now has an explicit mixed workbook tab model (`WorkbookSheetKind`, descriptors and order APIs) while the legacy `sheetCount()`/`sheetNames()` contract remains worksheet-only.
2. Added first-class `Chartsheet` create/load/save/rename/remove/reorder support, correct OPC content types/relationships/drawing ownership, imported preservation and exclusive package-closure cleanup on removal.
3. `StreamingWorkbookReader` distinguishes row-streamable worksheets from chartsheets and exposes the original mixed workbook tab order; materialized load adds `maxChartsheets`.
4. P1S generated multi-plot/secondary-axis charts work on chart-only sheets. An openpyxl-created Chartsheet fixture round-trips with untouched chartsheet XML byte-preserved.
5. Pivot field item semantics can now bind to typed logical cache values instead of unstable shared-item ordinals. `hideCacheValue()`/`showCacheValue()` and save-time `(kind,value)` resolution repair physical `x` indexes after cache reorder/compaction.
6. Package inventory and validation understand Chartsheets as first-class owners of drawing/chart relationships.
7. The C ABI exposes mixed workbook tabs/Chartsheets and P1S combo/scatter/bubble generation primitives; the C API smoke round-trips a Chartsheet combo.
8. Final main regression remains **198/198 suites, 3,632/3,632 checks PASS**; P1P/P1Q/P1R/P1S/P1T, C API, independent openpyxl 3.1.5 host validation, package validation and focused Clang ASan+UBSan pass.

### Recommended P1U three-pillar target

1. Basic XLSX: model active tab and sheet visibility across mixed Worksheet/Chartsheet order; broaden Excel/LibreOffice differential round-trip corpus and extract sheet/package reader-writer code from `Workbook.cpp`.
2. Charts: deepen Chartsheet page setup/view/protection handling, generated chart-family corpus, secondary-axis crossing/scaling and combo combinations beyond Bar+Line.
3. PivotTables: logical-value helpers for page/filter selections, shared-item append/remove compaction with reference repair, grouping/filter edge cases and broader PivotChart producer corpus.
4. Move the first-class P1S/P1T chart/Pivot surface through Python/C# runtime-tested bindings; C ABI is already advanced in P1T.
5. Keep openpyxl/LibreOffice host checks and package graph validation as mandatory release gates rather than relying on XL++ self-round-trip alone.


## P1U-A completed — template and Chartsheet production semantics

Completed on 2026-08-14:

1. Extended the first-class mixed sheet model with `Visible` / `Hidden` / `VeryHidden` state and active-tab identity across Worksheets and Chartsheets while retaining worksheet-only legacy APIs.
2. Added workbook-view `activeTab` / `firstSheet` read-write preservation and active-index repair across reorder/removal/hiding; model validation rejects invalid active/visibility combinations.
3. Deepened `.xltx` / `.xltm` support so template identity coexists with active Chartsheets, mixed visibility and generated/preserved VBA.
4. Added semantic Chartsheet properties, view/zoom, protection, page margins/setup and odd/even/first-page header/footer handling.
5. Split imported Chartsheet dirtiness into sheet-metadata vs chart-subtree domains: metadata edits patch only the Chartsheet XML and preserve DrawingML/ChartML bytes; chart mutation intentionally regenerates the chart closure.
6. Hardened source-backed repeated saves so neither metadata-only edits nor regenerated chart edits revert to the original imported subtree on a later save.
7. Added C ABI template/active/visibility accessors and exercised them in `xlpp_capi_smoke`.
8. Added an openpyxl-generated advanced Chartsheet template fixture and independent openpyxl 3.1.5 host checks for XLTX, XLTM+VBA, Chartsheet metadata and repeated chart regeneration.
9. Final verification: **198/198 main suites, 3,632/3,632 checks PASS**; P1P–P1U dedicated suites, C API, public headers, installed/source consumers, package validator and focused Clang ASan+UBSan all PASS.

### Recommended P1V template/chartsheet target

1. Extract Chartsheet/package writer and reader blocks from the ~581 KiB / 9.6k-line `Workbook.cpp` before adding substantial breadth.
2. Model additional workbook-window/view geometry and Chartsheet print/drawing-header-footer picture ownership while preserving unknown extensions.
3. Expand corpus gates with Microsoft Excel Desktop and LibreOffice produced `.xltx/.xltm` files, mixed visibility/active-tab states and chart-sheet metadata variants.
4. Add runtime-tested Python and C# wrappers for mixed tab state, templates and Chartsheet metadata.
5. Add differential mutation fuzzing for mixed sheet lifecycle sequences (rename/reorder/hide/remove/save/reload) with semantic/package validation after every step.


## P1V-A completed — template / Chartsheet depth and I/O decomposition

Completed on 2026-08-14:

1. Moved Chartsheet XML parsing/serialization and chart-only drawing relationship helpers into `WorkbookChartsheetIO.cpp`, reducing `Workbook.cpp` from roughly 9,607 to 9,442 lines and narrowing incremental sanitizer/optimized rebuild scope.
2. Extended the shared `PageSetup` model with paper height/width, page order, printer-default behavior, cell-comment/error rendering, DPI, copies and printer-settings relationship identity; both Worksheet and Chartsheet load/save paths round-trip the additions.
3. Extended Chartsheet protection with modern hash descriptor metadata (`algorithmName`, `hashValue`, `saltValue`, `spinCount`) while retaining legacy protection semantics.
4. Added first-class `CustomChartsheetView` modeling for GUID, scale, visible/hidden/veryHidden state, zoom-to-fit and nested margins/page-setup/header-footer metadata. Imported custom-view XML remains preservation-backed until mutated.
5. Added duplicate/missing custom-view GUID validation, incomplete modern-protection diagnostics and unresolved generated printer-settings relationship diagnostics.
6. Fixed a custom-view mutable-accessor state-loss defect found by the new regression suite.
7. Independent openpyxl 3.1.5 host validation accepts generated advanced XLTX/custom-view/worksheet page-setup artifacts; package graph validation reports zero errors.
8. Final main regression remains **198/198 suites, 3,632/3,632 checks PASS**; P1P–P1V dedicated suites, C API, public headers, installed/source consumers and focused Clang ASan+UBSan+leak detection pass.

### Recommended P1W template/chartsheet target

1. Continue decomposition by moving Chartsheet lifecycle/package-closure ownership and the remaining chart/package writer out of `Workbook.cpp` without weakening preservation semantics.
2. Model drawing-header-footer picture (`drawingHF`/`picture`) ownership, printer-settings parts and workbook-window/view geometry as first-class preservation-safe objects.
3. Add Excel Desktop and LibreOffice-produced `.xltx/.xltm`/Chartsheet corpora, including printer settings, custom views, protection variants and mixed hidden/active tab states.
4. Add mutation-sequence fuzzing for mixed Worksheet/Chartsheet lifecycle plus custom-view/page-setup changes, with model/package validation after each save/reload.
5. Bring P1U/P1V advanced template/Chartsheet metadata through runtime-tested Python and C# wrappers after the native contracts stabilize.


## P1W-A completed — Chartsheet auxiliary package ownership hardening

Completed on 2026-08-14:

1. Added first-class opaque Chartsheet printer-settings payload ownership. Generated/replaced settings receive their own `xl/printerSettings/printerSettingsN.bin` part, content type and Chartsheet relationship; imported payloads are loaded and can be replaced or cleared without disturbing unrelated package objects.
2. Fixed Chartsheet header/footer-picture preservation to use the correct `legacyDrawingHF` owner node and retain its VML relationship/part across metadata-only edits.
3. Changed imported chart regeneration so it retires only the prior drawing/chart closure rather than suppressing every child of the Chartsheet; VML header/footer, printer-settings and unknown sibling relationships remain preservation-safe.
4. Extended package graph validation/inventory for Chartsheet `legacyDrawingHF -> vmlDrawing` and `pageSetup -> printerSettings` ownership, making dangling/unowned auxiliary parts visible before release.
5. Added C ABI binary printer-settings set/size/copy/clear operations and exercised embedded-NUL round-trip in the C API smoke test.
6. Added `XLPP_P1W_ChartsheetPackageTests` plus an openpyxl host-check script covering generated/replaced/cleared printer settings, header/footer VML preservation, chart regeneration, repeated save and Chartsheet removal cleanup.
7. Final verification: **198/198 main suites, 3,632/3,632 checks PASS**; P1P–P1W dedicated suites, C API, public headers, installed/source consumers and focused Clang ASan+UBSan+leak detection pass.
8. `Workbook.cpp` remains roughly **579 KiB / 9,547 lines** because the strengthened relationship/closure merge logic still resides in the main TU.

### Recommended P1X template/chartsheet target

1. Extract generic package relationship merge/allocation/closure helpers and Chartsheet package writer/lifecycle code from `Workbook.cpp` before adding another broad object family.
2. Model header/footer image media ownership above the preserved VML layer, including explicit image replacement/removal while retaining unknown VML content.
3. Deepen workbook-window/view geometry and template-specific view state with Excel Desktop/LibreOffice corpus fixtures.
4. Add mutation-sequence fuzzing for mixed Worksheet/Chartsheet lifecycle plus page/custom-view/printer/header-footer operations, validating model and package graph after each save/reload.
5. Bring P1U–P1W template/Chartsheet APIs through runtime-tested Python and C# wrappers after the native ownership contracts stabilize.


## P1X-A completed — Chartsheet writer decomposition and relationship-collision hardening

Completed on 2026-08-14:

1. Extracted the full Chartsheet package-write path into `WorkbookChartsheetPackage.cpp/.h`, including generated/imported ownership, chart/drawing closure replacement, printer-settings emission, preservation relationship merge and repeated-save handling.
2. Reduced `Workbook.cpp` from roughly 9,547 to roughly 9,440 lines without changing public Chartsheet APIs.
3. Added relationship-type-aware generated-owner patching: a generated drawing collision updates only `<drawing>`, while printer-settings collisions update only `<pageSetup>`.
4. Fixed the prior global `r:id` replacement defect that could rewrite preserved `legacyDrawingHF` owner references when an unrelated generated relationship reused the same ID.
5. Added `XLPP_P1X_ChartsheetWriterTests` with a deliberate `rId1` collision between preserved VML and regenerated drawing ownership, plus repeated-save package graph validation.
6. Final verification: **198/198 main suites, 3,632/3,632 checks PASS**; P1P–P1X, C API, public headers, installed/source consumers and focused Clang ASan+UBSan+leak detection pass.

### Recommended P1Y template/chartsheet target

1. Extract reusable OPC relationship merge/allocation/closure helpers so worksheet and Chartsheet writers share one tested implementation rather than local copies.
2. Add explicit header/footer image-media ownership and replacement/removal above the preserved VML layer.
3. Add mixed-sheet lifecycle mutation fuzzing with relationship-ID collision seeds and semantic/package validation after every save/reload.
4. Continue reducing `Workbook.cpp` before adding another broad Chartsheet feature family.
