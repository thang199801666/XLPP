# P0Z-I — Phase 28–37 Refinement / v1.12.0

P0Z-I is a refinement and production-engineering release. It advances the Phase 28–37 roadmap without pretending that the long-term scope of every phase is exhausted. The release concentrates on scalable formula dispatch/recalculation, common Excel filter fidelity, preservation-first inspection of enterprise data parts, ABI/SDK governance, performance/reliability gates, and corpus-driven preservation.

## Phase 28 — Formula Engine 2.0 foundation

- Added an internal `FunctionRegistry` so formula-family dispatch is registered instead of growing a monolithic name switch.
- Added dependency-driven partial recalculation through `CalculationOptions::changedCells`.
- Dirty roots expand through the workbook dependency graph to transitive dependent formulas only.
- `CalculationReport` reports dirty roots and selected formula cells.
- Added typed `CalculationMode::{Automatic,AutomaticExceptDataTables,Manual}` mapping to OOXML calculation mode.
- Existing dynamic arrays, structured references, `LET`, iterative calculation and external reference resolution remain available.

Carry-over: LAMBDA-family functions, full reference-value semantics, DAG worker scheduling and Excel-complete function coverage remain future work.

## Phase 29 — Common Excel completion

`AutoFilter` now round-trips additional Excel filter forms:

- Top/Bottom N and percent filters;
- dynamic date/average filters;
- color filters;
- icon filters;
- date-group items.

Legacy value/custom-filter state remains backward compatible. Regression specifically protects the older `includeBlank` behavior while advanced filter setters enforce mutually exclusive OOXML forms.

## Phase 30 — External data and connections

Added preservation-first inspection APIs:

- `Workbook::inspectExternalData()`;
- external workbook links and referenced sheet/name metadata;
- workbook connection metadata;
- query-table metadata;
- inventory of Power Query/web-query/unknown connection-related parts.

These APIs inspect metadata without regenerating enterprise connection payloads that XL++ does not yet semantically author.

## Phase 31 — Data Model / OLAP research boundary

Added `Workbook::inspectDataModel()` with detection of:

- embedded Data Model parts;
- model relationships;
- OLAP Pivot cache definitions;
- preservation warnings/inventory.

The release is intentionally preservation-first: proprietary Data Model/OLAP payloads are not rewritten from a partial model.

## Phase 32 — API and ABI refinement

- Added public `XLPP/Version.h` version contract.
- C ABI is versioned as **ABI 2**.
- Added `xlpp_c_abi_version()` and stable additive `xlpp_capabilities()` bit negotiation.
- Added C ABI inspection summaries for External Data and Data Model presence.
- Continued the additive ABI rule: existing C structs/functions retain meaning and new functionality is introduced additively.

## Phase 33 — Python SDK foundation

- Python package version is synchronized to `1.12.0` from the root `VERSION` file.
- `Workbook` supports context-manager use and `Workbook.open(...)` style construction.
- Native inspection and dirty-calculation surfaces are exposed in pybind11.
- `cibuildwheel` configuration targets CPython 3.10–3.13 on Linux/Windows/macOS.
- `.github/workflows/pypi-publish.yml` runs the self-contained root build through cibuildwheel and supports trusted PyPI publication on a GitHub release. The release audit also corrected `MANIFEST.in` so recursive XL++ C++ sources, the pybind11 bridge and bundled zlib C sources are present in source distributions.

The release host does not have pybind11 installed, so source/parity/package configuration is verified here but a local Python extension binary is not claimed.

## Phase 34 — .NET SDK foundation

- Owning `Workbook` now uses `SafeHandle`; child wrappers remain documented non-owning views.
- Runtime native-library resolution supports RID-native package layout.
- Added version/capability and enterprise-inspection surface.
- NuGet metadata and `.github/workflows/nuget-publish.yml` are present; the workflow builds the native C ABI, runs managed tests, packs and publishes on release.

The release host does not have the .NET SDK, so managed binary execution remains a CI gate rather than a locally claimed result.

## Phase 35 — Performance engineering

- Added a native benchmark target independent of libxlsxwriter.
- Added `benchmarks/performance_budget.json`.
- Added `tools/benchmark_guard.py` for same-machine baseline/candidate regression checks.
- Performance policy is explicit: compare equivalent scenarios on the same runner rather than treating cross-machine timing as a release contract.

## Phase 36 — Reliability CI 2.0

`.github/workflows/reliability.yml` adds:

- strict-warning builds on Linux and Windows for pull requests/pushes;
- scheduled/manual ASan + UBSan regression;
- scheduled/manual Clang/libFuzzer load/validate/resave fuzzing;
- crash-artifact upload on fuzz failure.

Release-host verification:

- strict native build: PASS;
- native regression: **177/177 suites, 3,225 checks PASS**;
- ASan + UBSan: all **177 suites / 3,225 checks PASS** when executed in three bounded slices;
- Clang 17 libFuzzer: **1,000 runs** from 16 OpenPyXL/LibreOffice seed workbooks, no crash or sanitizer finding.

## Phase 37 — Enterprise corpus foundation

Added an enterprise corpus contract:

- `tests/corpus/enterprise/manifest.schema.json`;
- `tests/corpus/enterprise/manifest.json`;
- `tools/xlpp-corpus-exercise`;
- `tools/enterprise_corpus_runner.py`;
- package-part fingerprints and removed-part detection.

The current seed matrix executes four operations (`noop`, `edit-a1`, `add-sheet`, `copy-sheet`) across chart-heavy, pivot-heavy and LibreOffice enterprise fixtures.

Release result: **16/16 scenarios PASS, 0 unexpected removed parts**.

The corpus found and forced a real preservation fix: an empty `docProps/custom.xml` created by LibreOffice was previously dropped by an otherwise successful no-op save. XL++ now preserves the part, its relationship and content type.

## Additional release hardening

A final package-consumer audit exposed a GCC `-Wsubobject-linkage` warning in `StreamingWorkbookReader` when XL++ was consumed through `add_subdirectory()`. `SheetBinding` was moved from anonymous-namespace linkage into the internal XL++ namespace. Both package consumption modes now build/run cleanly:

- installed package + `find_package(XLPP CONFIG REQUIRED)`;
- source embedding via `add_subdirectory()`.

## Release gates

| Gate | v1.12.0 result |
|---|---|
| Native unit regression | 177/177 suites, 3,225/3,225 checks PASS |
| Strict native build | PASS, no project warnings observed |
| Architecture boundary | PASS |
| Binding parity | PASS |
| Binding manifest | PASS at 1.12.0 |
| C ABI build/link/smoke | PASS, ABI 2 |
| Package consumer: find_package | PASS |
| Package consumer: add_subdirectory | PASS |
| ASan + UBSan regression | 177/177 suites PASS in 3 slices |
| libFuzzer | 1,000 runs, no finding |
| Enterprise corpus | 16/16 scenarios, 0 unexpected removed parts |

## Deliberate boundaries after v1.12.0

P0Z-I does not claim that the entire long-term Phase 28–37 roadmap is finished. High-value carry-over includes:

- Excel Desktop COM/recovery-log execution on a real Windows+Excel runner;
- LAMBDA/MAP/REDUCE/SCAN/BYROW/BYCOL and multithreaded formula scheduling;
- semantic authoring of Power Query/Data Model/OLAP rather than preservation-first inspection;
- broader Python/.NET binary matrices and published-package validation;
- a much larger anonymized enterprise workbook corpus;
- advanced Pivot ecosystem, VBA UserForms/FRX, DrawingML/SmartArt authoring and remaining fidelity work described in the active capability matrix.
