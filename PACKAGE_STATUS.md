# XL++ Full Project Package — P0Z-I / v1.12.0

P0Z-I is the Phase 28–37 refinement and production-engineering release on top of P0Z-H. It keeps the existing workbook/chart/pivot/VBA/encryption/streaming surface while adding dependency-driven formula recalculation, advanced AutoFilter fidelity, enterprise data/Data Model inspection, SDK/ABI governance, performance/reliability workflows, and corpus-driven preservation checks.

## Major additions

- Formula Engine 2.0 foundation: `FunctionRegistry`, dirty-root/transitive recalculation and typed calculation modes.
- Advanced AutoFilter round-trip: Top10, dynamic, color, icon and date-group filters.
- Preservation-first `inspectExternalData()` and `inspectDataModel()` APIs.
- C ABI version **2** with additive runtime capability negotiation.
- Python wheel/PyPI workflow and context-manager/open-style API surface; release audit routes cibuildwheel through the self-contained root build and fixes recursive sdist source inclusion.
- .NET `SafeHandle` ownership foundation and NuGet publishing workflow.
- Native performance budget/benchmark regression guard.
- Reliability CI with strict builds, scheduled ASan/UBSan and Clang/libFuzzer.
- Enterprise preservation corpus runner with package-part fingerprints and unexpected-removed-part failure gating.
- Custom document-properties preservation fix discovered by the enterprise corpus.
- `StreamingWorkbookReader` linkage cleanup so `add_subdirectory()` consumers build without the GCC `-Wsubobject-linkage` warning found during release audit.

## Verification

- Native regression: **177 / 177 suites PASS; 3,225 / 3,225 checks PASS**.
- Strict native Release build (`XLPP_ENABLE_STRICT_WARNINGS=ON`): PASS with no project warnings observed.
- Full native ASan + UBSan coverage: **177 / 177 suites PASS; 3,225 checks**, executed in three bounded slices on this host.
- Clang 17 libFuzzer: **1,000 runs** over 16 OpenPyXL/LibreOffice seed workbooks; no crash or sanitizer finding.
- Enterprise corpus: **16 / 16 scenarios PASS; 0 unexpected removed parts**.
- `ArchitectureBoundaryTests`: PASS.
- `BindingParityTests`: PASS.
- `BindingManifestTests`: PASS at **1.12.0**.
- `XLPP_CApiSmoke`: PASS, including C ABI version/capability and enterprise inspection checks.
- Installed-package consumer using `find_package(XLPP CONFIG REQUIRED)`: PASS.
- Source consumer using `add_subdirectory()`: PASS.

## Binding qualification

The C ABI is binary-built and smoke-tested on the release host. Python and C# source surfaces, manifests and package workflows are synchronized to **1.12.0**. This host does not provide pybind11 or the .NET SDK, so local binary wheel/NuGet execution is not claimed; the repository workflows are the intended cross-platform binary gates.

## Deliberate boundaries

P0Z-I does not claim complete semantic authoring for Power Query, Data Model/OLAP, PivotCharts/slicers/timelines, VBA UserForms/FRX/ActiveX, arbitrary DrawingML/SmartArt, or the full Excel formula catalog. External enterprise payloads in these areas remain preservation-first unless explicitly modeled.

Microsoft Excel Desktop is not installed on this Linux host, so Excel COM/recovery-log validation remains an external Windows+Excel gate rather than a locally fabricated result.

See:

- `docs/P0ZI_PHASE28_TO_37_REFINEMENT.md`
- `docs/CURRENT_CAPABILITIES.md`
- `docs/capabilities.json`
- `docs/COMPATIBILITY_MATRIX.md`
- `tests/FEATURE_COVERAGE.md`
