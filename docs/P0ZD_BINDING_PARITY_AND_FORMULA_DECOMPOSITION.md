# P0Z-D — Binding Parity + Formula/Chart Decomposition / v1.7.0

P0Z-D continues the architecture work from P0Z-C while making Python and C# parity a release gate instead of a best-effort follow-up. The scope-defined native editing-core feature score remains **90.7/100**; this milestone is about implementation decomposition, binding completeness, and verification depth rather than claiming additional Excel Desktop breadth.

## Core decomposition

The formula subsystem no longer compiles as one evaluator translation unit. It is split into explicit responsibilities:

- `Formula/Calculation.cpp` — calculation state, dependency recursion, cache/spill/iteration orchestration (~444 lines).
- `Formula/FormulaParser.cpp` — formula grammar and operator parsing (~369 lines).
- `Formula/FormulaFunctions.cpp` — Excel function catalog and dispatch (~473 lines).
- `Formula/DependencyGraph.cpp` — dependency scanning/query model.
- `Formula/ReferenceTranslator.cpp` — structural/reference translation.
- `FormulaEvaluationContext.h` / `CalculationValue.h` — private interfaces shared by parser/functions without coupling them back to the concrete calculation engine.

CMake compiles these as five private object modules. Together with the P0Z-C domain splits the native core now has **31 private object modules** before aggregation into the same `xlpp_static` / `XLPP::xlpp` target.

Imported ChartML mutation is also decomposed by semantic responsibility:

- `ImportedChartLayoutEditor.cpp`
- `ImportedChartSeriesEditor.cpp`
- `ImportedChartPlotEditor.cpp`
- `ImportedChartXmlEditor.cpp` (small orchestrator)
- `ImportedChartPatcher.cpp` (package/drawing orchestration only)

## C ABI bridge

The C ABI now carries the native services required by the maintained high-level bindings, including:

- full load/save options, atomic/validation save options and memory byte I/O;
- load cancel/progress callbacks;
- calculation options, iterative calculation and external-workbook resolver callbacks;
- structural-edit options/reports and dependency-aware worksheet rename;
- formula dependency graph queries;
- chart-cache synchronization/tracking;
- workbook validation and diagnostics;
- VBA project/module lifecycle;
- scoped defined-name create/lookup and local-sheet-id introspection.

The C header is compiled as strict C11 (`-Wall -Wextra -Werror`) as part of release verification.

## Python parity

The pybind11 surface now covers the native P0X–P0Z services above, including memory I/O, callbacks, external resolvers, structural editing, dependency graph, chart-cache tracking, validation, VBA, scoped names and save safety options.

Read-only Python collection access for tables/images/charts/pivots now uses the native const getters. Merely reading these collections therefore no longer marks a worksheet dirty. `Worksheet.image_count` is exposed as well.

P0Z-D adds `bindings/python/tests/test_native_parity.py` and fixes old tests that were hard-coded to Windows DLL/output paths. On the release host the Python extension was compiled against CPython 3.13 using available pybind11 headers and the complete binding suite passed **135/135 tests**.

## C# parity

`Native` and `Workbook` are partial classes and advanced parity lives in `XlppNet.Advanced.cs` instead of making the legacy wrapper grow indefinitely. The managed surface now includes:

- `LoadOptions` / `SaveOptions` including callbacks and byte-array I/O;
- full calculation options with external-workbook resolver callbacks;
- structural edit and worksheet rename reports;
- dependency graph, chart-cache tracking, validation and diagnostics;
- VBA module lifecycle;
- scoped defined-name create/lookup and nullable `LocalSheetId`.

The release host does not contain `dotnet`, `csc`, `mcs` or `msbuild`, so P0Z-D does **not** claim a local C# binary build. GitHub CI now installs .NET 8 and runs the C# test project against the C ABI built in the same job. A source-level CTest parity gate prevents required symbols from silently disappearing between native/C/Python/C# surfaces.

## Binding parity gate

`cmake/check_binding_parity.cmake` is registered in CTest. It checks required native bridge/Python/C# surfaces and verifies that CMake, VERSION, vcpkg, C ABI, Python and C# package versions all agree on **1.7.0**.

## CI

The dedicated `bindings` GitHub Actions job:

1. installs Python 3.12, pybind11 and pytest;
2. installs .NET 8;
3. builds `xlpp_static`, the Python extension and C ABI from the same source revision;
4. runs the full Python binding tests;
5. runs the C ABI smoke test;
6. runs the C# xUnit project against the just-built native library.

This turns binding parity into a continuous requirement rather than a manual catch-up phase.

## Verification baseline

- native unit suites: **171/171**, **3,072/3,072 checks**;
- public standalone headers: **64/64**;
- architecture boundary gate: **PASS**;
- binding parity gate: **PASS**;
- C ABI runtime smoke: **PASS**;
- C header C11 `-Werror` syntax gate: **PASS**;
- Python binary binding: **135/135 tests PASS** on CPython 3.13 release host;
- strict safety-warning core build: **PASS** (conversion/style warning debt is still tracked separately);
- ASan + UBSan full native suite: **PASS**, sanitizer stderr empty.

## Remaining binding gaps

Low-level preservation/package internals are intentionally native-first and are not required to have identical ergonomic wrappers. C# binary verification depends on the CI SDK job on this host. Any future native high-level Workbook capability must be added to the parity gate and to Python/C# before a release is considered complete.
