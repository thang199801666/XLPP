# P0Z-G — Technical-Debt Hardening / v1.10.0

P0Z-G is a reliability/maintainability release on top of P0Z-F. It does not claim a new Excel-feature-parity score. The milestone closes technical debt that had remained after the architecture, lifetime, Pivot and chart expansion releases and converts several informal guarantees into build/test gates.

## 1. Strict-warning clean core

The full native `xlpp_static` core now builds cleanly with `XLPP_ENABLE_STRICT_WARNINGS=ON` on the release GCC toolchain: **0 warnings / 0 errors** under `-Wall -Wextra -Wpedantic -Wshadow -Wconversion -Wsign-conversion -Wnon-virtual-dtor -Wold-style-cast -Woverloaded-virtual` plus the existing safety diagnostics promoted to errors.

The cleanup was source-level rather than suppression-level. Formula, encryption, ZIP, VBA and style code were reformatted/typed to eliminate misleading indentation, signed/unsigned conversion noise and unused implementation remnants.

## 2. Power-loss-aware path saves

`SaveOptions::durableWrite` now defaults to `true`.

- atomic path saves flush the same-directory staging file before replacement;
- Windows uses `FlushFileBuffers` and write-through replacement semantics;
- POSIX uses `fsync` on the staging file and, after rename, the parent directory;
- non-atomic path saves also flush the completed destination when durability is enabled;
- stream saves are unchanged.

`durableWrite=false` remains available for latency-sensitive callers. Python exposes `SaveOptions.durable_write`. The stable C ABI struct was deliberately **not** enlarged; C callers retain durable-by-default behavior and can opt out through the additive `xlpp_workbook_save_durable()` function. C# exposes the matching `Save(path, durableWrite)` overload.

This is a best-effort portable durability guarantee at the OS/filesystem API boundary. It cannot promise persistence through storage hardware that ignores flush/barrier requests.

## 3. Worksheet-copy topology correctness

`Workbook::copyWorksheet()` now performs a dependency-aware deep topology copy rather than a plain value copy.

- explicit self-qualified formula references follow the cloned sheet;
- conditional-formatting/data-validation formulas, internal hyperlinks, chart series/error-bar references and Pivot source/location references use the existing rename translator on the clone;
- cross-sheet references continue to target the original external sheet;
- copied Tables receive deterministic workbook-unique identifiers;
- copied PivotTables receive deterministic workbook-unique identifiers;
- worksheet-local DefinedNames are cloned to the new `localSheetId` scope and explicit self references are translated;
- chart-cache dependency tracking is invalidated and host recalculation is requested.

A save/reload regression verifies formula, Table, Chart, Pivot and local DefinedName topology after the copy.

## 4. Pivot/Chart model validation

Workbook validation was extended from basic workbook/table/Pivot-name checks to the effective state that the serializers actually materialize.

Pivot validation covers source worksheet/range resolution, inferred lazy cache fields, explicit cache record widths, row/column/page/data-field resolution, axes, sorting, item indexes, numeric grouping, filters, measure/base fields and Top-N constraints. Empty Pivot location/source metadata remains valid when the writer has an Excel-compatible effective default.

Chart validation covers positive dimensions, plot/series spans, Histogram bins, Bubble scale, series/title/category/value/bubble-size/error-bar references, missing worksheets and malformed A1 ranges. External workbook references are intentionally allowed.

The validator therefore catches corrupt modeled output without rejecting the supported auto-cache/default-location workflows.

## 5. Continuous fuzzing and sanitizers

P0Z-G adds `XLPP_BUILD_FUZZERS=ON` and `tests/fuzz/XLPP_WorkbookLoadFuzzer.cpp` for Clang/libFuzzer. The harness performs bounded lenient package loading, semantic validation and second-stage reserialization, and is intended to run with ASan/UBSan.

Release verification includes:

- the existing deterministic ZIP/workbook mutation corpus;
- a Clang 17 libFuzzer build with ASan + UBSan;
- **500 libFuzzer iterations** seeded by real OpenPyXL and LibreOffice fixtures with no crash or sanitizer finding;
- the complete native suite under ASan + UBSan with no finding.

Example continuous run:

```bash
CC=clang CXX=clang++ cmake -S . -B build-fuzz \
  -DXLPP_BUILD_FUZZERS=ON -DXLPP_BUILD_TESTS=OFF \
  -DXLPP_BUILD_SAMPLES=OFF -DXLPP_BUILD_TOOLS=OFF
cmake --build build-fuzz --target XLPP_WorkbookLoadFuzzer
./build-fuzz/tests/fuzz/XLPP_WorkbookLoadFuzzer corpus/ -max_total_time=60
```

## 6. Architecture gates

`ArchitectureBoundaryTests` now additionally prevents Formula, Dependencies and Validation from including OOXML/Package implementation layers. This closes the target-level dependency-creep item left by the first P0Z architecture milestone while preserving the single public `XLPP::XLPP` library target.

The older technical-debt notes about an 8k-line `Workbook.cpp`, a monolithic regression translation unit and a central `WorkbookCodec.cpp` are historical: those were closed by P0Z-B through P0Z-D and are no longer current debt.

## Release verification

- Native regression: **174 / 174 suites PASS**.
- Checks: **3,164 / 3,164 PASS**.
- Full strict-warning core build: **0 warnings / 0 errors**.
- Full native regression under ASan + UBSan: **PASS, no findings**.
- Clang/libFuzzer smoke: **500 seeded iterations, no findings**.
- `ArchitectureBoundaryTests`: PASS.
- `BindingParityTests`: PASS.
- `BindingManifestTests`: PASS after v1.10.0 manifest regeneration.
- `XLPP_CApiSmoke`: PASS, including durability opt-out bridge.
- Public standalone-header build gate: PASS.

Microsoft Excel Desktop recovery-log/COM validation is not available on this Linux host, so direct Excel-host validation remains an external release gate. The host also lacks the .NET SDK and pybind11 package required to binary-build the C# and Python wrappers locally; source/manifest/C-ABI parity is gated here, while their binary jobs remain CI responsibilities.

## Remaining debt after P0Z-G

The highest-value remaining work is now narrower:

1. direct Microsoft Excel Desktop compatibility/repair-log corpus and automated host checks;
2. Pivot ecosystem features that are separate subsystems: calculated fields/items, PivotCharts, slicers/timelines and OLAP/Data Model;
3. complete Python/C# binary CI coverage on every supported platform/toolchain;
4. wider long-running fuzz corpora and production-workbook fuzz seeding;
5. worksheet-removal handle lifetime cannot be guaranteed for every surviving `std::deque` reference without a public storage/ABI design change;
6. rarer cross-object OOXML relationships and advanced DrawingML preservation still benefit from real-world corpus expansion.
