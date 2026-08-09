# XL++

[![Windows CI](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml)
[![Linux CI](https://github.com/thang199801666/XLPP/actions/workflows/linux-ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/linux-ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![C++20](https://img.shields.io/badge/C%2B%2B-20-blue)](https://en.cppreference.com/w/cpp/20)
[![API Reference](https://img.shields.io/badge/API-Reference-blue)](API_REFERENCE.md)
[![API Documentation](https://img.shields.io/badge/API-Doxygen-blue)](https://thang199801666.github.io/XLPP/)
[![PyPI version](https://img.shields.io/pypi/v/xlpp.svg)](https://pypi.org/project/xlpp/)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/xlpp.svg)](https://pypi.org/project/xlpp/)
[![NuGet version](https://img.shields.io/nuget/v/XLPP.svg)](https://www.nuget.org/packages/XLPP/)
[![Benchmarks](https://github.com/thang199801666/XLPP/actions/workflows/benchmark.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/benchmark.yml)

**High-performance C++20 Excel `.xlsx` read/write library.**  
One dependency (zlib). SIMD-accelerated parsing. Multi-threaded save. Python & C# bindings.

API reference: [API Reference](API_REFERENCE.md) | [Doxygen documentation](https://thang199801666.github.io/XLPP/)

```cpp
#include <XLPP/XLPP.h>
using namespace xlpp;

Workbook wb;
auto& ws = wb.addWorksheet("Sheet1");
ws.cell("A1").setValue(42);
ws.cell("B1").setValue("Hello");
ws.font().setBold(true);
wb.save("output.xlsx");
```

## Feature Comparison

`Partial` means the feature is available with limitations or is not fully
round-trip compatible. `Write only` means the library cannot read an existing
workbook.

| Feature | XLPP | openpyxl (Python) | XlsxWriter (Python) | ClosedXML (C#) | libxlsxwriter (C++) |
|---------|:----:|:-----------------:|:-------------------:|:---------------:|:-------------------:|
| Read and write `.xlsx` | Yes | Yes | Write only | Yes | Write only |
| `.xlsm` / VBA round-trip | Yes | Partial | Partial | Partial | Partial |
| Cell values and formulas | Yes | Yes | Yes | Yes | Yes |
| In-process formula calculation | Partial | No | No | Yes | No |
| Shared and array formulas | Yes | Partial | Partial | Partial | Partial |
| Fonts, fills, borders and number formats | Yes | Yes | Yes | Yes | Yes |
| Merged cells and freeze panes | Yes | Yes | Yes | Yes | Yes |
| Tables and auto-filters | Yes | Yes | Yes | Yes | Yes |
| Data validation | Yes | Yes | Yes | Yes | Yes |
| Conditional formatting | Yes | Yes | Yes | Yes | Yes |
| Charts | Yes | Yes | Yes | Yes | Yes |
| Pivot tables | Yes | Partial | No | Partial | No |
| Images and hyperlinks | Yes | Yes | Yes | Yes | Yes |
| Comments and document properties | Yes | Yes | Yes | Yes | Yes |
| Worksheet and workbook protection | Yes | Yes | Yes | Yes | Yes |
| Password-to-open OOXML encryption | Yes | Partial | No | Partial | No |
| Streaming large worksheets | Yes | Partial | Yes | Partial | Yes |
| Strict OOXML support | Yes | Partial | No | Partial | No |
| ZIP64 workbooks | Yes | Partial | Yes | Partial | Partial |
| Native C++ API | Yes | No | No | No | Yes |
| Python API | Yes | Yes | Yes | No | No |
| C# API | Yes | No | No | Yes | No |

P0M extends imported-chart editing with `dPt` data-point styling, rich chart-title/per-point label runs, gradient/pattern fills, color transforms, and advanced line/custom-dash metadata. As with P0L, these APIs patch only targeted ChartML subtrees and preserve sibling DrawingML/package relationships.

P0N extends preservation-aware chart editing with plot-area/legend manual layouts, axis rich titles, number formats, tick settings, major/minor units, crossing metadata, axis/gridline line formatting, and legend overlay/fill/line formatting. Imported charts retain native `axId` targeting so combined/secondary-axis charts are not flattened during selective edits.

P0O adds native axis scaling (`min`/`max`/`logBase`/orientation), numeric `crossesAt`, built-in/custom display units with optional rich labels, explicit major/minor-gridline lifecycle, and selective chart-area/plot-area fill and line formatting. These changes remain stable-ID/`axId` targeted; unrelated chart XML, sibling drawings and media are preserved.

P0P adds preservation-aware chart auxiliary objects: chart data tables, drop lines, high-low lines, up/down bars, and leader-line containers/formatting at plot or series label scope. Add/remove/edit operations patch only the owning ChartML subtree and preserve sibling DrawingML/media.

P0Q adds first-class StockChart inspection/generation (`Chart::Type::Stock`), generation-time auxiliary objects through `Chart::primaryPlot()`, and `ChartDataTable::textStyle` support for `dTable/txPr`. Both imported and newly generated stock charts support high-low and up/down bars without flattening preserved ChartML.
| SIMD-accelerated XML scanning | Yes | No | No | No | No |
| mmap zero-copy reading | Yes | No | No | No | No |
| Multi-threaded save | Yes | No | No | No | No |

## Performance

Cross-library benchmark sources and the raw GitHub Actions artifacts are
available in [`benchmarks/README.md`](benchmarks/README.md).

### Cross-library reference run

Measured on the Windows GitHub Actions runner, using the latest successful
benchmark run.

#### 10K rows × 15 cols

| Library | Write | Read |
|---------|------:|-----:|
| XLPP (Python) | **810.4 ms** | **402.6 ms** |
| openpyxl (Python) | 2079.9 ms | 1690.8 ms |
| XlsxWriter (Python) | 1384.8 ms | n/a |
| XLPP (C++) | **468.3 ms** | **252.4 ms** |
| libxlsxwriter (C++) | 263.1 ms | n/a |
| XLPP (C#) | **713.7 ms** | **355.2 ms** |
| ClosedXML (C#) | 1266.7 ms | 3945.6 ms |

#### 10K rows × 15 cols (with Lookup table)

| Library | Write | Read |
|---------|------:|-----:|
| XLPP (Python) | **749.9 ms** | **286.8 ms** |
| openpyxl (Python) | 2068.9 ms | 1679.1 ms |
| XlsxWriter (Python) | 940.1 ms | n/a |
| XLPP (C++) | **464.6 ms** | **234.5 ms** |
| libxlsxwriter (C++) | 260.5 ms | n/a |
| XLPP (C#) | **491.3 ms** | **285.7 ms** |
| ClosedXML (C#) | 402.0 ms | 603.7 ms |

#### 10K rows × 15 cols (with formula)

| Library | Write | Read |
|---------|------:|-----:|
| XLPP (Python) | **790.8 ms** | **306.0 ms** |
| openpyxl (Python) | 2092.8 ms | 1653.7 ms |
| XlsxWriter (Python) | 1401.4 ms | n/a |
| XLPP (C++) | **467.2 ms** | **247.8 ms** |
| libxlsxwriter (C++) | 293.4 ms | n/a |
| XLPP (C#) | 492.7 ms | 306.8 ms |
| ClosedXML (C#) | 484.6 ms | 596.4 ms |

The C++ XLPP/libxlsxwriter write comparison is included above; libxlsxwriter
is write-only. `n/a` is used only where a library cannot perform the read
operation, such as read-only XlsxWriter/libxlsxwriter or streaming readers not
provided by the wrapper. Timings vary by runner and are indicative rather than
a performance guarantee.

### Large-file streaming read

This benchmark measures reading rows without materializing the complete
workbook in memory. Values below are read times in milliseconds; the generated
files contain 10 columns.

| File | Cells | XLPP (C++) | openpyxl (Python) | ClosedXML (C#) |
|------|------:|-----------:|------------------:|---------------:|
| 100K rows × 10 cols | 1M | **258.7 ms** | 12655.00 ms | n/a |
| 500K rows × 10 cols | 5M | **1289.8 ms** | 57746.51 ms | n/a |
| 1M rows × 10 cols | 10M | **2562.2 ms** | 121278.43 ms | n/a |

ClosedXML has no streaming reader API, so `n/a` is intentional. The existing
non-streaming XLPP reference is approximately **940K cells/sec** at 1M cells;
it is not used as a streaming result.

## Quick Start

### C++

```cpp
#include <XLPP/XLPP.h>
using namespace xlpp;

// Create
Workbook wb;
auto& ws = wb.addWorksheet("Data");

// Write
ws.cell("A1").setValue("Name");
ws.cell("B1").setValue(42.5);
ws.cell("A1").font().setBold(true);
ws.cell("A1").font().color().setArgb("FFFF0000");
ws.cell("B1").setNumberFormat("#,##0.00");
ws.mergeCells("A2:C2");
ws.freezePanes("A3");

// Bulk append
ws.append({"Alice", 100, 9.99});
ws.append({"Bob",   50,  19.99});

// Save
wb.save("report.xlsx");

// Load
Workbook loaded;
loaded.load("report.xlsx");
std::cout << loaded.worksheet("Data")->cell("A1").stringValueOr("")
          << std::endl;  // "Name"
```

Build: `msbuild XL++.sln /p:Configuration=Release /p:Platform=x64` (VS2022)  
or `cmake -B build -G Ninja && cmake --build build` (any platform)

### Package preservation validator

The CMake build now includes `xlpp-package-validator`, which checks duplicate
relationship IDs, dangling targets, orphaned parts, inconsistent content types,
relationship syntax, and owner references for drawings/images/charts/tables/comments/external links/pivots. It also inventories
preserved DrawingML shapes/text boxes/connectors/groups and can compare package
parts before and after a round trip:

```bash
xlpp-package-validator workbook.xlsx
xlpp-package-validator before.xlsx --compare after.xlsx

# Machine-readable CI output
xlpp-package-validator workbook.xlsx --json
xlpp-package-validator before.xlsx --compare after.xlsx --json
```

See [OPC Preservation Core](docs/PRESERVATION_CORE.md) for current guarantees
and limitations.

### Python

See the complete [Python guide](docs/python.md) for installation, cell values,
formatting, tables, charts, validation, NumPy helpers, and round-trip
preservation details.

```python
import xlpp
from datetime import date

wb = xlpp.Workbook()
ws = wb.add_worksheet("Sales")

ws['A1'].value = "Item"
ws['B1'].value = 100
ws['C1'].value = date(2024, 1, 15)

ws['A1'].font().bold = True
ws.append(["Widget", 50, 19.99])

wb.save("sales.xlsx")

wb2 = xlpp.Workbook()
wb2.load("sales.xlsx")
print(wb2['Sales']['A1'].value)  # "Item"
```

```bash
pip install xlpp
```

Pre-built wheels are published to [PyPI](https://pypi.org/project/xlpp/) for
Windows, macOS (Apple Silicon) and Linux (`manylinux`) on every `v*` tag. To
build from source instead:

```bash
pip install pybind11 setuptools
cd bindings/python && pip install .
```

### C\#

```csharp
using XLPP;

using var wb = new Workbook();
var ws = wb.AddWorksheet("Report");

ws["A1"].Value = "Title";
ws["A1"].Font.SetBold(true);
ws["B1"].Value = 42.5;

wb.Save("report.xlsx");
```


## 3D / Surface Chart Foundation (P0R)

XL++ can now read and preservation-edit Bar3D, Line3D, Area3D, Pie3D, Surface and Surface3D ChartML. Imported charts expose `view3D`, native three-axis structure, and floor/side/back wall formatting. New Bar3D and Surface3D charts can be generated while preserving the same package-validation guarantees used by the 2D chart pipeline.

## Projected Pie, Doughnut & Radar Expansion (P0S)

XL++ now models and generates Pie-of-Pie and Bar-of-Pie (`ofPieChart`) with split controls, second-plot size and series-line formatting. Doughnut/Pie plots expose first-slice angle, Doughnut exposes hole size, and Radar plots expose native `radarStyle` while preserving series markers. Imported charts can be selectively updated by stable ID without rebuilding sibling drawing/chart objects.

## Chart Style, Theme & Series Cache Foundation (P0T)

XL++ now reads workbook theme colors used by charts, preserves chart-style/color-style relationship resources, and models cached chart data (`strCache` / `numCache`) for series titles, categories and values. Imported charts can selectively change the chart style ID or cached series data without rebuilding unrelated ChartML or style resources. Newly generated charts can also serialize first-class series caches.

Theme-aware chart colors expose the source scheme name plus color transforms and can resolve the base scheme color through the workbook theme palette. Direct XL++ output preserves chart-style/color-style parts byte-for-byte when they are not edited.

## Cache Synchronization & Theme Transform Engine (P0U)

XL++ can now rebuild chart `strCache` / `numCache` data directly from the worksheet ranges referenced by each series. `Workbook::synchronizeChartCaches()` supports quoted cross-sheet A1 references, sparse blank cells, title/category/value caches, and numeric format propagation. Unsupported external/2-D/union references are reported rather than rewritten.

The chart theme model also parses major/minor Latin fonts and format-scheme metadata and resolves final RGB/alpha after DrawingML color transforms (`tint`, `shade`, luminance/saturation transforms and alpha transforms). Cache validation helpers expose duplicate-index, ordering and sparse-state diagnostics.

## Dependency-aware Chart Cache Synchronization (P0V)

P0V adds workbook-level dependency snapshots for chart title/category/value references. `ChartCacheSyncOptions::changedReferencesOnly` compares the exact referenced cells, formulas, number formats and workbook date epoch against the previous synchronization and skips unchanged cache rebuilds. Unrelated worksheet edits therefore no longer cause chart-cache rewrites.

```cpp
xlpp::ChartCacheSyncOptions sync;
sync.changedReferencesOnly = true;
auto report = workbook.synchronizeChartCaches(sync);

report.referencesChecked;
report.referencesUnchanged;
report.dependenciesRegistered;
report.dependenciesChanged;
```

Save-time synchronization is also available as an explicit compatibility option. It runs on a private workbook copy, so the caller's in-memory chart objects are not mutated:

```cpp
xlpp::SaveOptions save;
save.synchronizeChartCaches = true;
save.synchronizeChangedChartCachesOnly = true;
workbook.save("report.xlsx", save);
```

The snapshot tracker can be cleared with `resetChartCacheDependencyTracking()`. Unsupported external, structured, union and two-dimensional references retain the P0U skip-and-diagnose behavior.


## P0Y Core Hardening — v1.3.0

P0Y keeps the P0X feature surface stable and hardens the native core for production use. The milestone adds a **strong exception guarantee for `Workbook::load()`**, same-directory **atomic path saves** by default, explicit workbook-model preflight validation, strict Excel A1 coordinate invariants, malformed/duplicate/CRC-checked ZIP handling, direct **ZIP64 streaming reads**, bounded-memory ZIP64 writes for file-backed entries, streaming decompression budgets, dependency-aware worksheet rename/removal, and scoped defined-name correctness.

The hardening pass also fixed two low-level memory-safety defects that normal regression tests did not expose: move ownership of an active streaming inflater and an out-of-bounds central-directory read found by ASan mutation coverage. The final sanitizer gate runs the full native suite under **ASan + UBSan**. Safety-critical compiler diagnostics (`return-type`, uninitialized values and array/string bounds) are promoted to errors in strict-warning mode.

P0Y native regression baseline: **171/171 suites and 3,072/3,072 checks PASS**. See [`docs/P0Y_CORE_HARDENING.md`](docs/P0Y_CORE_HARDENING.md) for exact guarantees, verification and remaining core debt.

## General-purpose Editing Core — P0X 90% Gate

P0X closes the broad 90% engine target with a **90.7/100 weighted general-purpose XLSX editing-core score**. The score covers package/streaming robustness, workbook/common features, formula calculation/dependencies, reference-safe structural editing, charts, pivots, encryption, preservation/interoperability and bindings/tooling. It is not a claim that every Microsoft Excel UI feature or function is implemented. The v1.2.0 closeout regression is **169/169 native suites and 2,992/2,992 checks**, with the full CTest unit/package/C-ABI registry passing.

Major P0X additions include:

- workbook-level transactional row/column structural editing with formula/name/table/filter/CF/DV/print/chart/pivot/drawing/hyperlink reference rewriting and rollback;
- public formula dependency graph;
- dynamic arrays/spill, structured references, `LET`, `INDIRECT`/`OFFSET`, reference introspection, iterative circular calculation and explicit external-workbook resolvers;
- semantic editing/regeneration of imported pivots plus layout/style/cache/data/report-filter options;
- Office Standard AES-128/192/256 + SHA-1 **writer** in addition to the existing Standard reader and Agile AES-256/SHA-512 read/write path;
- expanded Python/C/C# parity, automated C ABI smoke coverage, a buildable sample target and cross-platform CI workflow.

See [`docs/P0X_90_PERCENT_ENGINE.md`](docs/P0X_90_PERCENT_ENGINE.md) for the weighted capability matrix, exact scope and remaining gaps.

## Architecture

```
Application (C++ / Python / C# / C)
    │
    ▼
┌─────────────────────────┐
│  Public API (Workbook,  │
│  Worksheet, Cell, Style)│
├─────────────────────────┤
│  XML Serializer         │  SIMD-accelerated scanning
│  ZIP Engine             │  mmap I/O + ThreadPool
│  Cell Storage           │  uint64_t keys, row-major
├─────────────────────────┤
│  zlib (compression)     │
└─────────────────────────┘
```

## Project structure

```
XLPP/
├── include/XLPP/          # Public headers
├── src/XLPP/              # Implementation (XML, ZIP, Streaming)
├── bindings/
│   ├── python/            # pybind11
│   ├── c/                 # C API DLL
│   └── csharp/            # P/Invoke wrapper
├── tests/                 # 150+ unit test suites
├── docs/                  # Documentation
├── BUILDING.md            # Build instructions
└── ROADMAP.md             # Development roadmap
```

## License

MIT — see [LICENSE](LICENSE)



## P0Z-I Phase 28–37 Refinement — v1.12.0

P0Z-I shifts the next refinement batch toward production engineering: formula-family dispatch now uses a registry and supports dependency-driven dirty recalculation; AutoFilter round-trips Top10/dynamic/color/icon/date-group forms; External Data and Data Model/OLAP receive preservation-first inspection APIs; the C ABI is versioned at ABI 2 with runtime capability negotiation; Python/.NET packaging workflows are present; and benchmark, strict-warning, sanitizer, fuzz and enterprise-corpus gates are first-class project assets. Release verification is **177/177 suites, 3,225/3,225 checks PASS**, ASan+UBSan covers all suites, Clang 17 libFuzzer completes 1,000 seeded runs without a finding, and the enterprise corpus completes **16/16 scenarios with 0 unexpected removed parts**. See [`docs/P0ZI_PHASE28_TO_37_REFINEMENT.md`](docs/P0ZI_PHASE28_TO_37_REFINEMENT.md) and [`docs/CURRENT_CAPABILITIES.md`](docs/CURRENT_CAPABILITIES.md).

## P0Z-H VBA Authoring + Preservation — v1.11.0

P0Z-H turns the experimental VBA source layer into a substantially safer macro-project subsystem. XL++ now authors and reads Standard, Class and host Document modules, supports workbook/worksheet event source, persists stable worksheet VBA code names across reorder/delete/copy, carries `MODULEREADONLY` / `MODULEPRIVATE`, exposes and preserves project Name/Description/Help/Help Context/conditional-compilation constants, exports raw `vbaProject.bin`, and reports signature/source-editability state. Externally supplied or signed VBA projects remain non-destructive/preserve-only: XL++ refuses source/metadata rewrites that could discard designer/reference/protection/signature state. Release verification remains **174/174 suites PASS** with the expanded VBA regression corpus. See [`docs/P0ZH_VBA_AUTHORING_AND_PRESERVATION.md`](docs/P0ZH_VBA_AUTHORING_AND_PRESERVATION.md).

## P0Z-G Technical-Debt Hardening — v1.10.0

P0Z-G keeps the P0Z-F Pivot/Chart feature surface and closes core engineering debt: the full native core is strict-warning clean, path saves gain durable flush/fsync semantics by default, `copyWorksheet()` becomes dependency-aware across formulas/Chart/Pivot/Table/local DefinedNames, Pivot/Chart preflight validation is deeper, architecture dependency gates are stricter, and a Clang/libFuzzer load/validate/resave harness joins the deterministic mutation corpus. Release verification is **174/174 suites and 3,164/3,164 checks PASS**, including a full ASan+UBSan run and a 500-iteration seeded libFuzzer smoke. See [`docs/P0ZG_TECHNICAL_DEBT_HARDENING.md`](docs/P0ZG_TECHNICAL_DEBT_HARDENING.md).

## P0Z-F PivotTable + Complete Excel Chart Families — v1.9.0

P0Z-F expands generated/reloaded non-OLAP PivotTables and brings the native Chart model across the chart families exposed by current Excel. Classic ChartML now has correct XY Scatter/Bubble semantics and generated Combo/multi-axis plots; modern Histogram, Pareto, Box & Whisker, Waterfall, Funnel, Treemap, Sunburst and Filled Map use native ChartEx packaging/read-back. Pivot generation now covers cache records/lifecycle metadata, row/column/page/multiple-data fields, Cartesian items, hidden/repeated labels, numeric/date grouping, filters, subtotals, style/layout options and both classic and Excel-2010+ x14 `Show Values As` modes. C/Python/C# surfaces were advanced with the native model. See [`docs/P0ZF_PIVOT_AND_COMPLETE_CHART_FAMILIES.md`](docs/P0ZF_PIVOT_AND_COMPLETE_CHART_FAMILIES.md).

## P0Z-E Stable Handles + Generated Binding Manifest — v1.8.0

P0Z-E keeps the **90.7/100 scope-defined editing-core feature baseline** while strengthening native/binding lifetime semantics and release governance. Handle-exposed child collections now use stable-address storage, Formula functions compile as separate semantic families, and the core builds through **35 private object modules**. `BindingManifestTests` derives the public Workbook/Worksheet surface directly from native headers and fails when the checked-in Python/C# parity manifest drifts. Python is binary-verified on the release host with **136/136 tests PASS**, including retained-child-handle growth tests; C# carries equivalent managed coverage and remains binary-gated by the .NET CI job because this host has no .NET SDK. See [`docs/P0ZE_STABLE_HANDLES_AND_BINDING_MANIFEST.md`](docs/P0ZE_STABLE_HANDLES_AND_BINDING_MANIFEST.md) and [`bindings/PARITY.md`](bindings/PARITY.md).

## P0Z-D Binding Parity + Formula/Chart Decomposition — v1.7.0

P0Z-D keeps the **90.7/100 scope-defined editing-core feature baseline** while making Python/C# parity a release gate. Formula calculation is decomposed into engine/parser/function/dependency/reference modules, imported ChartML mutation is split into layout/series/plot editors, and the native core now compiles through **31 private object modules**. The C ABI now bridges memory I/O, load callbacks, external formula resolvers, dependency graph, structural/rename services, chart-cache tracking, validation, VBA and scoped defined names. Python is binary-verified on the release host with **135/135 binding tests PASS**; C# has the corresponding managed surface, source/P/Invoke parity checks pass locally, and the new .NET CI job is configured as its binary gate. See [`docs/P0ZD_BINDING_PARITY_AND_FORMULA_DECOMPOSITION.md`](docs/P0ZD_BINDING_PARITY_AND_FORMULA_DECOMPOSITION.md) and [`bindings/PARITY.md`](bindings/PARITY.md).

## P0Z-C Fine-Grained Core Decomposition — v1.6.0

P0Z-C keeps the public API and 90.7/100 scope-defined editing-core feature baseline unchanged while further decomposing the implementation. Chart serialization, pure ChartML mutation and package/drawing orchestration are separate services; Worksheet model code is split into core, structural, image and chart domains; the old regression monolith is split by test domain; ZIP/OPC/XML and Workbook/Worksheet model compilation are separate object modules. The native core now builds through 27 private object targets, and `ArchitectureBoundaryTests` prevents Package→Workbook and Model→OOXML/Package dependency regressions. See [`docs/P0ZC_CORE_DECOMPOSITION.md`](docs/P0ZC_CORE_DECOMPOSITION.md) and [`docs/architecture.md`](docs/architecture.md).

## P0Z-B Core Decomposition — v1.5.0

P0Z-B completes the next architecture step without changing the public API or feature score: the former catch-all `WorkbookCodec.cpp` is removed and replaced by `WorkbookReader`, `WorkbookWriter` and domain-owned Worksheet/Chart/Drawing/Pivot codecs. OOXML now builds as nine subdomain object modules (20 internal object targets across the full core) before aggregation into the same `XLPP::xlpp` library. The release also reduces read/write translation-unit coupling and removes repeated XML rescans in drawing-tag ordering. See [`docs/P0ZB_CORE_DECOMPOSITION.md`](docs/P0ZB_CORE_DECOMPOSITION.md) and [`docs/architecture.md`](docs/architecture.md).
