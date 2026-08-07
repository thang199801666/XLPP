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
