# XL++

[![Windows CI](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml)
[![Linux CI](https://github.com/thang199801666/XLPP/actions/workflows/linux-ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/linux-ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![C++20](https://img.shields.io/badge/C%2B%2B-20-blue)](https://en.cppreference.com/w/cpp/20)
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
| SIMD-accelerated XML scanning | Yes | No | No | No | No |
| mmap zero-copy reading | Yes | No | No | No | No |
| Multi-threaded save | Yes | No | No | No | No |

## Performance

Cross-library benchmark sources and the raw GitHub Actions artifacts are
available in [`benchmarks/README.md`](benchmarks/README.md).

### Cross-library reference run

Measured on the local Windows development machine. The GitHub Actions
workflow publishes runner-specific results as the `benchmark-results` artifact.

#### 10K rows × 15 cols

| Library | Write | Read |
|---------|------:|-----:|
| XLPP (Python) | **653.4 ms** | **309.0 ms** |
| openpyxl (Python) | 1081.5 ms | 1092.5 ms |
| XlsxWriter (Python) | 933.3 ms | n/a |
| XLPP (C++) | See CI artifact | See CI artifact |
| libxlsxwriter (C++) | See CI artifact | n/a |
| XLPP (C#) | See CI artifact | See CI artifact |
| ClosedXML (C#) | See CI artifact | See CI artifact |

#### 10K rows × 15 cols (with Lookup table)

| Library | Write | Read |
|---------|------:|-----:|
| XLPP (Python) | pending CI | pending CI |
| openpyxl (Python) | pending CI | pending CI |
| XlsxWriter (Python) | pending CI | n/a |
| XLPP (C++) | pending CI | pending CI |
| libxlsxwriter (C++) | pending CI | n/a |
| XLPP (C#) | pending CI | pending CI |
| ClosedXML (C#) | pending CI | pending CI |

#### 10K rows × 15 cols (with formula)

| Library | Write | Read |
|---------|------:|-----:|
| XLPP (Python) | pending CI | pending CI |
| openpyxl (Python) | pending CI | pending CI |
| XlsxWriter (Python) | pending CI | n/a |
| XLPP (C++) | pending CI | pending CI |
| libxlsxwriter (C++) | pending CI | n/a |
| XLPP (C#) | pending CI | pending CI |
| ClosedXML (C#) | pending CI | pending CI |

The C++ XLPP/libxlsxwriter write comparison is included in the GitHub Actions
artifact; libxlsxwriter is write-only. Timings vary by runner and are
indicative rather than a performance guarantee. `pending CI` in the tables
below means that scenario has not yet been added to the runner, not that the
base benchmark workflow failed.

### Large-file streaming read

This benchmark measures reading rows without materializing the complete
workbook in memory. Throughput is reported as processed cells per second; the
GitHub Actions artifact contains the raw timings and file sizes.

| File | Cells | XLPP (C++) | openpyxl (Python) | ClosedXML (C#) |
|------|------:|-----------:|------------------:|---------------:|
| 100K rows × 10 cols | 1M | pending CI | pending CI | pending CI |
| 500K rows × 10 cols | 5M | pending CI | pending CI | pending CI |
| 1M rows × 10 cols | 10M | pending CI | pending CI | pending CI |

The existing non-streaming XLPP reference is approximately **940K cells/sec**
at 1M cells. It is not used as a streaming result and is shown only as a
baseline until the large-file streaming job publishes measurements.

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

### Python

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
├── tests/                 # 75+ unit test suites
├── docs/                  # Documentation
├── BUILDING.md            # Build instructions
└── ROADMAP.md             # Development roadmap
```

## License

MIT — see [LICENSE](LICENSE)
