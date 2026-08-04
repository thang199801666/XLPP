# XL++

[![Windows CI](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml)
[![Linux CI](https://github.com/thang199801666/XLPP/actions/workflows/linux-ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/linux-ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![C++20](https://img.shields.io/badge/C%2B%2B-20-blue)](https://en.cppreference.com/w/cpp/20)
[![API Documentation](https://img.shields.io/badge/API-Doxygen-blue)](https://thang199801666.github.io/XLPP/)
[![PyPI version](https://img.shields.io/pypi/v/xlpp.svg)](https://pypi.org/project/xlpp/)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/xlpp.svg)](https://pypi.org/project/xlpp/)
[![Benchmarks](https://github.com/thang199801666/XLPP/actions/workflows/benchmark.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/benchmark.yml)

**High-performance C++20 Excel `.xlsx` read/write library.**  
One dependency (zlib). SIMD-accelerated parsing. Multi-threaded save. Python & C# bindings.

API reference: [Doxygen documentation](https://thang199801666.github.io/XLPP/)

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

## Features

| Category | Details |
|----------|---------|
| **I/O** | Read/write `.xlsx`, `.xlsm` (macros preserved on round-trip) |
| **Cells** | string, number, bool, date, error, formula (shared/array) |
| **Styles** | font, fill, border, alignment, number format, named styles |
| **Layout** | merge, freeze, row/col dimensions, insert/delete rows/cols |
| **Tables** | named tables with columns & style info |
| **Filters** | autoFilter, sort state, custom/value filters |
| **Validation** | data validation (list, whole, decimal, date, custom) |
| **Conditional** | cell-is & formula rules with differential formatting |
| **Charts** | bar, line, pie, scatter, doughnut, radar, area, bubble |
| **Pivot** | pivot tables with row/column/page/data fields |
| **Images** | PNG, JPEG with cell anchoring |
| **Comments** | legacy comments with author |
| **Hyperlinks** | external & internal targets |
| **Page setup** | orientation, paper size, margins, headers/footers, print area |
| **Protection** | workbook & worksheet protection with granular permissions |
| **Properties** | title, creator, subject, keywords, custom properties |
| **Streaming** | append-only writer & pull-based reader for 100K+ rows |
| **Strict OOXML** | ISO 29500 strict namespaces |
| **ZIP64** | packages exceeding 4 GB |
| **Macro safe** | VBA & custom XML survive load→save |

## Performance

Cross-library benchmark sources and the raw GitHub Actions artifacts are
available in [`benchmarks/README.md`](benchmarks/README.md).

### Cross-library reference run

Measured on the local Windows development machine. The GitHub Actions
workflow publishes runner-specific results as the `benchmark-results` artifact.

| Language | Library | Workload | Write | Read |
|----------|---------|----------|------:|-----:|
| Python | XLPP | 10K × 15 | **653.4 ms** | **309.0 ms** |
| Python | openpyxl | 10K × 15 | 1081.5 ms | 1092.5 ms |
| Python | XlsxWriter | 10K × 15 | 933.3 ms | n/a |
| C# | XLPP | 10K × 10 | **280.7 ms** | n/a |
| C# | ClosedXML | 10K × 10 | 930.2 ms | n/a |

The C++ XLPP/libxlsxwriter write comparison is included in the GitHub Actions
artifact; libxlsxwriter is write-only. Timings vary by runner and are
indicative rather than a performance guarantee.

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
