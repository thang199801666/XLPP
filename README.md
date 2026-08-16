# XL++

[![Windows CI](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml)
[![Linux CI](https://github.com/thang199801666/XLPP/actions/workflows/linux-ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/linux-ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![C++20](https://img.shields.io/badge/C%2B%2B-20-blue)](https://en.cppreference.com/w/cpp/20)
[![API Reference](https://img.shields.io/badge/API-Reference-blue)](API_REFERENCE.md)
[![PyPI version](https://img.shields.io/pypi/v/xlpp.svg)](https://pypi.org/project/xlpp/)
[![NuGet version](https://img.shields.io/nuget/v/XLPP.svg)](https://www.nuget.org/packages/XLPP/)

**High-performance C++20 Excel library** — reads and writes `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, legacy `.xls`, binary `.xlsb` and plain `.csv`.

Bundled zlib, Windows CNG or OpenSSL crypto, SIMD-accelerated XML parsing, multi-threaded save, byte-preserving round-trips, and full **Python / C# / C** bindings.

```cpp
#include <XLPP/XLPP.h>
using namespace xlpp;

Workbook wb;
auto& ws = wb.addWorksheet("Sheet1");
ws.cell("A1").setValue(42);
ws.cell("B1").setValue("Hello");
ws.cell("A1").font().setBold(true);
ws.append({"Widget", 100, 9.99});
wb.save("output.xlsx");          // .xls / .xlsb / .csv also supported
```

---

## Highlights

- **Formats**: `.xlsx`, `.xlsm` (VBA), `.xltx`/`.xltm` templates, legacy `.xls` (BIFF8), binary `.xlsb` (BIFF12), and `.csv`.
- **Styles**: fonts (14 attributes incl. theme colors + underline style), fills, borders (incl. diagonal), alignment, number formats, named styles.
- **Formatting**: 13 conditional-formatting rule families, auto-filters (value/custom/top10/dynamic/color/icon/date), data validation, tables with totals, sparklines.
- **Charts**: 17 types (2D/3D/ChartEx), trendlines, error bars, data labels, secondary axes, chart style/theme matrix, cache synchronization.
- **Pivot**: shared caches, grouping, filters/items, PivotChart linkage, typed cache values.
- **VBA**: modules, events, code names, UserForms/FRX preservation, project metadata.
- **Robustness**: structural row/column edits (reference rewriting), change tracking, transactional load, byte-preservation of untouched parts, input hardening.
- **Bindings**: Python (`pip install xlpp`), C# (NuGet), C ABI.

---

## Feature Comparison

`Partial` = available with limitations or not fully round-trip compatible. `Write only` = cannot read existing workbooks.

| Feature | XLPP | openpyxl (Py) | XlsxWriter (Py) | ClosedXML (C#) | libxlsxwriter (C++) |
|---|:---:|:---:|:---:|:---:|:---:|
| Read/write `.xlsx` | Yes | Yes | Write only | Yes | Write only |
| Templates `.xltx`/`.xltm` | Yes | Yes | Write only | Partial | Partial |
| Legacy `.xls` (BIFF8) | **Yes** | No | No | No | No |
| Binary `.xlsb` (BIFF12) | **Yes** | No | No | No | No |
| CSV import/export | **Yes** | Partial | Partial | Yes | No |
| `.xlsm` / VBA round-trip | Yes | Partial | Partial | Partial | Partial |
| Cell values & formulas (shared/array/dynamic) | Yes | Partial | Partial | Partial | Partial |
| Fonts, fills, borders, number formats | Yes | Yes | Yes | Yes | Yes |
| Tables, auto-filters | Yes | Yes | Yes | Yes | Yes |
| Conditional formatting (13 rule families) | **Yes** | Yes | Yes | Yes | Yes |
| Charts (17 types + trendline/errorBar/dataLabel) | **Yes** | Yes | Yes | Yes | Yes |
| Pivot tables | Advanced | Preserve/read | No | Partial | No |
| Sparklines | **Yes** | No | No | No | No |
| Images / hyperlinks / comments | Yes | Yes | Yes | Yes | Yes |
| Worksheet/workbook protection | Yes | Yes | Yes | Yes | Yes |
| Password-to-open encryption | Yes | Partial | No | Partial | No |
| Streaming large worksheets | Yes | Partial | Yes | Partial | Yes |
| Strict OOXML / ZIP64 | Yes | Partial | No | Partial | Partial |
| Structural row/col edits (reference rewrite) | **Yes** | No | No | No | No |
| Byte-preservation of untouched parts | **Yes** | No | No | No | No |
| C++ API | Yes | No | No | No | Yes |
| Python API | Yes | Yes | Yes | No | No |
| C# API | Yes | No | No | Yes | No |

---

## Performance

Cross-library benchmark sources live in [`benchmarks/`](benchmarks/README.md). Representative timings (100K cells, mixed types, Windows):

| Operation | XL++ (C++) | XL++ (Python) | xlsxwriter | openpyxl |
|---|--:|--:|--:|--:|
| Write | ~175 ms | ~200 ms | ~360 ms | ~520 ms |
| Read | — | ~180 ms | — | ~475 ms |

XL++ is roughly **1.8× faster than xlsxwriter** and **2.5× faster than openpyxl** — even through the Python binding. See the benchmark workflow for full matrix (including C# / ClosedXML and large-file streaming).

### Streaming read (large files)

| File | Cells | XL++ (C++) | openpyxl |
|------|------:|-----------:|---------:|
| 100K rows × 10 cols | 1M | ~260 ms | ~12,655 ms |
| 500K rows × 10 cols | 5M | ~1,290 ms | ~57,747 ms |

---

## Quick Start

### C++

```cpp
#include <XLPP/XLPP.h>
using namespace xlpp;

Workbook wb;
auto& ws = wb.addWorksheet("Data");
ws.cell("A1").setValue("Name");
ws.cell("B1").setValue(42.5);
ws.cell("A1").font().setBold(true);
ws.mergeCells("A2:C2");
ws.freezePanes("A3");
ws.append({"Alice", 100, 9.99});
wb.save("report.xlsx");

Workbook loaded;
loaded.load("report.xlsx");
// load("legacy.xls"), load("binary.xlsb"), ws.saveCsv("data.csv") also supported
```

Build: `msbuild XL++.sln /p:Configuration=Release /p:Platform=x64` (VS)  
or `cmake -B build -G Ninja && cmake --build build`

### Python

```python
import xlpp

wb = xlpp.Workbook()
ws = wb.add_worksheet("Sales")
ws["A1"].value = "Item"
ws["B1"].value = 100
ws.append(["Widget", 50, 19.99])
wb.save("sales.xlsx")
```

```bash
pip install xlpp
```

### C\#

```csharp
using XLPP;
using var wb = new Workbook();
var ws = wb.AddWorksheet("Report");
ws["A1"].Value = "Title";
ws["B1"].Value = 42.5;
wb.Save("report.xlsx");
```

---

## Architecture

```
Application (C++ / Python / C# / C)
    │
    ▼
┌─────────────────────────┐
│  Public API (Workbook,  │
│  Worksheet, Cell, Style)│
├─────────────────────────┤
│  Serializer (OOXML/     │  SIMD XML scanning
│  XLS/BIFF8, XLSB/BIFF12)│  mmap I/O + ThreadPool
│  ZIP Engine             │  uint64_t cell keys
│  Cell Storage           │  compact lazy payloads
├─────────────────────────┤
│  zlib (compression)     │
└─────────────────────────┘
```

## Project structure

```
XLPP/
├── include/XLPP/          # Public headers (~70 files, 247 types)
├── src/XLPP/              # Implementation (~100 files, 27K lines)
│   ├── Workbook/          # 22 focused modules (split from a 9.9K-line monolith)
│   ├── Legacy/            # XLS (BIFF8) + XLSB (BIFF12) readers/writers
│   ├── Packaging/         # ZIP, OPC relationship graph
│   └── ...                # XML, Streaming, Encryption, VBA, Pivot
├── bindings/
│   ├── python/            # pybind11 (167 public symbols)
│   ├── csharp/            # P/Invoke wrapper
│   └── c/                 # C ABI
├── tests/                 # 213+ unit suites, 3,900+ checks
├── docs/                  # Documentation
└── ROADMAP.md             # Development roadmap
```

## Code size

| Component | Lines |
|---|--:|
| Core (src + include) | ~34.5K |
| Tests | ~32.2K |
| Bindings (Python + C# + C) | ~8.3K |
| **Total** | **~75K** |

---

## License

MIT — see [LICENSE](LICENSE)

Detailed feature history, gap analysis and roadmap: [ROADMAP.md](ROADMAP.md), [docs/](docs/).
