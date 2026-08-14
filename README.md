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

**High-performance C++20 Excel `.xlsx` / `.xlsm` / `.xltx` / `.xltm` read/write library.**  
Bundled-zlib support, Windows CNG or OpenSSL crypto, SIMD-accelerated parsing, multi-threaded save, and Python/C# bindings.

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



## P1X-A Chartsheet writer decomposition

P1X-A moves the complete Chartsheet package writer out of `Workbook.cpp` into `WorkbookChartsheetPackage.cpp` and hardens relationship-ID collision repair for preservation-backed `.xltx/.xltm` Chartsheets. A generated drawing relationship can now be reallocated without rewriting an unrelated preserved `legacyDrawingHF` owner that happens to use the same `rId`. The P1X gate includes a deliberate collision fixture, repeated-save package-graph validation, full regression and focused ASan/UBSan.

## P1W-A template / Chartsheet auxiliary ownership

P1W-A hardens `.xltx` / `.xltm` Chartsheet package ownership around features that live beside the chart itself. `Chartsheet` can now own an opaque `printerSettingsN.bin` payload through `pageSetup@r:id`; imported header/footer picture ownership through `legacyDrawingHF` / VML is preserved across metadata edits and chart regeneration; and the package validator now checks/inventories both auxiliary object families. Chart regeneration retires only the old drawing/chart closure instead of deleting unrelated sibling relationships.

The P1W release gate includes full regression, dedicated package-ownership tests, openpyxl 3.1.5 host validation, package-graph validation, C API/public-header/installed/source consumer checks and focused ASan+UBSan coverage. See [`docs/P1W_TEMPLATE_CHARTSHEET_PACKAGE_OWNERSHIP.md`](docs/P1W_TEMPLATE_CHARTSHEET_PACKAGE_OWNERSHIP.md).

P1V introduced advanced Chartsheet page setup/protection/custom views and moved Chartsheet XML I/O into `WorkbookChartsheetIO.cpp`; P1W builds on that model without decoding platform-specific printer structures.

## Feature Comparison

`Partial` means the feature is available with limitations or is not fully
round-trip compatible. `Write only` means the library cannot read an existing
workbook.

| Feature | XLPP | openpyxl (Python) | XlsxWriter (Python) | ClosedXML (C#) | libxlsxwriter (C++) |
|---------|:----:|:-----------------:|:-------------------:|:---------------:|:-------------------:|
| Read and write `.xlsx` | Yes | Yes | Write only | Yes | Write only |
| Excel templates `.xltx` / `.xltm` | Yes | Yes | Write only | Partial | Partial |
| `.xlsm` / VBA round-trip | Partial* | Partial | Partial | Partial | Partial |
| Cell values and formulas | Yes | Yes | Yes | Yes | Yes |
| Shared and array formulas | Yes | Partial | Partial | Partial | Partial |
| Fonts, fills, borders and number formats | Yes | Yes | Yes | Yes | Yes |
| Merged cells and freeze panes | Yes | Yes | Yes | Yes | Yes |
| Tables and auto-filters | Yes | Yes | Yes | Yes | Yes |
| Data validation | Yes | Yes | Yes | Yes | Yes |
| Conditional formatting | Yes | Yes | Yes | Yes | Yes |
| Charts | Yes | Yes | Yes | Yes | Yes |
| Pivot tables | Advanced* | Preserve/read | No | Partial | No |
| Images and hyperlinks | Yes | Yes | Yes | Yes | Yes |
| Comments and document properties | Yes | Yes | Yes | Yes | Yes |
| Worksheet and workbook protection | Yes | Yes | Yes | Yes | Yes |
| Password-to-open encryption | Yes† | Partial | No | Partial | No |
| Streaming large worksheets | Yes | Partial | Yes | Partial | Yes |
| Strict OOXML support | Yes | Partial | No | Partial | No |
| ZIP64 workbooks | Yes | Partial | Yes | Partial | Partial |
| Native C++ API | Yes | No | No | No | Yes |
| Python API | Yes | Yes | Yes | No | No |
| C# API | Yes | No | No | Yes | No |

`†` **P1I encryption scope:** XL++ writes/reads Agile AES-128/192/256-CBC with SHA-1/SHA-256/SHA-384/SHA-512 password profiles, and writes/reads Standard CryptoAPI AES-128/192/256 + SHA-1 password packages. P1I can inspect certificate key-encryptor metadata and decrypt by password when certificate entries coexist with the required password entry; certificate-only decryption/writing, RC4 and Extensible Encryption remain outside the supported writer/reader set.

`*` **Pivot/VBA scope through P1S:** XL++ can create/read common worksheet PivotTable/cache/cache-record structures, model shared caches, grouping, PivotChart linkage, filters/items, row/column/page/data fields, selectively edit imported cache/table metadata, and preserve untouched imported Pivot OOXML. P1S adds exact physical cache value kinds (missing/number/string/Boolean/error/date-time). OLAP, slicer/timeline and some calculated-member/advanced extension semantics remain partial. VBA/UserForm support is substantially deeper than simple `vbaProject.bin` preservation but still keeps unsupported designer/control payloads lossless rather than claiming complete semantic authoring.

P0M extends imported-chart editing with `dPt` data-point styling, rich chart-title/per-point label runs, gradient/pattern fills, color transforms, and advanced line/custom-dash metadata. As with P0L, these APIs patch only targeted ChartML subtrees and preserve sibling DrawingML/package relationships.

P0N extends preservation-aware chart editing with plot-area/legend manual layouts, axis rich titles, number formats, tick settings, major/minor units, crossing metadata, axis/gridline line formatting, and legend overlay/fill/line formatting. Imported charts retain native `axId` targeting so combined/secondary-axis charts are not flattened during selective edits.

P0O adds native axis scaling (`min`/`max`/`logBase`/orientation), numeric `crossesAt`, built-in/custom display units with optional rich labels, explicit major/minor-gridline lifecycle, and selective chart-area/plot-area fill and line formatting. These changes remain stable-ID/`axId` targeted; unrelated chart XML, sibling drawings and media are preserved.

P0P adds preservation-aware chart auxiliary objects: chart data tables, drop lines, high-low lines, up/down bars, and leader-line containers/formatting at plot or series label scope. Add/remove/edit operations patch only the owning ChartML subtree and preserve sibling DrawingML/media.

P0Q adds first-class StockChart inspection/generation (`Chart::Type::Stock`), generation-time auxiliary objects through `Chart::primaryPlot()`, and `ChartDataTable::textStyle` support for `dTable/txPr`. Both imported and newly generated stock charts support high-low and up/down bars without flattening preserved ChartML.
| SIMD-accelerated XML scanning | Yes | No | No | No | No |
| mmap zero-copy reading | Yes | No | No | No | No |
| Multi-threaded save | Yes | No | No | No | No |

## Three-pillar interoperability milestone (P1S-A)

P1S-A targets Basic XLSX read/write, Charts and PivotTables as explicit competitive gates. `Workbook` now round-trips XLTX/XLTM template identity; Scatter/Bubble generation uses the correct XY schema; generated charts can contain independent plots with a secondary Y axis; and Pivot caches can retain exact physical value kinds instead of inferring every cache value from text. A dedicated P1S regression covers all three areas, and the generated artifacts were also opened by an independent openpyxl 3.1.5 consumer. See [`docs/P1S_THREE_PILLAR_INTEROPERABILITY.md`](docs/P1S_THREE_PILLAR_INTEROPERABILITY.md).

P1T introduced first-class Chartsheets; P1U deepens them with mixed-tab active/visibility state, Chartsheet view/protection/page setup/header-footer semantics, and preservation-safe template round-trips.

## Core input-hardening milestone (P1Q-A)

P1Q-A hardens untrusted package handling without another `Cell`/`Style` ABI-layout change. Streaming reads now support explicit ZIP resource limits, normalize and validate workbook relationship targets, and bound the incremental XML buffer (64 MiB default, configurable/disableable). Truncated streamed XML and short CFB FAT/miniFAT chains are rejected instead of being treated as clean EOF/partial streams; CFB directory cycles/invalid references/duplicate paths are also detected. Main regression remains **198/198 suites / 3,632 checks PASS**, with a dedicated malformed-input suite plus focused ASan+UBSan passing. See [`docs/P1Q_CORE_HARDENING.md`](docs/P1Q_CORE_HARDENING.md).

## Core lazy-density milestone (P1P-A)

P1P-A moves the remaining inline `Style` and `FormulaMetadata` payloads behind deep-copy lazy value storage. On the local GCC/libstdc++ probe, `sizeof(Cell)` falls from 352 B to **152 B**; a same-profile one-million-cell `-O1` run lowers peak RSS by about **30%** and bulk-build time by about **38%** while preserving Store-save output size. Formula loading now keeps ordinary formulas metadata-allocation-free and explicitly clears stale shared/array metadata when a cell is converted to a normal formula. Main regression is **198/198 suites / 3,632 checks PASS**; dedicated formula-density regression, C API, header and installed-package checks pass. See [`docs/P1P_LAZY_MODEL_DENSITY.md`](docs/P1P_LAZY_MODEL_DENSITY.md). Native C++ consumers must rebuild because `Cell` layout changes again.

## Core scaling milestone (P1O-A)

P1O-A reduces the default local `Cell` footprint from 424 B to 352 B on the development GCC/libstdc++ probe by lazily materializing rare optional payloads, adds an insertion hint for monotonic bulk `Worksheet::append()`, initializes the canonical empty-sheet extents cache so bulk append never needs a first-query full rescan, and exposes peak RSS in the native core benchmark. It also moves VBA/UserForm, Pivot, chart-cache/reference and chart-style implementation out of `Workbook.cpp`, reducing the main translation unit from about 724 KB / 12.3K lines to about 525 KB / 8.8K lines. Public method signatures remain source compatible, but the `Cell` layout change requires native C++ consumers to rebuild. See [`docs/P1O_CORE_DECOMPOSITION_DENSITY.md`](docs/P1O_CORE_DECOMPOSITION_DENSITY.md).

## Core structural safety, transactions and strict validation (P1K)

P1K continues the correctness-first direction of P1J. Workbook-level row/column edits transform references across formulas, defined names, tables/filters/validation/conditional formatting, internal hyperlinks, chart references and caches, Pivot sources/cache schema/records, print ranges and drawing anchors, and are now transactional by default. Excel grid limits (`XFD1048576`) and worksheet/table/merge naming/geometry invariants are enforced before invalid XML can be produced.

```cpp
auto edit = wb.deleteColumns("Data", 2, 1);
wb.renameWorksheet("Data", "Input Data");

auto integrity = wb.validateModelIntegrity();
if (!integrity.ok()) {
    // inspect stable issue codes before save
}

xlpp::StructuralEditOptions editOptions;
editOptions.rollbackOnFailure = true;
editOptions.validateResult = true;

xlpp::SaveOptions save;
save.validateModelBeforeSave = true;
save.rejectModelWarningsBeforeSave = true;
save.validatePackageBeforeWrite = true;
wb.save("output.xlsx", save);
```

`Workbook::removeWorksheet()` now performs preservation-aware ownership cleanup for imported drawing/chart descendants and refuses to remove the final sheet. `Workbook::renameWorksheet()` is the reference-safe rename path; direct `Worksheet::rename()` validates the local name but cannot update sibling/workbook references because worksheets intentionally do not hold parent pointers.

P1K recognizes 3-D qualifiers such as `Start:End!A1`. Rename/remove operations can rewrite or invalidate 3-D endpoints safely; row/column structural edits preserve such references and report them as unsupported instead of guessing coordinate semantics. Transaction rollback restores worksheets in place, so caller-held `Worksheet&` identities remain stable after cancellation or failure.

## Password-to-open encryption (P1I)

XL++ supports real Office package encryption, distinct from worksheet/workbook protection. The default remains modern Agile AES-256-CBC/SHA-512 with HMAC data integrity, 4096-byte `EncryptedPackage` segmentation and a 100,000-spin password KDF. P1H made the Agile AES/hash profile configurable and added a Standard AES compatibility writer. P1I removes plaintext inner-ZIP temporary files from encrypted file load/save, adds certificate-key metadata inspection, and adds stricter encrypted-input policies.

```cpp
xlpp::SaveOptions save;
save.encryption.enabled = true;
save.encryption.password = "secret";
save.encryption.mode = xlpp::PackageEncryptionMode::Agile;
save.encryption.keyBits = 128;
save.encryption.hashAlgorithm = xlpp::PackageEncryptionHash::Sha256;
save.encryption.spinCount = 100000;
wb.save("protected.xlsx", save);

// Legacy/compatibility profile:
save.encryption.mode = xlpp::PackageEncryptionMode::Standard;
save.encryption.keyBits = 128; // 128/192/256
wb.save("protected-standard.xlsx", save);

const auto info = xlpp::Workbook::inspectPasswordEncryptionFile("protected.xlsx");
// info.keyEncryptorCount / passwordKeyEncryptorCount / certificateKeyEncryptors

xlpp::LoadOptions load;
load.passwordToOpen = "secret";
load.maxEncryptionInfoBytes = 1024u * 1024u;
load.allowStandardEncryption = false;       // optional enterprise policy
load.requireAgileDataIntegrity = true;      // optional enterprise policy
```

Agile supports AES-128/192/256 with SHA-1/SHA-256/SHA-384/SHA-512 in XL++. Standard uses AES-128/192/256 with SHA-1 and the format-defined 50,000-iteration derivation. `LoadOptions::maxEncryptionSpinCount`, `maxDecryptedPackageBytes`, and `maxEncryptionInfoBytes` provide resource guards for untrusted encrypted documents; `allowStandardEncryption` and `requireAgileDataIntegrity` provide format/integrity policy controls. Unicode passwords are accepted as UTF-8 by the API and encoded as UTF-16LE for Office key derivation.

On Windows, XL++ uses the system CNG/BCrypt provider; non-Windows CMake builds use OpenSSL `Crypto`. The default Agile AES-256/SHA-512 output and the Standard AES-128 compatibility writer have been opened independently with LibreOffice in the development sandbox. P1I understands mixed Agile key-encryptor lists well enough to locate the password descriptor and inspect certificate metadata; certificate-private-key decryption/writing, RC4 and Extensible Encryption are not yet implemented.

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

## Automatic Cache Dependency Tracking & Style Resolution (P0V)

XL++ now exposes the chart-cache dependency graph and can synchronize only caches whose A1 ranges intersect tracked cell changes. Tracking combines worksheet cell/range access with per-cell mutation revisions, so a retained `Cell&` changed after a save/reset is still detected. `Workbook::synchronizeChangedChartCaches()` avoids rebuilding unrelated title/category/value caches, while `chartCacheDependencies()` exposes owner/source sheet, series, cache kind and resolved range information for diagnostics or tooling. Formula cells without a current cached worksheet value can retain the matching existing chart-cache point instead of silently deleting it.

Worksheet change tracking is cleared after load and can be explicitly reset independently of the normal worksheet dirty flag. Chart style/color-style resources now expose resource IDs, color-style method and referenced colors; theme placeholders `+mj-lt` / `+mn-lt` resolve through the parsed major/minor Latin theme fonts, and color-style entries can be resolved to final RGB through the workbook theme.

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

### P0W chart-cache formula dependencies

Incremental chart-cache synchronization can now detect changes to simple A1 precedents of formula cells. XL++ deliberately does not evaluate formulas: when only a precedent changed, it preserves the chart cache and requests host recalculation. Imported chart color-style palettes can also be materialized selectively onto series through the workbook theme.


## P1N-A core memory-density milestone

The P1N-A development tree reduces the local GCC/libstdc++ `Cell` footprint from 1,072 to 424 bytes and `Style` from 616 to 184 bytes while retaining public source-level method behavior. It also removes temporary-tree allocation from `Worksheet::trackedCellChangeCount()`. Full regression is 196/196 suites and 3,600/3,600 checks passing. See `docs/P1N_CORE_MEMORY_DENSITY.md` for benchmarks, correctness coverage and the required C++ ABI rebuild note.


## P1R-A transactional load and package hardening

P1R-A gives materialized `Workbook::load()` a strong exception guarantee: failed loads no longer destroy the caller's existing workbook. Stream loading is bounded before temporary-file materialization, `LoadOptions` adds worksheet/cell/shared-string/defined-name object limits, critical OOXML numeric fields use exact parsing, and materialized worksheet cell enumeration is zero-copy. Owner-aware package relationships and duplicate content-type declarations are validated by a dedicated internal package-reader layer. See `docs/P1R_CORE_TRANSACTIONAL_PACKAGE_HARDENING.md`.

## P1U-A — template and Chartsheet production semantics

P1U-A concentrates on `.xltx` / `.xltm` and chart-only sheets. Mixed Worksheet/Chartsheet tabs now share active-tab and `visible`/`hidden`/`veryHidden` workbook state; Chartsheets model sheet properties, zoom/view, protection, margins/page setup and odd/even/first-page headers/footers. Imported Chartsheet metadata can be patched without regenerating the owned drawing/chart subtree, and repeated saves preserve both metadata-only patches and intentionally regenerated imported charts. Generated XLTX/XLTM artifacts and patched imported Chartsheets pass openpyxl 3.1.5 host validation and the XL++ package validator. See [`docs/P1U_TEMPLATE_CHARTSHEET_PRODUCTION.md`](docs/P1U_TEMPLATE_CHARTSHEET_PRODUCTION.md).

## P1T-A — first-class Chartsheets and semantic Pivot items

P1T-A strengthens the three openpyxl-comparison pillars again. Workbook tab order can now contain both Worksheets and first-class Chartsheets without changing legacy worksheet-only APIs. Generated or imported Chartsheets load/save, reorder, rename/remove and own chart drawings through a validated OPC graph; streaming readers report chart-only tabs separately. Pivot field items can bind to a typed logical cache value so hidden/item semantics survive shared-item reordering instead of depending on a stale physical `x` index. The generated Chartsheet + combo/secondary-axis package and repaired Pivot package are accepted by an independent openpyxl 3.1.5 host check. See `docs/P1T_THREE_PILLAR_SHEETMODEL_CHARTSHEET_PIVOT_IDENTITY.md`.
