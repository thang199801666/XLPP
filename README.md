# XL++

[![CI](https://github.com/thang199801666/XLPP/actions/workflows/ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/xlpp.svg)](https://pypi.org/project/xlpp/)
[![Documentation](https://img.shields.io/badge/docs-online-4051B5)](https://thang199801666.github.io/XLPP/)
[![C++20](https://img.shields.io/badge/C%2B%2B-20-00599C)](https://en.cppreference.com/w/cpp/20)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

XL++ is a native C++20 library for reading, creating, editing, and preserving
Excel `.xlsx` and `.xlsm` workbooks. It is available for C++, Python, C#, and C.

[Documentation](https://thang199801666.github.io/XLPP/) ·
[Python guide](https://thang199801666.github.io/XLPP/python/) ·
[API reference](API_REFERENCE.md) ·
[Building](BUILDING.md)

## Python

```bash
python -m pip install --upgrade xlpp
```

```python
from datetime import date
import xlpp

book = xlpp.Workbook()
sheet = book.add_worksheet("Sales")

sheet.append(["Product", "Units", "Price", "Date"])
sheet.append(["Widget", 100, 9.99, date(2026, 8, 9)])
sheet["A1"].font().bold = True
sheet["E2"].set_formula("=B2*C2")

book.save("sales.xlsx")
```

Published wheels currently cover CPython 3.10–3.13 on Windows x64, Linux
x86_64, and macOS Apple Silicon. See the
[installation guide](https://thang199801666.github.io/XLPP/python/installation/)
for source builds and platform details.

## C++

```cpp
#include <XLPP/XLPP.h>

int main() {
    xlpp::Workbook book;
    auto& sheet = book.addWorksheet("Report");

    sheet.cell("A1").setValue("Total");
    sheet.cell("B1").setValue(42.5);
    sheet.cell("C1").setFormula("=B1*2");

    book.save("report.xlsx");
}
```

```bash
cmake -S . -B build
cmake --build build --config Release
```

## Highlights

- Cell values, formulas, dates, errors, rich text, comments, and hyperlinks
- Styles, tables, filters, validation, conditional formatting, and protection
- Charts, chart caches, pivot tables, images, and drawing metadata
- Formula calculation, dependency graphs, and reference-aware structural edits
- VBA projects, password-to-open encryption, Strict OOXML, and ZIP64
- Streaming reader/writer, memory I/O, NumPy helpers, and parallel save paths
- Preservation and inspection of advanced workbook package content

See [current capabilities](docs/CURRENT_CAPABILITIES.md) and the
[compatibility matrix](docs/COMPATIBILITY_MATRIX.md) for exact support levels.

## Bindings

| Language | Package or guide |
| --- | --- |
| Python | [`pip install xlpp`](https://pypi.org/project/xlpp/) · [guide](https://thang199801666.github.io/XLPP/python/) |
| C# | [NuGet](https://www.nuget.org/packages/XLPP/) · [quick start](docs/quickstart-csharp.md) |
| C | [C ABI source](bindings/c/) |

The Python API follows the public native model with Python-style `snake_case`
names. Generated parity checks detect drift in public Workbook and Worksheet
methods.

## Compatibility boundary

XL++ does not automate Microsoft Excel and does not claim every Excel feature
or formula behavior. Some enterprise features are fully modeled, while others
are available through targeted editing, inspection, or preservation-only paths.

Complete semantic authoring is not currently claimed for Power Query, Data
Model/OLAP, slicers, timelines, SmartArt, ActiveX, or UserForms. Test complex
production workbooks with the target Excel version before deployment.

See [preservation guarantees](docs/PRESERVATION_CORE.md) and
[known gaps](docs/MISSING_FEATURES_AND_DEVELOPMENT_ROADMAP.md).

## Development

```bash
cmake -S . -B build -DXLPP_BUILD_TESTS=ON
cmake --build build --config Release
ctest --test-dir build -C Release --output-on-failure
```

Python source checkout:

```bash
python -m pip install -e .
python -m pytest bindings/python/tests -q
```

## License

[MIT](LICENSE)
