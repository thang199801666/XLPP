<div class="xlpp-hero" markdown>

# Build Excel workflows on a native engine

XL++ reads, creates, edits, calculates, and preserves `.xlsx` and `.xlsm`
workbooks from C++20, Python, C#, or C. It is designed for applications that
need native performance without automating Microsoft Excel.

[Install Python :material-language-python:](python/installation.md){ .md-button .md-button--primary }
[Python tutorial](python/index.md){ .md-button }
[GitHub :fontawesome-brands-github:](https://github.com/thang199801666/XLPP){ .md-button }

</div>

<div class="grid cards" markdown>

-   :material-speedometer:{ .lg .middle } **Native workbook engine**

    ---

    C++ cell storage, ZIP streaming, accelerated XML scanning, memory I/O, and
    parallel save paths behind a Python-friendly API.

-   :material-file-excel:{ .lg .middle } **Broad Excel model**

    ---

    Cells, styles, tables, filters, validation, charts, pivots, formulas, VBA,
    encryption, structural editing, and large-file streaming.

-   :material-shield-check:{ .lg .middle } **Preservation aware**

    ---

    Unsupported and enterprise package parts can survive unrelated edits, with
    inspection APIs and package validation for explicit safety boundaries.

-   :material-language-python:{ .lg .middle } **Native Python binding**

    ---

    The pybind11 module follows the public C++ model. Generated parity checks
    catch Workbook and Worksheet API drift.

</div>

## Python in 60 seconds

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

Then continue with the [Python documentation](python/index.md), browse the
[cookbook](python/cookbook.md), or compare equivalent
[openpyxl and XL++ code](python/openpyxl-migration.md).

## Choose a path

=== "Python"

    - [Installation and wheel support](python/installation.md)
    - [Workbook and worksheet tutorial](python/workbooks-and-worksheets.md)
    - [Feature/API coverage](python/api-coverage.md)
    - [Troubleshooting](python/troubleshooting.md)

=== "C++"

    - [C++ quick start](quickstart-cpp.md)
    - [Native API reference](https://github.com/thang199801666/XLPP/blob/main/API_REFERENCE.md)
    - [Building from source](https://github.com/thang199801666/XLPP/blob/main/BUILDING.md)

=== "C#"

    - [C# quick start](quickstart-csharp.md)
    - [NuGet package](https://www.nuget.org/packages/XLPP/)

## Know the compatibility boundary

XL++ has broad OOXML coverage, but it is not the Excel desktop application.
Some advanced Excel features are fully modeled, some support targeted edits,
and some are preservation/inspection-first. In particular, complete UI-level
authoring for Power Query, Data Model/OLAP, slicers, timelines, SmartArt,
ActiveX, UserForms, and every Excel formula behavior is not claimed.

Read these before deploying workbook round trips to production:

- [Current capabilities](CURRENT_CAPABILITIES.md)
- [Compatibility matrix](COMPATIBILITY_MATRIX.md)
- [Preservation core](PRESERVATION_CORE.md)
- [Known gaps and roadmap](MISSING_FEATURES_AND_DEVELOPMENT_ROADMAP.md)
