# Python documentation

The `xlpp` package gives Python direct access to the XL++ C++20 workbook
engine. This documentation is organized as a task-oriented tutorial and
cookbook, with API mapping tables for developers coming from openpyxl.

## Start here

1. [Install XL++](installation.md) and confirm wheel compatibility.
2. Follow [Workbook and worksheets](workbooks-and-worksheets.md).
3. Learn [Cells, ranges, and rows](cells-ranges-rows.md).
4. Add [styles and layout](styles-and-layout.md).
5. Open the topic guide for charts, pivots, formulas, or streaming.

```python
import xlpp

book = xlpp.Workbook()
sheet = book.add_worksheet("Data")
sheet["A1"].value = "Hello from XL++"
book.save("hello.xlsx")
```

## Guides by feature

<div class="grid cards" markdown>

-   **Workbook model**

    [Files, memory I/O, sheets, metadata, validation →](workbooks-and-worksheets.md)

-   **Cells and data**

    [Values, formulas, ranges, records, rich text →](cells-ranges-rows.md)

-   **Presentation**

    [Fonts, fills, borders, dimensions, printing →](styles-and-layout.md)

-   **Data features**

    [Tables, filters, validation, conditional formatting →](data-features.md)

-   **Charts and drawings**

    [Chart families, series, caches, images, imported edits →](charts-and-drawings.md)

-   **Pivot tables**

    [Caches, fields, grouping, layout, preservation →](pivot-tables.md)

-   **Calculation**

    [Formula authoring, evaluator, dependencies, translation →](formulas.md)

-   **Large files**

    [Streaming, NumPy, memory, performance techniques →](streaming-and-performance.md)

-   **Advanced workbooks**

    [VBA, encryption, external data, enterprise inspection →](vba-encryption-enterprise.md)

-   **Recipes**

    [Tips, patterns, safe round trips, debugging →](cookbook.md)

</div>

## API style at a glance

Simple values are properties. Operations and access to native child objects
are methods:

```python
sheet.name = "Revenue"                 # property
sheet["A1"].value = 42                 # property
sheet["A1"].font().bold = True         # native child accessor
sheet.merge_cells("A1:C1")             # operation
book.calculate_formulas()              # operation
```

The Python API uses `snake_case`, while C++ uses `camelCase`. See
[Feature/API coverage](api-coverage.md) for domain-by-domain mapping and
[Migrate from openpyxl](openpyxl-migration.md) for side-by-side examples.

!!! note "Parity and Excel compatibility are different"

    Binding parity means a public XL++ native model or operation is reachable
    from Python. It does not mean XL++ implements every feature in the Excel
    application. Preservation-only boundaries are documented explicitly.

