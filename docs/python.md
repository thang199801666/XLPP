# XL++ Python Guide

The `xlpp` package is the Python binding for the XL++ C++20 workbook engine.
It reads, creates, edits, calculates, and saves Excel `.xlsx` and `.xlsm`
files without automating Microsoft Excel.

This guide documents the Python API shipped by this repository. Python names
use `snake_case`; the corresponding native C++ methods generally use
`camelCase`.

## Installation

Install the latest release from PyPI:

```bash
python -m pip install --upgrade xlpp
```

Check the installed version:

```python
import xlpp

print(xlpp.__version__)
```

### Wheel availability

The current release workflow builds these binary wheels:

| Platform | Architecture | CPython |
| --- | --- | --- |
| Windows | AMD64 | 3.10–3.13 |
| Linux manylinux | x86_64 | 3.10–3.13 |
| macOS 10.15+ | Apple Silicon | 3.10–3.13 |

`python_requires` is Python 3.8 or newer. Platforms or Python versions without
a published wheel require a source build with a C++20 compiler, Python
development headers, and enough memory to compile the pybind11 extension.

NumPy and openpyxl are not runtime dependencies. Install NumPy only when using
the array helpers:

```bash
python -m pip install numpy
```

## First workbook

```python
from datetime import date, datetime
import xlpp

book = xlpp.Workbook()
sheet = book.add_worksheet("Sales")

sheet.append(["Product", "Units", "Price", "Date"])
sheet.append(["Widget", 100, 9.99, date(2026, 8, 9)])
sheet.append(["Gadget", 50, 19.99, datetime(2026, 8, 10, 9, 30)])

sheet["A1"].font().bold = True
sheet["E2"].set_formula("=B2*C2")

book.save("sales.xlsx")
```

Load it again:

```python
book = xlpp.Workbook.open("sales.xlsx")
sheet = book["Sales"]

print(book.sheet_names)
print(sheet["A2"].value)
print(sheet["E2"].formula)
```

You can also use an existing `Workbook` object:

```python
book = xlpp.Workbook()
book.load("sales.xlsx")
```

## API conventions

### Workbook and worksheet access

Workbooks support sheet-name and zero-based integer indexing:

```python
first = book[0]
sales = book["Sales"]

print(len(book))
print(book.sheet_count)
print(book.active.name)

for sheet in book:
    print(sheet.name)
```

Missing worksheets and invalid indices raise exceptions. `get_worksheet()` and
`worksheet()` return `None` when a named worksheet does not exist.

### Native child objects

Methods such as `cell.font()`, `cell.fill()`, `sheet.auto_filter()`, and
`pivot.cache` expose objects owned by their parent. Keep the parent workbook
alive while using these objects.

Collection-backed handles are designed to remain stable while related native
collections grow. Do not use a child object after removing or clearing its
owner.

### Properties versus methods

Simple values normally use properties:

```python
sheet.name = "Revenue"
sheet["A1"].value = "Title"
chart.title = "Revenue by month"
```

Operations and native reference accessors use methods:

```python
sheet["A1"].font().bold = True
sheet.merge_cells("A1:C1")
book.calculate_formulas()
```

## Workbook I/O

### Files and memory

```python
book.save("output.xlsx")

payload = book.save_bytes()
assert isinstance(payload, bytes)

copy = xlpp.Workbook()
copy.load_bytes(payload)
```

### Save options

```python
options = xlpp.SaveOptions()
options.compression_level = xlpp.CompressionLevel.BEST
options.parallel_workers = 4
options.parallel_sheets = True
options.atomic_write = True
options.validate_before_save = True

book.save("output.xlsx", options)
```

Useful save controls include compression strategy, parallel row/sheet writing,
Strict OOXML namespaces, formula calculation before save, chart-cache
synchronization, durable writes, validation, and encryption.

### Load options and resource limits

```python
options = xlpp.LoadOptions()
options.max_file_bytes = 512 * 1024 * 1024
options.max_entry_bytes = 128 * 1024 * 1024
options.max_total_bytes = 1024 * 1024 * 1024
options.max_entries = 50_000

book = xlpp.Workbook.open("input.xlsx", options)
```

Set `options.lenient = True` only for deliberate recovery of malformed input.
Normal strict loading reports invalid package content earlier.

## Cells, ranges, and rows

### Values

```python
sheet["A1"].value = "Text"
sheet["B1"].value = 42
sheet["C1"].value = 3.14
sheet["D1"].value = True
sheet["E1"].value = date(2026, 8, 9)
sheet["F1"].value = datetime(2026, 8, 9, 14, 30)
sheet["G1"].value = None
```

| Python input | Excel cell |
| --- | --- |
| `str` | Text |
| `int`, `float` | Number |
| `bool` | Boolean |
| `datetime.date` | Date |
| `datetime.datetime` | Date and time |
| `None` | Empty |

Excel numeric values are returned as Python `float`.

### Cell inspection

```python
cell = sheet.cell("B2")

print(cell.address, cell.row, cell.column)
print(cell.value_type())
print(cell.is_numeric())
print(cell.has_formula())

cell.clear()
```

Use `try_cell()` when absent cells should not be created or treated as errors:

```python
cell = sheet.try_cell("Z99")
if cell is not None:
    print(cell.value)
```

### Ranges and rows

```python
block = sheet.range("A1:C10")
print(block.address())
print(block.values())

for row in sheet:
    print(row.number, row.values())
```

Bulk row insertion is available through `append()` and `from_records()`:

```python
sheet.from_records(
    [["Alice", 30], ["Bob", 25]],
    columns=["Name", "Age"],
)

records = sheet.to_records()
```

### Rich text, comments, and hyperlinks

```python
rich = xlpp.RichText()
run = xlpp.RichTextRun("Important")
run.bold = True
run.color = "FFFF0000"
rich.add_run(run)
sheet["A1"].set_rich_text(rich)

sheet["A2"].set_comment(xlpp.Comment("Review this row", "Finance"))
sheet["A2"].set_hyperlink(xlpp.Hyperlink("https://example.com"))
```

## Formulas

### Formula authoring

```python
sheet["C2"].set_formula("=A2+B2")
sheet["D2"].set_shared_formula("=A2*2", 1, "D2:D20")
sheet["E2"].set_array_formula("=SUM(A2:B20)", "E2:E2")
sheet["F2"].set_dynamic_array_formula("=_xlfn.SORT(A2:A20)", "F2:F20")
```

For newer Excel functions, `xlfn()` supplies the compatibility prefix:

```python
name = xlpp.xlfn("FILTER")
sheet["G2"].set_formula(f"={name}(A2:A20,B2:B20>0)")
```

### Calculation

```python
options = xlpp.CalculationOptions()
options.recursive_dependencies = True
options.update_cached_values = True
options.spill_dynamic_arrays = True

report = book.calculate_formulas(options)
print(report.formula_cells_evaluated)
print(report.unsupported_formulas)
print(report.warnings)
```

The native evaluator covers a broad function catalog and reference model, but
does not reproduce every Excel function or every edge-case behavior. Unsupported
expressions are reported rather than silently presented as fully calculated.

### Dependency and reference helpers

```python
graph = book.dependency_graph()
for edge in graph.edges:
    print(edge.dependent_sheet, edge.dependent_cell, edge.symbol)

renamed = xlpp.rename_worksheet_references(
    "=OldName!A1",
    "OldName",
    "NewName",
)
print(renamed.value)
```

The module also exposes `build_formula_dependency_graph()`,
`translate_formula_references()`, `translate_range_references()`, and
`invalidate_worksheet_references()`.

## Formatting and layout

Formatting accessors return native style objects:

```python
cell = sheet["A1"]
cell.font().bold = True
cell.font().size = 14
cell.font().color().set_argb("FFFFFFFF")

cell.fill().pattern_type = "solid"
cell.fill().foreground().set_argb("FF1F4E78")
cell.border().bottom().style = "thin"
cell.alignment().horizontal = "center"
cell.alignment().vertical = "center"
cell.alignment().wrap_text = True
cell.set_number_format("#,##0.00")
```

Worksheet layout:

```python
sheet.merge_cells("A1:C1")
sheet.freeze_panes("A2")
sheet.column_dimension("A").width = 24
sheet.row_dimension(1).height = 28
sheet.set_print_area("A1:F100")
sheet.print_titles_rows = "1:1"
sheet.hide_gridlines = True
```

Named styles, page setup, margins, worksheet views, workbook protection, and
worksheet protection are also exposed.

## Structural editing

Simple worksheet operations:

```python
sheet.insert_rows(2, 3)
sheet.delete_columns(5, 1)
```

Workbook-level structural editing can update dependent formulas, names,
tables, filters, validations, chart references, pivot references, drawings,
and hyperlinks transactionally:

```python
options = xlpp.StructuralEditOptions()
options.transactional = True
options.update_defined_names = True
options.fail_on_invalid_reference = True

report = book.insert_rows("Sales", 2, 3, options)
if not report.success:
    raise RuntimeError(report.warnings)
```

For dependency-aware renaming:

```python
report = book.rename_worksheet("Sales", "Revenue")
print(report.references_updated)
```

## Tables, filters, validation, and conditional formatting

```python
table = sheet.add_table("SalesTable", "A1:D100")
table.display_name = "SalesData"
table.style_info.name = "TableStyleMedium4"
table.style_info.show_row_stripes = True

filter_ = sheet.auto_filter()
filter_.reference = "A1:D100"

validation = sheet.data_validations.add(
    xlpp.DataValidationType.LIST,
    "E2:E100",
)
validation.formula1 = '"Open,Closed,Pending"'
validation.allow_blank = True

sheet.conditional_formatting.add_rule(
    "B2:B100",
    xlpp.ConditionalRuleType.FORMULA,
    "B2>100",
)
```

Imported advanced AutoFilter forms, including Top10, dynamic, color, icon,
and date-group filters, are represented by the public filter model.

## Charts

Create a chart from worksheet references:

```python
chart = xlpp.Chart(xlpp.ChartType.BAR)
chart.title = "Monthly sales"
chart.x_axis_title = "Month"
chart.y_axis_title = "Sales"
chart.grouping = xlpp.ChartGrouping.CLUSTERED
chart.width = 640
chart.height = 360

series = chart.add_series(xlpp.ChartSeries("Sales"))
series.categories_reference = "'Sales'!$A$2:$A$13"
series.values_reference = "'Sales'!$B$2:$B$13"

sheet.add_chart(chart)
```

The Python model exposes chart families, combined plots, native axes, series,
markers, data labels, trendlines, error bars, cached points, rich text, manual
layouts, 3D views, walls/floor, theme colors, and fill/line formatting.

For an imported chart, use its stable ID with worksheet chart-editing methods
to patch targeted ChartML fields without rebuilding unrelated drawing objects.
Synchronize formula-backed caches when needed:

```python
options = xlpp.ChartCacheSyncOptions()
options.changed_references_only = True
report = book.synchronize_chart_caches(options)
print(report.caches_updated)
```

## Pivot tables

The binding exposes the native pivot table, cache, field, reference, filter,
layout, and grouping models:

```python
pivot = xlpp.PivotTable("SalesPivot")
pivot.location = "G2"
pivot.layout = xlpp.PivotLayout.TABULAR
pivot.cache.source_data = "'Sales'!$A$1:$D$100"
pivot.cache.refresh_on_load = True
for field_name in ["Product", "Units", "Price", "Date"]:
    pivot.cache.add_field(field_name)

pivot.add_row_field("Product")
pivot.add_data_field("Units", "sum")
sheet.add_pivot_table(pivot)
```

PivotCharts, slicers, timelines, OLAP caches, and Data Model content have
preservation and inspection support. Their complete Excel UI-level authoring
semantics are not yet available.

## Images

```python
image = xlpp.Image.from_file("logo.png", "A1")
image.width_pixels = 160
image.height_pixels = 60
sheet.add_image(image)
```

For in-memory content:

```python
image = xlpp.Image("A1", png_bytes, "png")
assert isinstance(image.bytes, bytes)
```

PNG and JPEG are supported for authored images. Imported drawing-anchor and
package-origin metadata is available through `anchor_info`, `stable_id`, and
the source-part properties.

## Streaming large workbooks

The streaming writer avoids materializing the full worksheet model:

```python
writer = xlpp.StreamingWorkbookWriter(
    "large.xlsx",
    xlpp.SharedStringMode.BOUNDED_LRU,
    4096,
)
sheet = writer.add_worksheet("Data")

for row_number in range(1_000_000):
    sheet.append([row_number, f"row-{row_number}"])

writer.close()
```

Stream rows from an existing workbook:

```python
options = xlpp.StreamingReaderOptions()
options.max_file_bytes = 2 * 1024 * 1024 * 1024

reader = xlpp.StreamingWorkbookReader("large.xlsx", options)
for row_number, cells in reader.worksheet("Data"):
    print(row_number, [cell.value for cell in cells])
```

Iteration currently constructs a Python list for the selected worksheet. For
early termination or lower Python-side materialization, use `for_each_row()`
and return `False` from the callback.

## NumPy and pandas

```python
import numpy as np

values = np.arange(20, dtype=float).reshape(4, 5)
sheet.write_array(values, row=2, column=1)
result = sheet.to_array(2, 1, 5, 5)
```

For pandas, pass records explicitly:

```python
sheet.from_records(df.values.tolist(), columns=list(df.columns))
records = sheet.to_records()
df2 = pd.DataFrame(records[1:], columns=records[0])
```

`to_array()` is numeric and fills missing or non-numeric values with `0.0`.
Use records when mixed types must be preserved.

## VBA and macro-enabled workbooks

Load and save `.xlsm` files normally to preserve an existing VBA project:

```python
book = xlpp.Workbook.open("template.xlsm")
print(book.has_vba_project)
print([module.name for module in book.vba_modules])

book.set_vba_module_text("Module1", "Public Sub Run()\nEnd Sub\n")
book.save("result.xlsm")
```

Binary project APIs include `set_vba_project()`, `add_vba_project()`,
`vba_project_bytes`, `save_vba_project()`, and `remove_vba_project()`.

Signed projects and unsupported project forms require care: editing VBA can
invalidate signatures, and complete UserForms/ActiveX semantic authoring is
outside the current model.

## Encryption

Open an encrypted workbook:

```python
options = xlpp.LoadOptions()
options.password = "correct horse battery staple"
book = xlpp.Workbook.open("encrypted.xlsx", options)
```

Save with Agile AES-256/SHA-512 encryption:

```python
options = xlpp.SaveOptions()
options.encryption_password = "correct horse battery staple"
options.encryption_mode = xlpp.OfficeEncryptionMode.AGILE_AES256_SHA512
options.encryption_spin_count = 100_000
book.save("encrypted.xlsx", options)
```

Standard AES/SHA-1 profiles and encryption inspection fields are also exposed.
Do not hard-code passwords in production source code.

## External data and enterprise inspection

```python
external = book.inspect_external_data()
model = book.inspect_data_model()
enterprise = book.inspect_enterprise_features()

for feature in enterprise.features:
    print(feature.kind, feature.part_name, feature.semantic_editable)
```

The enterprise inventory covers PivotCharts, slicers, timelines, OLAP pivot
caches, Data Model, Power Query, SmartArt, ActiveX, and VBA UserForms. Selected
refresh/source metadata has targeted editing APIs. Most of these features are
preservation-first rather than fully authorable.

## Preservation and validation

Unknown or unsupported OPC package parts are retained when possible. You can
inspect preserved content and validate the workbook model:

```python
print([part.name for part in book.preserved_parts])
print(book.preserved_relationships)

report = book.validate()
for issue in report.issues:
    print(issue.severity, issue.message)
```

Always test representative production workbooks, especially those containing
Power Query, Data Model, signed VBA, embedded controls, or third-party OOXML
extensions. See [Preservation Core](PRESERVATION_CORE.md).

## Exceptions

Common mappings are:

| Condition | Python exception |
| --- | --- |
| Invalid argument, coordinate, or duplicate name | `ValueError` |
| Invalid index or range access | `IndexError` |
| Missing file, ZIP, encryption, or serialization failure | `RuntimeError` |
| XL++ domain exception | `xlpp.XLPPError` |

Reports such as `CalculationReport`, `StructuralEditReport`, and
`EnterpriseEditReport` provide warnings and counters for operations where a
partial or unsupported result should be inspected explicitly.

## Moving from openpyxl

The basic object model is intentionally familiar, but XL++ is not an API-drop-in
replacement:

| openpyxl | XL++ |
| --- | --- |
| `Workbook()` | `xlpp.Workbook()` |
| `load_workbook(path)` | `xlpp.Workbook.open(path)` |
| `wb.create_sheet("Data")` | `wb.add_worksheet("Data")` |
| `ws["A1"] = 42` | `ws["A1"].value = 42` |
| `ws["A1"].font.bold = True` | `ws["A1"].font().bold = True` |
| `ws.append(values)` | `ws.append(values)` |
| `wb.save(path)` | `wb.save(path)` |

XL++ uses native objects and explicit methods for style references. Validate
your formulas, date behavior, charts, pivots, and preservation requirements
before migrating a production workflow.

## Build and test from source

From the repository root:

```bash
python -m pip install --upgrade pip setuptools wheel pybind11
python -m pip install --editable .
python -m pip install pytest numpy openpyxl
python -m pytest bindings/python/tests -q
```

The source package compiles the XL++ sources and bundled zlib directly into the
extension. On Windows, use a Visual Studio developer environment. On Linux and
macOS, use a compiler with C++20 and `std::filesystem` support.

## Further reading

- [Python quick start](quickstart-python.md)
- [Binding parity](https://github.com/thang199801666/XLPP/blob/main/bindings/PARITY.md)
- [Current capabilities](CURRENT_CAPABILITIES.md)
- [Compatibility matrix](COMPATIBILITY_MATRIX.md)
- [Performance](performance.md)
- [Building XL++](https://github.com/thang199801666/XLPP/blob/main/BUILDING.md)
