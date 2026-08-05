# XLPP Python Guide

XLPP is a native C++ Excel library with a Python extension module. It reads and
writes `.xlsx` files and keeps workbook parts that are not exposed by the
Python API when they are loaded and saved again.

## Installation

Install the released wheel from PyPI:

```bash
python -m pip install --upgrade xlpp
```

XLPP supports CPython 3.9 through 3.14 on the platforms covered by the
published wheels. A compiler is not required when a compatible wheel exists.

To install from a source checkout:

```bash
python -m pip install --upgrade pip setuptools wheel pybind11
python -m pip install .
```

The source build compiles XLPP, the bundled zlib sources, and the pybind11
extension in one step. NumPy is optional and is only needed for array helpers.

Check the installed version:

```python
import xlpp
print(xlpp.__version__)
```

## Basic Workbook

```python
from datetime import date, datetime
import xlpp

book = xlpp.Workbook()
sheet = book.add_worksheet("Sales")

sheet["A1"].value = "Product"
sheet["B1"].value = "Units"
sheet["C1"].value = "Date"
sheet.append(["Widget", 100, date(2024, 1, 15)])
sheet.append(["Gadget", 50, datetime(2024, 1, 16, 9, 30)])

book.save("sales.xlsx")
```

Supported cell assignments are:

| Python value | Excel value |
| --- | --- |
| `str` | Text |
| `int`, `float` | Number |
| `bool` | Boolean |
| `datetime.date` | Date |
| `datetime.datetime` | Date and time |
| `None` | Empty cell |

Excel numbers are represented as `float` when read back.

Load an existing workbook with:

```python
book = xlpp.Workbook()
book.load("sales.xlsx")
sheet = book["Sales"]
print(sheet["A2"].value)
```

`Workbook` can be indexed by sheet name or zero-based index. It is also
iterable:

```python
print(book.sheet_names)
print(book.active.name)
for sheet in book:
    print(sheet.name, sheet.dimensions)
```

## Cells and Formulas

```python
cell = sheet.cell("B2")
cell.value = 100
cell.set_formula("=B2*1.2")

print(cell.address)
print(cell.row, cell.column)
print(cell.value_type())
print(cell.has_formula())
print(cell.formula)

cell.clear()
```

Use `try_cell()` when a missing cell should not raise:

```python
cell = sheet.try_cell("Z99")
if cell is not None:
    print(cell.value)
```

Dynamic-array formulas can be written with:

```python
sheet["E2"].set_dynamic_array_formula("_xlfn.SORT(A2:A20)", "E2:E20")
```

For newer Excel functions, `xlpp.xlfn()` adds the required prefix:

```python
sheet["F2"].set_formula(f"={xlpp.xlfn('FILTER')}(A2:A20,B2:B20>0)")
```

## Formatting

Formatting objects are accessed as methods because they return references to
native style objects:

```python
title = sheet["A1"]
title.font().bold = True
title.font().size = 14
title.font().color().set_argb("FFFFFFFF")
title.fill().pattern_type = "solid"
title.fill().foreground().set_argb("FF1F4E78")
title.alignment().horizontal = "center"
title.alignment().vertical = "center"
title.alignment().wrap_text = True
title.border().bottom().style = "thin"
sheet["B2"].number_format = "#,##0.00"
```

ARGB colors use eight hexadecimal digits: alpha, red, green, blue. For
example, `FFFF0000` is opaque red.

## Layout and Worksheet Operations

```python
sheet.merge_cells("A1:C1")
sheet.freeze_panes("A2")

sheet.column_dimension("A").width = 24
sheet.row_dimension(1).height = 24

sheet.auto_filter().reference = "A1:C100"
sheet.hide_gridlines = True
```

Worksheet management:

```python
copy = book.copy_worksheet(sheet, "Sales Copy")
book.remove_worksheet("Sales Copy")
```

Worksheet names must be unique. Missing worksheets and invalid cell
references raise Python exceptions instead of returning invalid native
objects.

## Tables

```python
table = sheet.add_table("SalesTable", "A1:C3")
table.display_name = "SalesDisplay"
table.show_header_row = True
table.show_totals_row = False
table.style_info.name = "TableStyleMedium4"
table.style_info.show_row_stripes = True
table.add_column("Product")
```

The table reference is an Excel range such as `A1:C100`. Add the table after
the header and data cells have been written.

## Data Validation and Conditional Formatting

```python
validation = sheet.data_validations.add(xlpp.DataValidationType.LIST, "D2:D100")
validation.formula1 = '"Open,Closed,Pending"'
validation.allow_blank = True

sheet.conditional_formatting.add_rule(
    "B2:B100",
    xlpp.ConditionalRuleType.FORMULA,
    "B2>0",
)
```

Use `DataValidationType.LIST`, `WHOLE`, `DECIMAL`, `DATE`, `TIME`, `TEXT_LENGTH`,
or `CUSTOM` according to the required Excel validation rule.

## Charts

```python
chart = xlpp.Chart(xlpp.ChartType.BAR)
chart.title = "Sales by Product"
chart.x_axis_title = "Product"
chart.y_axis_title = "Units"
chart.grouping = xlpp.ChartGrouping.CLUSTERED
chart.width = 640
chart.height = 360

series = chart.add_series(xlpp.ChartSeries("Units"))
series.categories_reference = "'Sales'!$A$2:$A$3"
series.values_reference = "'Sales'!$B$2:$B$3"
sheet.add_chart(chart)
```

Chart references use normal Excel formulas and should include the worksheet
name when the source is on another sheet. Available chart types include bar,
line, pie, doughnut, radar, area, scatter, and bubble variants.

## Hyperlinks, Comments, and Properties

```python
cell = sheet["A2"]
cell.set_hyperlink(xlpp.Hyperlink("https://example.com"))
cell.hyperlink.display = "Open product page"
cell.hyperlink.tooltip = "Product details"
cell.set_comment(xlpp.Comment("Review this row", "Finance"))

book.properties.title = "Sales Report"
book.properties.creator = "XLPP"
book.properties.subject = "Monthly sales"
```

## NumPy and Records

NumPy helpers are optional. Install NumPy separately:

```bash
python -m pip install numpy
```

Write and read numeric arrays:

```python
import numpy as np

values = np.arange(20, dtype=float).reshape(4, 5)
sheet.write_array(values, row=2, column=1)
result = sheet.to_array(2, 1, 5, 5)
```

For mixed values, use records:

```python
sheet.from_records(
    [["Alice", 30, 9.5], ["Bob", 25, 8.0]],
    columns=["Name", "Age", "Score"],
)
records = sheet.to_records()
```

`to_array()` returns numeric values and fills missing/non-numeric cells with
`0.0`. `to_records()` preserves mixed cell values.

## Save and Load Options

```python
save_options = xlpp.SaveOptions()
save_options.compression_level = xlpp.CompressionLevel.BEST
save_options.parallel_workers = 4
book.save("sales.xlsx", save_options)

load_options = xlpp.LoadOptions()
load_options.lenient = True
book.load("possibly_damaged.xlsx", load_options)
```

Use lenient loading only when recovering data from a malformed workbook. A
normal load is preferable because it reports invalid package content early.

## Preservation and Compatibility

XLPP preserves many package parts that are not represented by the Python API,
including unknown XML parts, media, charts, pivot-related parts, and VBA
binary parts when they are present in the input package. Always test a
load/save round-trip with a representative workbook before modifying files in
production.

The library does not calculate formulas. It writes formula expressions and
preserves cached values where available; Excel or another calculation engine
must recalculate formulas after changes.

## Exceptions

Common Python exception mappings are:

| Native error | Python exception |
| --- | --- |
| Invalid argument or duplicate name | `ValueError` |
| Invalid index or range access | `IndexError` |
| File, ZIP, or serialization failure | `RuntimeError` |

## Testing a Source Checkout

```bash
python -m pip install --upgrade pytest pybind11 numpy openpyxl
python setup.py build_ext --inplace
python -m pytest bindings/python/tests -q
```

The published wheels are built by GitHub Actions with cibuildwheel. Local
native libraries are not required for the release workflow.
