# Quick Start — Python

## Install

```bash
pip install pybind11 setuptools
cd bindings/python
pip install .
```

## Usage

```python
import xlpp
from datetime import date, datetime

# Create workbook
wb = xlpp.Workbook()
ws = wb.add_worksheet("Sales")

# Write cells (auto type conversion)
ws['A1'].value = "Product"          # str -> string
ws['B1'].value = 100                # int -> number
ws['C1'].value = 9.99               # float -> number
ws['D1'].value = True               # bool -> boolean
ws['E1'].value = date(2024, 1, 15)  # date -> DateTime
ws['F1'].value = datetime.now()     # datetime -> DateTime
ws['G1'].value = None               # None -> empty
ws['H1'].value = "=A2+B2"           # set_formula() for formulas

# Style (fluent chaining, returns self)
ws['A1'].font().bold = True
ws['A1'].font().size = 14
ws['A1'].font().italic = True
ws['A1'].font().color().set_argb("FFFF0000")   # red text
ws['A1'].fill().pattern_type = "solid"
ws['A1'].fill().foreground().set_argb("FFCCFFCC")  # light green bg
ws['A1'].border().bottom().style = "thin"
ws['A1'].alignment().horizontal = "center"
ws['B1'].number_format = "#,##0.00"

# Bulk append (like openpyxl)
ws.append(["Widget", 100, 9.99, date(2024, 1, 15)])
ws.append(["Gadget",   50, 19.99, date(2024, 1, 16)])
ws.append(["Sprocket", 200, 4.99, date(2024, 1, 17)])

# Layout
ws.merge_cells("A5:C5")
ws['A5'].value = "Summary"
ws.freeze_panes("A2")

# Metadata
wb.properties.title = "Sales Report"
wb.properties.creator = "XL++"

# Save
wb.save("sales.xlsx")

# Save with options
opts = xlpp.SaveOptions()
opts.parallel_workers = 4      # 4 threads
opts.parallel_sheets = True     # per-sheet parallelism
opts.parallel_rows = True       # per-row parallelism (single sheet)
opts.compression_level = xlpp.CompressionLevel.BEST
wb.save("sales.xlsx", opts)

# Load
wb2 = xlpp.Workbook()
wb2.load("sales.xlsx")
ws2 = wb2['Sales']
print(ws2['A1'].value)          # "Product"
print(ws2['B2'].value)          # 100.0

# Iteration
for row in ws2:
    print(row.number, row.values())

# Cell access
cell = ws2['A1']
print(cell.address)              # "A1"
print(cell.row, cell.column)     # (1, 1)
print(cell.value)                # current value
print(cell.value_type())         # "string", "numeric", "bool", ...
print(cell.is_numeric())         # True/False
print(cell.has_formula())        # True/False

# Safe access
c = ws2.try_cell("Z99")
if c is None:
    print("Cell does not exist")

# Multi-sheet
wb2.add_worksheet("Inventory")
wb2.add_worksheet("HR")
print(wb2.sheet_names)           # ['Sales', 'Inventory', 'HR']
print(len(wb2))                  # 3
print(wb2.active.name)           # 'Sales'

# Copy worksheet
wb2.copy_worksheet(ws2, "Sales Copy")
wb2.remove_worksheet("HR")

# Hyperlinks + comments
cell = ws2['A1']
cell.set_hyperlink(xlpp.Hyperlink("https://example.com"))
cell.set_comment(xlpp.Comment("This is important", "Alice"))

# AutoFilter
af = ws2.auto_filter()
af.reference = "A1:C100"
col = af.column(0)
col.add_value("Widget")
```

## API mapping (openpyxl → XL++)

| openpyxl | XL++ |
|----------|------|
| `wb = Workbook()` | `wb = xlpp.Workbook()` |
| `ws = wb.active` | `ws = wb.active` |
| `ws['A1'] = 42` | `ws['A1'].value = 42` |
| `ws['A1'].font.bold = True` | `ws['A1'].font().bold = True` |
| `ws.append([1, 2, 3])` | `ws.append([1, 2, 3])` |
| `ws.merge_cells('A1:B2')` | `ws.merge_cells('A1:B2')` |
| `wb.save('f.xlsx')` | `wb.save('f.xlsx')` |
| `wb = load_workbook('f.xlsx')` | `wb.load('f.xlsx')` |
| `ws.max_row` | `ws.max_row` |
| `ws.dimensions` | `ws.dimensions` |
| `cell.value` | `cell.value` |
| `cell.column_letter` | (use `cell.column`) |

## Type conversion

| Python | XL++ |
|--------|------|
| `str` | `CellValue{string}` |
| `int`, `float` | `CellValue{double}` |
| `bool` | `CellValue{bool}` |
| `datetime.date` | `DateTime{date only}` |
| `datetime.datetime` | `DateTime{date + time}` |
| `None` | `std::monostate` (empty) |

## NumPy & Pandas integration

Bulk operations bypass the per-cell Python overhead for maximum throughput.

### NumPy array write/read

```python
import numpy as np

# Write a 2D array to a worksheet (row1, col1 by default)
arr = np.random.rand(1000, 10)
ws.write_array(arr)                    # writes to A1:J1000
ws.write_array(arr, 2, 1)              # writes to A2:J1001
ws.write_array(arr, 1, 1, transpose=True)  # transposed

# Read a region back into a numpy array
out = ws.to_array(1, 1, 1000, 10)      # shape (1000, 10)
out = ws.to_array()                    # full used range
```

**Performance**: 1000×10 array = **3 ms** (vs ~200 ms per-cell).

### Records (list-of-lists)

```python
# Write with header row
ws.from_records(
    [["Alice", 30, 9.5], ["Bob", 25, 8.0]],
    columns=["Name", "Age", "Score"]
)

# Read back (list of rows, header first)
records = ws.to_records()
# [['Name', 'Age', 'Score'], ['Alice', 30.0, 9.5], ...]
```

### Pandas DataFrame

```python
import pandas as pd

df = pd.DataFrame({"Name": ["A", "B"], "Age": [30, 25], "Score": [9.5, 8.0]})

# DataFrame -> worksheet (via from_records)
ws.from_records(df.values.tolist(), columns=list(df.columns))

# Worksheet -> DataFrame (via to_records)
records = ws.to_records()
df2 = pd.DataFrame(records[1:], columns=records[0])
```

> Note: Excel stores all numbers as double, so `int` columns become `float` on round-trip.
