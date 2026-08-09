# Python Quick Start

This page is the shortest path from installation to a working workbook. See
the [complete Python guide](python.md) for charts, pivots, VBA, encryption,
streaming, structural editing, and preservation.

## Install

```bash
python -m pip install --upgrade xlpp
```

Published wheels currently cover CPython 3.10–3.13 on Windows x64, Linux
x86_64, and macOS Apple Silicon.

## Create a workbook

```python
from datetime import date
import xlpp

book = xlpp.Workbook()
sheet = book.add_worksheet("Sales")

sheet.append(["Product", "Units", "Price", "Date"])
sheet.append(["Widget", 100, 9.99, date(2026, 8, 9)])
sheet.append(["Gadget", 50, 19.99, date(2026, 8, 10)])

sheet["A1"].font().bold = True
sheet["A1"].fill().pattern_type = "solid"
sheet["A1"].fill().foreground().set_argb("FFD9EAF7")
sheet["C2"].set_number_format("$#,##0.00")
sheet["E2"].set_formula("=B2*C2")

sheet.freeze_panes("A2")
sheet.column_dimension("A").width = 24

book.save("sales.xlsx")
```

## Read a workbook

```python
book = xlpp.Workbook.open("sales.xlsx")
sheet = book["Sales"]

print(book.sheet_names)
print(sheet["A2"].value)
print(sheet["E2"].formula)

for row in sheet:
    print(row.number, row.values())
```

## Work with memory

```python
payload = book.save_bytes()

copy = xlpp.Workbook()
copy.load_bytes(payload)
```

## openpyxl mapping

| openpyxl | XL++ |
| --- | --- |
| `Workbook()` | `xlpp.Workbook()` |
| `load_workbook(path)` | `xlpp.Workbook.open(path)` |
| `wb.create_sheet("Data")` | `wb.add_worksheet("Data")` |
| `ws["A1"] = 42` | `ws["A1"].value = 42` |
| `ws["A1"].font.bold = True` | `ws["A1"].font().bold = True` |
| `ws.append(values)` | `ws.append(values)` |
| `wb.save(path)` | `wb.save(path)` |

XL++ is not API-compatible with openpyxl. Style objects are accessed through
methods because they reference native objects owned by the workbook.

## Next steps

- [Complete Python guide](python.md)
- [Current capabilities](CURRENT_CAPABILITIES.md)
- [Compatibility matrix](COMPATIBILITY_MATRIX.md)
- [Binding parity](https://github.com/thang199801666/XLPP/blob/main/bindings/PARITY.md)
- [Source build instructions](https://github.com/thang199801666/XLPP/blob/main/BUILDING.md)
