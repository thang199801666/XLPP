# Cells, ranges, and rows

## Cell access

```python
cell = sheet["B2"]
same_cell = sheet.cell("B2")
same_again = sheet.cell(2, 2)

print(cell.address)   # B2
print(cell.row)       # 2
print(cell.column)    # 2
```

Cell coordinates are one-based. Invalid Excel coordinates raise `ValueError`.

`try_cell()` checks whether a cell already exists without forcing normal access:

```python
cell = sheet.try_cell("Z99")
if cell is None:
    print("No stored cell at Z99")
```

## Values and types

```python
from datetime import date, datetime

sheet["A1"].value = "text"
sheet["A2"].value = 42
sheet["A3"].value = 3.14
sheet["A4"].value = True
sheet["A5"].value = date(2026, 8, 9)
sheet["A6"].value = datetime(2026, 8, 9, 14, 30)
sheet["A7"].value = None
```

| Python input | Stored cell kind | Typical Python read value |
| --- | --- | --- |
| `str` | String | `str` |
| `int`, `float` | Numeric | `float` |
| `bool` | Boolean | `bool` |
| `date` | Excel serial + date format | `xlpp.DateTime`/converted date semantics |
| `datetime` | Excel serial + date-time format | `xlpp.DateTime`/converted date-time semantics |
| `None` | Empty | `None` |

Explicit setters are useful when type intent matters:

```python
cell.set_string_value("00123")
cell.set_numeric_value(123.0)
cell.set_bool_value(True)
cell.set_date(2026, 8, 9)
```

Inspect type and state:

```python
print(cell.value_type())
print(cell.has_value())
print(cell.is_string())
print(cell.is_numeric())
print(cell.is_bool())
print(cell.is_date())
```

!!! tip "Preserve leading zeros"

    Write identifiers such as postal codes and account numbers with
    `set_string_value()`, or apply an appropriate number format. Assigning an
    integer does not preserve textual leading zeros.

## Errors

```python
cell.set_error(xlpp.CellError.DIVISION_BY_ZERO)
if cell.is_error():
    print(cell.error())
```

Error codes include common Excel values such as `#DIV/0!`, `#N/A`, `#NAME?`,
`#NULL!`, `#NUM!`, `#REF!`, and `#VALUE!` through `CellError`.

## Formulas

```python
sheet["C2"].set_formula("=A2+B2")
print(sheet["C2"].has_formula())
print(sheet["C2"].formula)

sheet["C2"].clear_formula()
```

Formula strings should normally start with `=`. Formula assignment is explicit;
assigning a string beginning with `=` to `.value` stores text.

Advanced forms:

```python
sheet["D2"].set_shared_formula("=A2*2", 1, "D2:D20")
sheet["E2"].set_array_formula("=SUM(A2:B20)", "E2:E2")
sheet["F2"].set_dynamic_array_formula("=_xlfn.SORT(A2:A20)", "F2:F20")
```

See [Formulas and calculation](formulas.md) for evaluation, dependencies, and
reference translation.

## Ranges

```python
block = sheet.range("A1:C10")
same = sheet.range(1, 1, 10, 3)

print(block.address())
print(block.row_count)
print(block.column_count)
print(block.values())
print(block.formulas())
```

Apply an operation to a range:

```python
block.set_value(0)

def make_bold(cell):
    cell.font().bold = True

block.for_each(make_bold)
block.clear()
```

`cells()` returns a flat list; `rows()` returns nested row lists.

## Append and records

```python
sheet.append(["Name", "Age", "Score"])
sheet.append(["Alice", 30, 9.5])
sheet.append(["Bob", 25, 8.0])
```

Write records with a header:

```python
sheet.from_records(
    [["Alice", 30, 9.5], ["Bob", 25, 8.0]],
    columns=["Name", "Age", "Score"],
)

records = sheet.to_records()
```

!!! tip "Prefer bulk entry points"

    `append()`, `from_records()`, and `write_array()` reduce Python-to-C++ call
    overhead compared with deeply nested per-cell loops.

## Row iteration

```python
for row in sheet:
    print(row.number)
    print(row.values())
    for cell in row.cells:
        print(cell.address, cell.value)
```

For very large workbooks, use the [streaming reader](streaming-and-performance.md)
instead of loading the complete workbook model.

## Rich text

```python
rich = xlpp.RichText()

title = xlpp.RichTextRun("Quarterly ")
title.bold = True
rich.add_run(title)

accent = xlpp.RichTextRun("Report")
accent.italic = True
accent.color = "FF1F4E78"
rich.add_run(accent)

sheet["A1"].set_rich_text(rich)
print(sheet["A1"].rich_text_value.plain_text())
```

## Hyperlinks and comments

```python
link = xlpp.Hyperlink("https://example.com")
link.display = "Open report"
link.tooltip = "View the source"
sheet["A2"].set_hyperlink(link)

sheet["A2"].set_comment(xlpp.Comment("Verify this value", "Finance"))
```

Use `has_hyperlink()`, `hyperlink()`, `clear_hyperlink()`, `has_comment()`,
`comment()`, and `clear_comment()` for lifecycle operations.

## Cell helpers

```python
neighbor = sheet["B2"].offset(1, 0)
ref = xlpp.CellReference.parse("$C$7")

print(xlpp.is_valid_cell_coordinate(7, 3))
print(xlpp.CellReference.column_name(3))
```

Date helpers include `days_from_civil()`, `civil_from_days()`,
`to_iso8601_date()`, and `to_iso8601()`.
