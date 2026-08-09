# Styles and layout

## Why style accessors are methods

`font()`, `fill()`, `border()`, `alignment()`, and `style()` return native
objects owned by the cell. They are methods to make this reference behavior
explicit:

```python
cell = sheet["A1"]
cell.font().bold = True
cell.fill().pattern_type = "solid"
```

## Fonts

```python
font = sheet["A1"].font()
font.name = "Aptos"
font.size = 14
font.bold = True
font.italic = False
font.underline = True
font.color().set_argb("FFFFFFFF")
```

ARGB colors use eight hexadecimal digits: alpha, red, green, blue.
`FFFF0000` is opaque red and `00FF0000` is fully transparent red.

## Fills

```python
fill = sheet["A1"].fill()
fill.pattern_type = "solid"
fill.foreground().set_argb("FF1F4E78")
fill.background().set_argb("FFFFFFFF")
```

## Borders

```python
border = sheet["A1"].border()
border.top().style = "thin"
border.bottom().style = "double"
border.left().style = "thin"
border.right().style = "thin"
border.bottom().color().set_argb("FF000000")
```

## Alignment

```python
alignment = sheet["A1"].alignment()
alignment.horizontal = "center"
alignment.vertical = "center"
alignment.wrap_text = True
alignment.shrink_to_fit = False
alignment.text_rotation = 0
alignment.indent = 0
```

## Number formats

```python
sheet["B2"].set_number_format("#,##0.00")
sheet["C2"].set_number_format("$#,##0.00")
sheet["D2"].set_number_format("0.00%")
sheet["E2"].set_number_format("yyyy-mm-dd")
```

!!! tip "Dates are numbers plus formatting"

    Excel stores dates as serial numbers. Assign a Python `date`/`datetime` or
    use `set_date()`/`set_datetime()` so the value and date format agree.

## Reuse styles

Named styles reduce repeated style construction:

```python
header = xlpp.NamedStyle("Header")
header.style().font().bold = True
header.style().fill().pattern_type = "solid"
header.style().fill().foreground().set_argb("FF1F4E78")
book.add_named_style(header)

book.apply_named_style(sheet["A1"], "Header")
```

You can also copy a `Style` object and assign it through native style APIs.
Avoid creating thousands of nearly identical styles; Excel workbooks have
practical style-count limits and style catalogs affect file size.

## Merged cells

```python
sheet.merge_cells("A1:D1")
print(sheet.is_merged("B1"))
print(sheet.merged_ranges)
sheet.unmerge_cells("A1:D1")
```

Store content in the top-left cell of a merged range.

## Freeze panes and views

```python
sheet.freeze_panes("B2")
print(sheet.frozen_pane)
sheet.clear_freeze_panes()

sheet.hide_gridlines = True
view = sheet.sheet_view()
```

## Row and column dimensions

```python
sheet.column_dimension("A").width = 24
sheet.column_dimension(2).hidden = True
sheet.row_dimension(1).height = 30
sheet.row_dimension(10).hidden = True
```

Dimensions also expose outline, style, and visibility metadata represented by
the native model.

## Printing

```python
sheet.set_print_area("A1:F100")
sheet.print_titles_rows = "1:2"
sheet.print_titles_cols = "A:B"

setup = sheet.page_setup()
setup.orientation = xlpp.PageOrientation.LANDSCAPE
setup.fit_to_width = 1
setup.fit_to_height = 0

margins = sheet.page_margins()
margins.left = 0.25
margins.right = 0.25
```

Headers, footers, print options, paper size, scaling, and page breaks are
available through their corresponding worksheet models.

## Protection

```python
protection = sheet.protection()
protection.sheet = True
protection.set_password("edit-password")
```

Worksheet/workbook protection is an editing restriction, not file encryption.
Use [password-to-open encryption](vba-encryption-enterprise.md#encryption) when
the file content itself must be protected.

## Styling efficiently

- Build one style and reuse it instead of mutating every property repeatedly.
- Apply formatting only to the used range.
- Use named styles for repeated semantic roles.
- Use table styles for tabular stripes instead of formatting every row.
- Avoid styling huge empty regions; styled empty cells still occupy workbook
  model and package space.
