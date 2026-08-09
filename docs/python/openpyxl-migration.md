# Migrate from openpyxl

XL++ deliberately feels familiar to openpyxl users, but it is not a drop-in
replacement. openpyxl is a pure-Python OOXML library; XL++ exposes a native C++
workbook model and adds formula calculation, encryption, package preservation,
and native streaming services with different object-lifetime rules.

The examples below target the openpyxl 3.1 API documented in its
[official tutorial](https://openpyxl.readthedocs.io/en/3.1/tutorial.html).

## Create and load

=== "openpyxl"

    ```python
    from openpyxl import Workbook, load_workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    wb.save("book.xlsx")

    wb = load_workbook("book.xlsx")
    ```

=== "XL++"

    ```python
    import xlpp

    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Data")
    wb.save("book.xlsx")

    wb = xlpp.Workbook.open("book.xlsx")
    ```

XL++ starts with no worksheet. Add one before writing.

## Cells and rows

=== "openpyxl"

    ```python
    ws["A1"] = "Name"
    ws.cell(row=2, column=1, value="Alice")
    ws.append(["Bob", 25])
    value = ws["A1"].value
    ```

=== "XL++"

    ```python
    ws["A1"].value = "Name"
    ws.cell(2, 1).value = "Alice"
    ws.append(["Bob", 25])
    value = ws["A1"].value
    ```

## Formulas

=== "openpyxl"

    ```python
    ws["C2"] = "=A2+B2"
    ```

=== "XL++"

    ```python
    ws["C2"].set_formula("=A2+B2")
    report = wb.calculate_formulas()
    ```

openpyxl writes formulas but does not calculate them. XL++ can calculate a
broad supported subset and update cached values, with diagnostics for unsupported
or failed formulas.

## Styles

=== "openpyxl"

    ```python
    from openpyxl.styles import Font, PatternFill, Alignment

    ws["A1"].font = Font(bold=True, color="FFFFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="FF1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["B2"].number_format = "#,##0.00"
    ```

=== "XL++"

    ```python
    ws["A1"].font().bold = True
    ws["A1"].font().color().set_argb("FFFFFFFF")
    ws["A1"].fill().pattern_type = "solid"
    ws["A1"].fill().foreground().set_argb("FF1F4E78")
    ws["A1"].alignment().horizontal = "center"
    ws["B2"].set_number_format("#,##0.00")
    ```

XL++ style methods return references to native child objects; keep the workbook
alive while using them.

## Dimensions, merge, and panes

| Task | openpyxl | XL++ |
| --- | --- | --- |
| Merge | `ws.merge_cells("A1:C1")` | `ws.merge_cells("A1:C1")` |
| Unmerge | `ws.unmerge_cells("A1:C1")` | `ws.unmerge_cells("A1:C1")` |
| Freeze | `ws.freeze_panes = "A2"` | `ws.freeze_panes("A2")` |
| Column width | `ws.column_dimensions["A"].width = 24` | `ws.column_dimension("A").width = 24` |
| Row height | `ws.row_dimensions[1].height = 30` | `ws.row_dimension(1).height = 30` |
| Print area | `ws.print_area = "A1:F100"` | `ws.set_print_area("A1:F100")` |

## Tables

=== "openpyxl"

    ```python
    from openpyxl.worksheet.table import Table, TableStyleInfo

    table = Table(displayName="SalesData", ref="A1:D100")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4")
    ws.add_table(table)
    ```

=== "XL++"

    ```python
    table = ws.add_table("SalesTable", "A1:D100")
    table.display_name = "SalesData"
    table.style_info.name = "TableStyleMedium4"
    ```

## AutoFilter

=== "openpyxl"

    ```python
    ws.auto_filter.ref = "A1:D100"
    ws.auto_filter.add_filter_column(3, ["Open", "Pending"])
    ```

=== "XL++"

    ```python
    filter_ = ws.auto_filter()
    filter_.reference = "A1:D100"
    column = filter_.column(3)
    column.add_value("Open")
    column.add_value("Pending")
    ```

Both libraries configure filter instructions; Excel applies the actual row
visibility/filter action when the workbook is opened. See openpyxl's official
[filters and sorts guide](https://openpyxl.readthedocs.io/en/stable/filters.html).

## Data validation

=== "openpyxl"

    ```python
    from openpyxl.worksheet.datavalidation import DataValidation

    validation = DataValidation(
        type="list",
        formula1='"Open,Closed,Pending"',
        allow_blank=True,
    )
    ws.add_data_validation(validation)
    validation.add("D2:D100")
    ```

=== "XL++"

    ```python
    validation = ws.data_validations.add(
        xlpp.DataValidationType.LIST,
        "D2:D100",
    )
    validation.formula1 = '"Open,Closed,Pending"'
    validation.allow_blank = True
    ```

## Conditional formatting

=== "openpyxl"

    ```python
    from openpyxl.formatting.rule import FormulaRule

    ws.conditional_formatting.add(
        "C2:C100",
        FormulaRule(formula=["C2>100"]),
    )
    ```

=== "XL++"

    ```python
    ws.conditional_formatting.add_rule(
        "C2:C100",
        xlpp.ConditionalRuleType.FORMULA,
        "C2>100",
    )
    ```

## Charts

=== "openpyxl"

    ```python
    from openpyxl.chart import BarChart, Reference

    chart = BarChart()
    chart.title = "Sales"
    data = Reference(ws, min_col=2, min_row=1, max_row=13)
    categories = Reference(ws, min_col=1, min_row=2, max_row=13)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    ws.add_chart(chart, "F2")
    ```

=== "XL++"

    ```python
    chart = xlpp.Chart(xlpp.ChartType.BAR)
    chart.title = "Sales"
    series = chart.add_series(xlpp.ChartSeries("Sales"))
    series.categories_reference = "'Data'!$A$2:$A$13"
    series.values_reference = "'Data'!$B$2:$B$13"
    ws.add_chart(chart)
    ```

XL++ additionally exposes imported stable IDs and targeted ChartML editing,
chart cache synchronization, combined plots, native axes, rich text, 3D, and
advanced formatting models.

## Images, hyperlinks, and comments

| Task | openpyxl | XL++ |
| --- | --- | --- |
| Image | `ws.add_image(Image(path), "A1")` | `ws.add_image(xlpp.Image.from_file(path, "A1"))` |
| Hyperlink | `cell.hyperlink = url` | `cell.set_hyperlink(xlpp.Hyperlink(url))` |
| Comment | `cell.comment = Comment(text, author)` | `cell.set_comment(xlpp.Comment(text, author))` |

## Large-file modes

openpyxl documents these as
[optimized read-only and write-only modes](https://openpyxl.readthedocs.io/en/stable/optimized.html).

=== "openpyxl"

    ```python
    wb = load_workbook("large.xlsx", read_only=True)
    for row in wb["Data"].rows:
        consume([cell.value for cell in row])
    wb.close()

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    ws.append([1, 2, 3])
    wb.save("large.xlsx")
    ```

=== "XL++"

    ```python
    reader = xlpp.StreamingWorkbookReader("large.xlsx")
    reader.for_each_row("Data", consume)

    writer = xlpp.StreamingWorkbookWriter("large.xlsx")
    ws = writer.add_worksheet("Data")
    ws.append([1, 2, 3])
    writer.close()
    ```

## VBA

=== "openpyxl"

    ```python
    wb = load_workbook("template.xlsm", keep_vba=True)
    wb.save("result.xlsm")
    ```

=== "XL++"

    ```python
    wb = xlpp.Workbook.open("template.xlsm")
    print(wb.has_vba_project)
    wb.set_vba_module_text("Module1", source)
    wb.save("result.xlsm")
    ```

openpyxl's `keep_vba` preserves VBA but does not make it editable. XL++ exposes
binary project access and module/project-property APIs. Signed projects still
require an external re-signing workflow.

## Pivots

openpyxl primarily preserves and exposes imported pivot definitions; it does
not provide a high-level pivot creation workflow. XL++ exposes pivot cache,
field, filter, grouping, layout, and generated/imported model APIs. Both require
care with PivotCharts, slicers, timelines, OLAP, and Data Model relationships.

## Features with no close openpyxl equivalent

XL++ exposes native services for:

- In-process formula calculation and dependency graphs
- Reference-aware transactional row/column structural editing
- Chart cache dependency tracking and synchronization
- Password-to-open OOXML encryption read/write and inspection
- Versioned C ABI and C# binding parity
- Enterprise feature inventory and targeted refresh/source metadata edits
- Package preservation inventory and validator tooling

## Complete feature mapping index

| Feature | openpyxl pattern | XL++ Python pattern | Important difference |
| --- | --- | --- | --- |
| Create workbook | `Workbook()` | `xlpp.Workbook()` | XL++ starts without a sheet |
| Load workbook | `load_workbook(path)` | `Workbook.open(path)` | XL++ also accepts explicit load limits/password |
| Memory I/O | `load_workbook(BytesIO(data))` / save to `BytesIO` | `load_bytes()` / `save_bytes()` | XL++ returns native `bytes` directly |
| Worksheet lookup | `wb[name]` | `wb[name]` | Similar |
| Safe lookup | membership/check names | `get_worksheet(name)` | XL++ returns `None` |
| Copy sheet | `wb.copy_worksheet(ws)` | `wb.copy_worksheet(ws, name)` | XL++ requires destination name |
| Dependency-aware rename | manual/reference-sensitive | `wb.rename_worksheet(old, new)` | XL++ returns an update report |
| Cell value | `ws["A1"] = value` | `ws["A1"].value = value` | Formula-looking strings remain text in XL++ |
| Formula | assign `"=..."` | `set_formula("=...")` | XL++ has native calculation/reporting |
| Formula cached values | `data_only=True` reads cache | `calculate_formulas()` updates cache | XL++ evaluator is broad but not complete Excel |
| Date epoch | `wb.epoch` | `wb.date_1904` | Different representation |
| Rich text | load with `rich_text=True`; rich text objects | `RichText` / `RichTextRun` | Explicit native runs |
| Style | immutable/copyable style descriptors | mutate `font()`, `fill()`, etc. | XL++ child objects are parent-owned |
| Named style | `NamedStyle`; add to workbook | `NamedStyle`; `add_named_style()` | Application syntax differs |
| Merge | `merge_cells()` | `merge_cells()` | Similar |
| Dimensions | mapping collections | `row_dimension()` / `column_dimension()` | Method-based access |
| Views/panes | properties | view methods/properties | Native model includes imported metadata |
| Print settings | properties | `page_setup()`, `page_margins()`, etc. | Method-based child access |
| Comments/hyperlinks | assign objects/properties | `set_comment()` / `set_hyperlink()` | Explicit lifecycle methods |
| Images | `Image` + `add_image()` | `xlpp.Image` + `add_image()` | XL++ accepts Python bytes and exposes anchor metadata |
| Tables | construct `Table`, then add | `add_table(name, ref)` | XL++ returns attached table |
| AutoFilter | `ws.auto_filter` | `ws.auto_filter()` | Both configure; Excel applies filtering |
| Data validation | construct then attach/range-add | collection `add(type, ref)` | XL++ returns attached validation |
| Conditional formatting | rule objects/factories | rule objects/factories or `add_rule()` | Different model names |
| Charts | family-specific chart classes + `Reference` | `ChartType` + `ChartSeries` reference strings | XL++ has imported stable-ID editors/cache sync |
| Pivots | preserve/inspect imported pivots | create/edit cache/fields/layout/filter model | XL++ offers semantic model authoring |
| Structural insert/delete | worksheet operations | worksheet or transactional workbook operations | Workbook operation rewrites broader references |
| Defined names | `DefinedName` dictionaries | `DefinedName` + scoped lookup | Similar OOXML concept, different API |
| Workbook validation | no equivalent aggregate service | `book.validate()` | XL++ report includes model/package concerns |
| Read-only large files | `load_workbook(read_only=True)` | `StreamingWorkbookReader` | Separate native streaming type |
| Write-only large files | `Workbook(write_only=True)` | `StreamingWorkbookWriter` | Explicit `close()` in XL++ |
| Numeric arrays | per-cell or third-party helpers | `write_array()` / `to_array()` | Native NumPy bridge |
| VBA preservation | `keep_vba=True` | automatic project detection/preservation | XL++ also exposes modules/binary APIs |
| VBA authoring | not supported | module/project APIs | Signatures still require re-signing |
| File encryption | not supported by core openpyxl | load/save encryption options | XL++ supports selected Office profiles |
| Formula dependencies | tokenizer/manual analysis | `dependency_graph()` | Native graph and reports |
| Chart cache refresh | manual/cache behavior limited | `synchronize_chart_caches()` | Dependency-aware option available |
| External data | package models/links | inspection/preservation APIs | XL++ adds aggregate inventory |
| Data Model/OLAP | limited preservation behavior | inspection + targeted refresh metadata | Neither is a full model execution engine |
| Power Query | no semantic authoring | inspection + selected metadata | Preservation-first |
| Slicers/timelines | limited extension preservation | inventory/preservation | No complete authoring claim |
| SmartArt/ActiveX/UserForms | preservation can be limited | inventory/preservation | No complete authoring claim |

## Migration checklist

1. Replace workbook creation/load and sheet creation semantics.
2. Change direct cell assignment to `.value` and formulas to `set_formula()`.
3. Convert style object assignment to native style accessors.
4. Replace openpyxl `Reference` chart objects with Excel reference strings.
5. Select normal, streaming, records, or NumPy APIs for each data path.
6. Validate dates, number formats, formulas, and cached values.
7. Test charts, pivots, VBA, external links, and preservation round trips.
8. Compare output using both Excel and automated interoperability tests.
