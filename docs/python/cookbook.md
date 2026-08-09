# Cookbook and tips

Practical patterns for common XL++ Python workflows.

## Build a report with a styled header

```python
import xlpp

book = xlpp.Workbook()
sheet = book.add_worksheet("Report")
sheet.append(["Product", "Units", "Price", "Total"])

for cell in sheet.range("A1:D1").cells():
    cell.font().bold = True
    cell.font().color().set_argb("FFFFFFFF")
    cell.fill().pattern_type = "solid"
    cell.fill().foreground().set_argb("FF1F4E78")
    cell.alignment().horizontal = "center"

sheet.freeze_panes("A2")
sheet.auto_filter().reference = "A1:D1000"
```

## Export iterable records efficiently

```python
writer = xlpp.StreamingWorkbookWriter("export.xlsx")
sheet = writer.add_worksheet("Data")
sheet.append(["id", "name", "amount"])

for record in query_results:
    sheet.append([record.id, record.name, record.amount])

writer.close()
```

Use the streaming writer when you do not need random access or rich workbook
features after rows are appended.

## Return a workbook from a web endpoint

```python
book = xlpp.Workbook()
sheet = book.add_worksheet("Result")
sheet.append(["Status", "OK"])

payload = book.save_bytes()

# Framework-specific response API:
return Response(
    payload,
    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    headers={"Content-Disposition": 'attachment; filename="result.xlsx"'},
)
```

## Load a workbook received as bytes

```python
book = xlpp.Workbook()
book.load_bytes(uploaded_bytes)
```

Set `LoadOptions` limits before accepting untrusted uploads.

## Preserve a template while replacing data

```python
book = xlpp.Workbook.open("template.xlsm")
sheet = book["Input"]

sheet.range("A2:F10000").clear()
for row in new_rows:
    sheet.append(row)

options = xlpp.SaveOptions()
options.atomic_write = True
options.validate_before_save = True
book.save("generated.xlsm", options)
```

!!! tip

    Clear only the intended data range. Avoid recreating the workbook when the
    template contains VBA, drawings, charts, pivots, or enterprise parts.

## Copy a style from another cell

```python
source = sheet["A1"].style()
target = sheet["A2"].style()

target.font().name = source.font().name
target.font().size = source.font().size
target.font().bold = source.font().bold
target.number_format = source.number_format
```

For widespread reuse, create a named style instead of copying properties one
at a time.

## Apply formatting to a range

```python
def format_currency(cell):
    cell.set_number_format("$#,##0.00")

sheet.range("D2:D1000").for_each(format_currency)
```

## Add totals without losing formulas

```python
last_row = sheet.max_row
total_row = last_row + 1
sheet.cell(total_row, 1).value = "Total"
sheet.cell(total_row, 4).set_formula(f"=SUM(D2:D{last_row})")
```

## Rename a sheet safely

```python
report = book.rename_worksheet("Raw Data", "Input")
if not report.success:
    raise RuntimeError(report.warnings)
```

This is safer than directly assigning `sheet.name` when the workbook has
dependent references.

## Insert rows transactionally

```python
options = xlpp.StructuralEditOptions()
options.transactional = True
options.update_defined_names = True
options.fail_on_invalid_reference = True

report = book.insert_rows("Input", 2, 10, options)
if not report.success:
    raise RuntimeError(report.warnings)
```

## Recalculate only affected formulas

```python
changed = xlpp.CalculationCell()
changed.sheet = "Input"
changed.cell = "B2"

options = xlpp.CalculationOptions()
options.changed_cells = [changed]
options.recursive_dependencies = True

report = book.calculate_formulas(options)
```

## Keep chart caches current

```python
options = xlpp.ChartCacheSyncOptions()
options.changed_references_only = True

report = book.synchronize_chart_caches(options)
print(report.caches_updated, report.references_unchanged)
```

## Add a dropdown from a hidden list sheet

```python
lists = book.add_worksheet("Lists")
for index, value in enumerate(["Open", "Closed", "Pending"], start=1):
    lists.cell(index, 1).value = value

book.add_defined_name(xlpp.DefinedName("StatusValues", "'Lists'!$A$1:$A$3"))

validation = sheet.data_validations.add(
    xlpp.DataValidationType.LIST,
    "D2:D1000",
)
validation.formula1 = "=StatusValues"
```

## Read rows until a condition is met

```python
reader = xlpp.StreamingWorkbookReader("input.xlsx")

def consume(row_number, cells):
    values = [cell.value for cell in cells]
    process(values)
    return row_number < 100_000

reader.for_each_row("Data", consume)
```

## Encrypt an output workbook

```python
options = xlpp.SaveOptions()
options.encryption_password = password
options.encryption_mode = xlpp.OfficeEncryptionMode.AGILE_AES256_SHA512
book.save("secure.xlsx", options)
```

## Diagnose enterprise content before editing

```python
enterprise = book.inspect_enterprise_features()
for feature in enterprise.features:
    if not feature.semantic_editable:
        print("Preservation-only:", feature.kind, feature.part_name)
```

## Validate a round trip

```python
before = xlpp.Workbook.open("input.xlsx")
before["Input"]["A1"].value = "Updated"

options = xlpp.SaveOptions()
options.atomic_write = True
options.validate_before_save = True
before.save("output.xlsx", options)

check = xlpp.Workbook.open("output.xlsx")
report = check.validate()
assert not report.issues, [issue.message for issue in report.issues]
```

For advanced workbooks, also run `xlpp-package-validator` against the before
and after files.

## Tips and traps

### Do

- Keep parent workbook objects alive while using native child handles.
- Use explicit formula setters rather than assigning formula-looking strings.
- Use bulk and streaming APIs for large datasets.
- Set realistic load limits for untrusted files.
- Save complex template edits to a new output path.
- Inspect operation reports and warnings.
- Test with the target Excel version and representative files.

### Avoid

- Styling millions of empty cells.
- Rebuilding imported charts to change one title or format.
- Assuming worksheet protection encrypts data.
- Assuming refresh-on-load executes external queries inside XL++.
- Assuming binding parity means every Excel UI feature is authorable.
- Editing signed VBA without a re-signing plan.

