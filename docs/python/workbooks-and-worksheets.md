# Workbooks and worksheets

## Create and save

```python
import xlpp

book = xlpp.Workbook()
sales = book.add_worksheet("Sales")
inventory = book.add_worksheet("Inventory")

book.properties.title = "Operations workbook"
book.properties.creator = "XL++"
book.save("operations.xlsx")
```

Worksheet names must be valid and unique. Duplicate or invalid names raise
`ValueError`.

## Open an existing workbook

```python
book = xlpp.Workbook.open("operations.xlsx")
```

Equivalent explicit form:

```python
book = xlpp.Workbook()
book.load("operations.xlsx")
```

Use `Workbook.open()` when creating and loading can be one expression. Use
`load()` when reusing a workbook instance or configuring it before load.

## Access sheets

```python
first = book[0]
sales = book["Sales"]

print(book.sheet_names)
print(book.sheet_count)
print(book.active.name)

for sheet in book:
    print(sheet.name, sheet.dimensions())
```

Safe lookup returns `None`:

```python
sheet = book.get_worksheet("Optional")
if sheet is None:
    sheet = book.add_worksheet("Optional")
```

## Copy, rename, and remove

```python
copy = book.copy_worksheet(book["Sales"], "Sales Copy")
book.rename_worksheet("Sales Copy", "Archive")
book.remove_worksheet("Archive")
```

Prefer `book.rename_worksheet()` over assigning `sheet.name` when dependent
formulas, names, chart references, pivots, or hyperlinks should be translated.
The returned report contains update counters and warnings.

```python
report = book.rename_worksheet("Sales", "Revenue")
print(report.references_updated)
print(report.warnings)
```

## Memory I/O

```python
payload = book.save_bytes()

received = xlpp.Workbook()
received.load_bytes(payload)
```

This is useful for web responses, object storage, queues, database blobs, and
tests that should not touch the filesystem.

## Save options

```python
options = xlpp.SaveOptions()
options.compression_level = xlpp.CompressionLevel.BEST
options.compression_strategy = xlpp.CompressionStrategy.DEFAULT
options.parallel_workers = 4
options.parallel_sheets = True
options.atomic_write = True
options.validate_before_save = True
options.calculate_formulas_before_save = False
options.synchronize_chart_caches = False

book.save("operations.xlsx", options)
```

!!! tip "Use atomic writes for production files"

    `atomic_write` protects the destination from partial replacement when save
    fails. `durable_write` requests stronger persistence guarantees at a higher
    I/O cost.

## Load options and limits

```python
options = xlpp.LoadOptions()
options.max_file_bytes = 512 * 1024 * 1024
options.max_entry_bytes = 128 * 1024 * 1024
options.max_total_bytes = 1024 * 1024 * 1024
options.max_entries = 50_000

book = xlpp.Workbook.open("input.xlsx", options)
```

These limits reduce exposure to malformed or unexpectedly expanded ZIP
packages. Use `lenient = True` only for deliberate recovery workflows.

## Workbook metadata

```python
book.properties.title = "Quarterly report"
book.properties.subject = "Revenue"
book.properties.creator = "Reporting service"
book.properties.keywords = "finance,quarterly"

custom = xlpp.CustomProperty("Environment", "Production")
book.custom_properties().add(custom)

book.calc_properties().full_calc_on_load = True
book.date_1904 = False
```

## Defined and named styles

```python
name = xlpp.DefinedName("TaxRate", "0.2")
book.add_defined_name(name)

style = xlpp.NamedStyle("Header")
book.add_named_style(style)
book.apply_named_style(book["Sales"]["A1"], "Header")
```

Defined names can be workbook-scoped or worksheet-scoped. Prefer the explicit
`local_sheet_id` argument when resolving duplicate scoped names.

## Validation and diagnostics

```python
report = book.validate()
for issue in report.issues:
    print(issue.severity, issue.message)
```

Validation is complementary to opening the file in Excel. For complex round
trips, also use XL++ package validation and compare before/after packages.

## Object lifetime

Worksheets, cells, styles, charts, and similar child objects are owned by their
parent workbook or worksheet. Keep the owner alive:

```python
book = xlpp.Workbook.open("input.xlsx")
cell = book["Sales"]["A1"]
print(cell.value)  # book is still alive
```

Do not retain a child after clearing the workbook or removing its owner.
