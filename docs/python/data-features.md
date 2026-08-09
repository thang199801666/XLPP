# Tables, filters, validation, and conditional formatting

## Tables

Write headers and data before adding a table:

```python
sheet.append(["Product", "Region", "Units", "Status"])
sheet.append(["Widget", "East", 100, "Open"])
sheet.append(["Gadget", "West", 50, "Closed"])

table = sheet.add_table("SalesTable", "A1:D3")
table.display_name = "SalesData"
table.show_header_row = True
table.show_totals_row = False
table.style_info.name = "TableStyleMedium4"
table.style_info.show_row_stripes = True
```

Add or inspect columns:

```python
column = table.add_column("Status")
for column in table.columns:
    print(column.name)
```

!!! warning "Table names and ranges"

    Table names must be unique and valid Excel identifiers. The table range
    should cover its header and all intended data rows.

## AutoFilter

```python
filter_ = sheet.auto_filter()
filter_.reference = "A1:D100"

column = filter_.column(3)
column.add_value("Open")
column.add_value("Pending")
```

The model supports imported simple values, custom filters, Top10, dynamic,
color, icon, and date-group filters. Use the typed filter fields when editing
those forms; preserving an imported filter does not require rewriting it.

Clear or disable filters through the `AutoFilter` and column APIs rather than
only blanking the range reference.

## Data validation

List validation:

```python
validation = sheet.data_validations.add(
    xlpp.DataValidationType.LIST,
    "D2:D100",
)
validation.formula1 = '"Open,Closed,Pending"'
validation.allow_blank = True
validation.show_error_message = True
validation.error_title = "Invalid status"
validation.error = "Choose a value from the list."
```

Numeric validation:

```python
validation = sheet.data_validations.add(
    xlpp.DataValidationType.WHOLE,
    "C2:C100",
)
validation.op = xlpp.DataValidationOperator.BETWEEN
validation.formula1 = "0"
validation.formula2 = "1000000"
```

Validation types include list, whole, decimal, date, time, text length, and
custom. Formulas follow Excel syntax and may reference named ranges.

## Conditional formatting

Formula rule:

```python
rule = sheet.conditional_formatting.add_rule(
    "C2:C100",
    xlpp.ConditionalRuleType.FORMULA,
    "C2>100",
)
```

Cell comparison:

```python
rule = sheet.conditional_formatting.add_rule(
    "C2:C100",
    xlpp.ConditionalRule.cell_is(
        xlpp.ConditionalOperator.GREATER_THAN,
        "100",
    ),
)
```

The public model also represents color scales, data bars, and icon sets.
Imported differential styles are preserved and exposed through rule fields.

## Protection interaction

Validation and filters may appear disabled in Excel when worksheet protection
does not allow the corresponding operation. Configure sheet protection and
filter/table features together, then test the result in the target Excel
version.

## openpyxl comparison

| Task | openpyxl | XL++ |
| --- | --- | --- |
| Add table | `ws.add_table(table)` | `ws.add_table(name, ref)` |
| Filter range | `ws.auto_filter.ref = ref` | `ws.auto_filter().reference = ref` |
| Add validation | create `DataValidation`, then `ws.add_data_validation()` | `ws.data_validations.add(type, ref)` |
| Add CF formula | `ws.conditional_formatting.add(ref, rule)` | `ws.conditional_formatting.add_rule(ref, type, formula)` |

See [Migrate from openpyxl](openpyxl-migration.md) for complete side-by-side
recipes.
