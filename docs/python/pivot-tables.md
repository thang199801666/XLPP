# Pivot tables

XL++ exposes pivot table, cache, field, field-reference, filter, grouping, and
layout models. It can generate model-backed pivots and inspect/edit imported
pivots while preserving related package content.

## Create a pivot table

Prepare rectangular source data:

```python
sheet.append(["Product", "Region", "Units", "Revenue"])
sheet.append(["Widget", "East", 100, 999.0])
sheet.append(["Widget", "West", 50, 499.5])
sheet.append(["Gadget", "East", 75, 1499.25])
```

Configure the pivot and its cache:

```python
pivot = xlpp.PivotTable("SalesPivot")
pivot.location = "F2"
pivot.layout = xlpp.PivotLayout.TABULAR
pivot.style_name = "PivotStyleMedium9"
pivot.show_row_stripes = True

pivot.cache.source_data = "'Sales'!$A$1:$D$4"
pivot.cache.refresh_on_load = True
pivot.cache.save_data = True

for name in ["Product", "Region", "Units", "Revenue"]:
    pivot.cache.add_field(name)

pivot.add_row_field("Product")
pivot.add_column_field("Region")
pivot.add_data_field("Revenue", "sum")

sheet.add_pivot_table(pivot)
```

The cache field count must match the width of the source range.

## Cache records

When embedding pivot cache data explicitly:

```python
pivot.cache.add_record(["Widget", "East", "100", "999.0"])
pivot.cache.add_record(["Widget", "West", "50", "499.5"])
```

Use `clear_records()` when the workbook should refresh from the source rather
than retain old cache records.

## Field configuration

```python
field = pivot.row_fields[0]
field.show_all = True
field.sort_type = "ascending"
field.repeat_item_labels = True
field.add_subtotal("sum")
field.set_item_hidden(2, True)
```

Page fields, column fields, and data field references expose captions,
subtotal modes, number-format IDs, show-data-as settings, base field/item,
selection, and hidden-item metadata.

## Grouping

```python
field = pivot.row_fields[0]
grouping = field.grouping
grouping.kind = xlpp.PivotGroupingKind.DATE
grouping.date_part = xlpp.PivotDatePart.MONTHS
```

Numeric and date grouping fields include start/end values, interval, and date
parts. Ensure cache values are compatible with the selected grouping.

## Filters

```python
filter_ = xlpp.PivotFilter()
filter_.field_index = 1
filter_.type = "captionEqual"
filter_.value1 = "East"
pivot.add_filter(filter_)
```

The filter model exposes value, caption, date, and Top10-related metadata used
by native pivot definitions.

## Imported pivots

```python
book = xlpp.Workbook.open("existing-pivot.xlsx")
sheet = book["Report"]

for pivot in sheet.pivot_tables:
    print(pivot.name, pivot.location, pivot.cache.source_data)
```

Unchanged imported pivots are preservation-aware. Mutating the semantic pivot
collection opts into model-driven output for the affected pivot domain.

!!! warning "Pivot ecosystem boundary"

    PivotCharts, slicers, timelines, OLAP caches, and Data Model connections
    span several package parts and extension schemas. XL++ can inventory and
    preserve these features and exposes selected targeted metadata edits, but
    does not claim complete Excel UI-level semantic authoring for them.

## Refresh behavior

```python
pivot.cache.refresh_on_load = True
pivot.cache.enable_refresh = True
pivot.cache.background_query = False
```

Setting refresh-on-load asks Excel to update from the source when opened. It
does not make XL++ connect to external databases or execute Power Query.

## Common problems

### Source width mismatch

If save reports that pivot source width does not match cache field count, add
one cache field for every source column or correct the source range.

### Stale cache

Set `refresh_on_load`, update cache records explicitly, or use an Excel refresh
workflow. Chart cache synchronization is separate from pivot cache refresh.

### Unsupported enterprise source

Inspect the workbook with `inspect_external_data()`, `inspect_data_model()`,
and `inspect_enterprise_features()` before making structural changes.
