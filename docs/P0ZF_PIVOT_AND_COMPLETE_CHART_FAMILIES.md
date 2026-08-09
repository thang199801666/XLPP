# P0Z-F — PivotTable Expansion + Complete Excel Chart Families / v1.9.0

**Date:** 2026-08-08

P0Z-F expands XL++ in two areas that previously remained incomplete: generated PivotTables and chart-family coverage. The target is practical non-OLAP PivotTable authoring/round-trip plus a native model for every chart family exposed by current Microsoft Excel, while preserving the existing preservation-aware chart editor.

## Chart coverage

### Classic ChartML families

XL++ now models/generates the classic chart families and major Excel variants:

- Column and horizontal Bar, including clustered/stacked/100%-stacked grouping.
- Line and 3-D Line.
- Pie, 3-D Pie, Doughnut, Pie-of-Pie and Bar-of-Pie.
- Area and 3-D Area.
- XY Scatter with the Excel `xVal`/`yVal` payload and scatter-style semantics.
- Bubble with `xVal`/`yVal`/`bubbleSize`, scale, negative-bubble and size-representation options.
- Stock, including high-low/up-down-bar support; volume-stock compositions can be expressed as a Combo plot.
- Radar.
- Surface and 3-D Surface.
- 3-D Bar/Column.
- Combo charts through multiple `Chart::Plot` objects, explicit series ranges and primary/secondary axis groups.

The classic serializer no longer treats Scatter/Bubble as category charts: it emits two value axes and Excel-compatible XY/Bubble series payloads.

### Modern ChartEx families

The following Excel modern chart families are generated as Office ChartEx parts (`application/vnd.ms-office.chartex+xml`) and referenced through the Office 2014 `chartEx` relationship/drawing namespace:

- Histogram
- Pareto
- Box & Whisker
- Waterfall
- Funnel
- Treemap
- Sunburst
- Filled Map

Histogram binning/underflow/overflow, Box & Whisker statistics/quartile options, and Waterfall connector-line options are represented in the public plot model. The ChartEx reader reconstructs the supported modern type and cached series data after XL++ load/save.

## PivotTable coverage

P0Z-F substantially expands generated and reloaded **non-OLAP** PivotTables.

### Cache model

- Cache source range, fields and typed records.
- Shared-item generation for axis fields and typed numeric/string/boolean/error/missing values.
- Cache record read-back into the public semantic model.
- Refresh/save-data/enable-refresh lifecycle flags.
- Missing-items limit, background query, memory optimization, upgrade-on-refresh, subquery/advanced-drill flags and `refreshedBy` metadata.
- Numeric grouping and date grouping metadata (`fieldGroup/rangePr`).

### Field/layout model

- Row, column, page and multiple data fields.
- Synthetic Excel **Values** field (`x=-2`) when multiple data fields require a values axis.
- Cartesian row/column item matrices generated from visible cache members.
- Compact, Outline and Tabular layouts.
- Row/column grand totals, formatting/style flags, row/column headers/stripes and page-field layout.
- Per-field manual sorting, subtotals, blank-row/page-break behavior, hidden items, dropdowns, multi-select and selected page item.
- Repeat item labels through the Excel x14 PivotField extension.

### Filters and value display

- Caption/value/date filter types are preserved through the generic OOXML filter type model.
- Top/Bottom Count/Percent/Sum style filters are represented by the Pivot filter + Top10 payload.
- Classic `Show Values As`: difference, percent, percent difference, running total, percent of row/column/total and index.
- Excel 2010+ `Show Values As`: percent of parent/parent row/parent column, percent of running total, rank ascending and rank descending. These use the required x14 `dataField` extension rather than writing an invalid legacy `showDataAs` value.
- Base field/base item, data-field caption, aggregation and number-format ID.

## Bindings

- Python exposes the full new Chart/Pivot semantic objects directly.
- C ABI exposes all new chart types/options, combo plots, modern-chart options, Pivot cache/field/grouping/filter lifecycle, style/display controls and data-field configuration.
- C# mirrors the C ABI for the new chart and Pivot authoring surface.
- Binding/version parity is a release gate; package metadata is synchronized at **1.9.0**.

## Regression evidence

The new regression coverage verifies:

- all eight modern ChartEx types are packaged with the correct content type/relationship and reload to their semantic chart types;
- Column+Line Combo with secondary axes;
- Scatter XY semantics and Bubble size references;
- Pivot cache records, grouping, filters, hidden items, repeat labels, multi-data Values field and Cartesian axis items;
- x14 modern `Show Values As` generation and reload.

The native unit suite remains the local release gate. Microsoft Excel Desktop is not installed in this Linux validation environment, so direct Excel recovery-log/COM re-save validation remains an external interoperability gate rather than a locally verified claim.

## Explicit remaining Pivot ecosystem gaps

P0Z-F targets worksheet-source, non-OLAP PivotTables. Excel features that depend on a different subsystem are still separate roadmap items: Power Pivot/Data Model and OLAP cube hierarchies, slicers/timelines, calculated fields/items and PivotCharts. Those should not be described as completed by this milestone.
