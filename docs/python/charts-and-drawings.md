# Charts and drawings

## Create a chart

Prepare source data first:

```python
sheet.append(["Month", "Sales", "Target"])
sheet.append(["Jan", 120, 100])
sheet.append(["Feb", 175, 150])
sheet.append(["Mar", 150, 160])
```

Create and attach a chart:

```python
chart = xlpp.Chart(xlpp.ChartType.BAR)
chart.title = "Monthly sales"
chart.x_axis_title = "Month"
chart.y_axis_title = "Units"
chart.grouping = xlpp.ChartGrouping.CLUSTERED
chart.width = 640
chart.height = 360

sales = chart.add_series(xlpp.ChartSeries("Sales"))
sales.categories_reference = "'Sales'!$A$2:$A$4"
sales.values_reference = "'Sales'!$B$2:$B$4"

target = chart.add_series(xlpp.ChartSeries("Target"))
target.categories_reference = "'Sales'!$A$2:$A$4"
target.values_reference = "'Sales'!$C$2:$C$4"

sheet.add_chart(chart)
```

References are Excel formula references. Quote sheet names and use absolute
coordinates when charts may be moved independently of their source.

## Chart families

`ChartType` includes classic and modern model values:

- Bar/column, line, area, pie, doughnut, scatter, bubble, and radar
- Stock, surface, 3D bar/line/area/pie, projected pie
- Histogram, Pareto, waterfall, box-and-whisker, funnel
- Treemap, sunburst, and filled map

Support differs by family: classic charts have broad generation and editing;
some modern chart types are primarily imported-model and preservation paths.
Check [Feature/API coverage](api-coverage.md) before generating a specialized
chart from scratch.

## Series

```python
series = xlpp.ChartSeries("Revenue")
series.categories_reference = "'Data'!$A$2:$A$13"
series.values_reference = "'Data'!$B$2:$B$13"
marker = xlpp.ChartMarkerFormat()
marker.present = True
marker.symbol = "circle"
marker.size = 7
series.marker_format = marker
series.smooth = False
chart.add_series(series)
```

The model exposes title/category/value/bubble references, cached data, marker
formatting, data labels, trendlines, error bars, per-point formatting, and
secondary-axis metadata.

## Trendlines and error bars

```python
trend = xlpp.ChartTrendline()
trend.type = xlpp.ChartTrendlineType.LINEAR
trend.display_equation = True
trend.display_r_squared = True
series.trendlines = [trend]

bars = xlpp.ChartErrorBars()
bars.direction = xlpp.ChartErrorBarDirection.Y
bars.bar_type = xlpp.ChartErrorBarType.BOTH
bars.value_type = xlpp.ChartErrorValueType.PERCENTAGE
bars.value = 5.0
series.error_bars = [bars]
```

When editing imported charts, prefer worksheet methods targeting the chart
stable ID and series index. Those methods patch the selected ChartML subtree
and preserve unrelated chart XML.

## Chart and plot formatting

```python
fill = xlpp.ChartFillFormat()
fill.present = True
fill.kind = xlpp.ChartFillKind.SOLID
fill.color.kind = xlpp.ChartColorKind.SRGB
fill.color.value = "FF1F4E78"
chart.chart_area_fill_format = fill

line = xlpp.ChartLineFormat()
line.present = True
line.width_points = 1.5
line.color.kind = xlpp.ChartColorKind.SRGB
line.color.value = "FF000000"
chart.chart_area_line_format = line
```

Formatting models include solid/gradient/pattern fills, DrawingML color
transforms, custom dash stops, text styles, legend/plot manual layouts, and
chart theme resolution.

## Axes and combined plots

```python
for axis in chart.axes:
    print(axis.id, axis.kind, axis.position)

plot = chart.primary_plot()
print(plot.type)
```

`add_plot()` builds combined charts and can select secondary axes. Imported
charts retain native axis IDs so selective edits do not flatten combined or
secondary-axis structures.

## 3D charts

```python
view = xlpp.ChartView3D()
view.present = True
view.has_rotation_x = True
view.rotation_x = 20
view.has_rotation_y = True
view.rotation_y = 30
chart.view_3d = view
```

Floor, side-wall, and back-wall models are available through
`floor_format`, `side_wall_format`, and `back_wall_format`.

## Synchronize chart caches

Excel charts often contain cached copies of formula-backed series values.
After changing source cells:

```python
options = xlpp.ChartCacheSyncOptions()
options.synchronize_titles = True
options.synchronize_categories = True
options.synchronize_values = True
options.changed_references_only = True

report = book.synchronize_chart_caches(options)
print(report.caches_updated)
print(report.references_skipped)
print(report.warnings)
```

!!! tip "Cache synchronization is explicit"

    Update caches only when consumers need cached chart data or when Excel will
    not recalculate it immediately. `changed_references_only` avoids rewriting
    caches whose dependency snapshots did not change.

## Edit imported charts

```python
for chart in sheet.charts:
    print(chart.stable_id, chart.title, chart.source_chart_part)

chart_id = sheet.charts[0].stable_id
sheet.set_chart_title(chart_id, "Updated title")
sheet.set_chart_style(chart_id, 10)
```

Worksheet methods cover imported title/rich text, plot and legend layout,
axes, gridlines, series metadata, labels, trendlines, error bars, caches,
point formatting, auxiliary objects, 3D view, and area/wall formatting.
Methods return `False` when the stable ID or requested child is not found.

## Images

```python
image = xlpp.Image.from_file("logo.png", "A1")
image.width_pixels = 180
image.height_pixels = 64
sheet.add_image(image)
```

In-memory image:

```python
with open("logo.png", "rb") as stream:
    image = xlpp.Image("A1", stream.read(), "png")

sheet.add_image(image)
```

Authored image formats are PNG and JPEG. Imported images expose anchor and
package-origin metadata:

```python
for image in sheet.images:
    print(image.stable_id, image.anchor_info.type, image.source_media_part)
```

## Preservation tips

- Use stable-ID worksheet edits for imported charts.
- Do not replace a chart merely to change one field.
- Synchronize caches after changing referenced values when required.
- Validate workbooks containing modern chart extensions in the target Excel
  version.
- Keep representative package before/after fixtures for complex DrawingML.
