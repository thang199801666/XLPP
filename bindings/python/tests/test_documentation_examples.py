"""Runtime smoke tests for the public examples used by the Python docs."""

from datetime import date

import xlpp


def test_documented_workbook_cell_style_and_data_features():
    book = xlpp.Workbook()
    sheet = book.add_worksheet("Sales")
    sheet.append(["Product", "Units", "Price", "Date"])
    sheet.append(["Widget", 100, 9.99, date(2026, 8, 9)])
    sheet["E2"].set_formula("=B2*C2")

    sheet["A1"].font().bold = True
    sheet["A1"].font().color().set_argb("FFFFFFFF")
    sheet["A1"].fill().pattern_type = "solid"
    sheet["A1"].fill().foreground().set_argb("FF1F4E78")
    sheet["A1"].alignment().horizontal = "center"
    sheet["C2"].set_number_format("$#,##0.00")
    sheet.freeze_panes("A2")
    sheet.column_dimension("A").width = 24
    sheet.set_print_area("A1:E100")

    named = xlpp.NamedStyle("Header")
    named.style().font().bold = True
    book.add_named_style(named)
    book.apply_named_style(sheet["A1"], "Header")

    table = sheet.add_table("SalesTable", "A1:E2")
    table.display_name = "SalesData"
    table.style_info.name = "TableStyleMedium4"
    sheet.auto_filter().reference = "A1:E2"

    validation = sheet.data_validations.add(
        xlpp.DataValidationType.LIST, "F2:F100"
    )
    validation.formula1 = '"Open,Closed,Pending"'
    validation.allow_blank = True
    validation.op = xlpp.DataValidationOperator.BETWEEN

    rule = xlpp.ConditionalRule.cell_is(
        xlpp.ConditionalOperator.GREATER_THAN, "100"
    )
    sheet.conditional_formatting.add_rule("B2:B100", rule)

    payload = book.save_bytes()
    loaded = xlpp.Workbook()
    loaded.load_bytes(payload)
    assert loaded["Sales"]["A2"].value == "Widget"


def test_documented_chart_and_pivot_models_round_trip():
    book = xlpp.Workbook()
    sheet = book.add_worksheet("Sales")
    sheet.append(["Product", "Region", "Units", "Revenue"])
    sheet.append(["Widget", "East", 100, 999.0])

    chart = xlpp.Chart(xlpp.ChartType.BAR)
    chart.title = "Sales"
    series = chart.add_series(xlpp.ChartSeries("Revenue"))
    series.categories_reference = "'Sales'!$A$2:$A$2"
    series.values_reference = "'Sales'!$D$2:$D$2"
    marker = xlpp.ChartMarkerFormat()
    marker.present = True
    marker.symbol = "circle"
    marker.size = 7
    series.marker_format = marker
    sheet.add_chart(chart)

    pivot = xlpp.PivotTable("SalesPivot")
    pivot.location = "F2"
    pivot.layout = xlpp.PivotLayout.TABULAR
    pivot.cache.source_data = "'Sales'!$A$1:$D$2"
    pivot.cache.refresh_on_load = True
    for name in ["Product", "Region", "Units", "Revenue"]:
        pivot.cache.add_field(name)
    pivot.add_row_field("Product")
    pivot.add_column_field("Region")
    pivot.add_data_field("Revenue", "sum")
    sheet.add_pivot_table(pivot)

    payload = book.save_bytes()
    loaded = xlpp.Workbook()
    loaded.load_bytes(payload)
    assert loaded["Sales"].chart_count == 1
    assert len(loaded["Sales"].pivot_tables) == 1


def test_documented_calculation_dependency_and_translation_helpers():
    book = xlpp.Workbook()
    sheet = book.add_worksheet("Sales")
    sheet["A1"].value = 2
    sheet["B1"].set_formula("=A1*2")

    changed = xlpp.CalculationCell()
    changed.sheet = "Sales"
    changed.cell = "A1"
    options = xlpp.CalculationOptions()
    options.changed_cells = [changed]
    report = book.calculate_formulas(options)
    assert report.success
    assert sheet["B1"].value == 4.0

    graph = book.dependency_graph()
    assert graph.depends_on("Sales", "B1", "Sales", "A1")

    edit = xlpp.StructuralEdit(
        "Sales", xlpp.StructuralEditKind.INSERT_ROWS, 1, 1
    )
    translated = xlpp.translate_formula_references("=A1", "Sales", edit)
    assert translated.changed


def test_documented_streaming_and_enterprise_inspection(tmp_path):
    path = tmp_path / "stream.xlsx"
    writer = xlpp.StreamingWorkbookWriter(
        str(path), xlpp.SharedStringMode.BOUNDED_LRU, 128
    )
    sheet = writer.add_worksheet("Data")
    sheet.append(["id", "value"])
    sheet.append([1, "one"])
    writer.close()

    reader = xlpp.StreamingWorkbookReader(str(path))
    rows = list(reader.worksheet("Data"))
    assert rows[1][1][1].value == "one"

    book = xlpp.Workbook.open(str(path))
    assert book.inspect_enterprise_features().features == []
    assert book.inspect_external_data().has_connections is False
    assert book.inspect_data_model().present is False


def test_documented_advanced_model_snippets_construct():
    cell = xlpp.Cell("A1")
    cell.set_error(xlpp.CellError.DIVISION_BY_ZERO)
    assert cell.is_error()

    series = xlpp.ChartSeries("Revenue")
    trend = xlpp.ChartTrendline()
    trend.type = xlpp.ChartTrendlineType.LINEAR
    trend.display_equation = True
    series.trendlines = [trend]

    bars = xlpp.ChartErrorBars()
    bars.direction = xlpp.ChartErrorBarDirection.Y
    bars.bar_type = xlpp.ChartErrorBarType.BOTH
    bars.value_type = xlpp.ChartErrorValueType.PERCENTAGE
    bars.value = 5.0
    series.error_bars = [bars]

    fill = xlpp.ChartFillFormat()
    fill.present = True
    fill.kind = xlpp.ChartFillKind.SOLID
    fill.color.kind = xlpp.ChartColorKind.SRGB
    fill.color.value = "FF1F4E78"
    assert fill.color.value == "FF1F4E78"

    field = xlpp.PivotField("Date")
    field.grouping.kind = xlpp.PivotGroupingKind.DATE
    field.grouping.date_part = xlpp.PivotDatePart.MONTHS
    assert field.grouping.active
