"""Parity regression for native workbook services added after the first bindings."""

import xlpp


def test_formula_external_resolver_and_dependency_graph():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Calc")
    ws["A1"].set_formula("='[External.xlsx]Data'!A1+1")

    options = xlpp.CalculationOptions()
    options.external_reference_resolver = (
        lambda workbook, sheet, address: 41.0
        if (workbook, sheet, address) == ("External.xlsx", "Data", "A1")
        else None
    )
    report = wb.calculate_formulas(options)
    assert report.success
    assert report.external_references_resolved == 1
    assert ws["A1"].value == 42.0

    ws2 = wb.add_worksheet("Consumer")
    ws2["A1"].set_formula("=Calc!A1")
    graph = wb.dependency_graph()
    assert graph.depends_on("Consumer", "A1", "Calc", "A1")
    assert graph.precedents_of("Consumer", "A1")
    assert graph.dependents_of("Calc", "A1")


def test_structural_rename_validation_chart_tracking_and_vba():
    wb = xlpp.Workbook()
    data = wb.add_worksheet("Data")
    data["A1"].value = 10
    calc = wb.add_worksheet("Calc")
    calc["A1"].set_formula("=Data!A1")

    rename = wb.rename_worksheet("Data", "Input")
    assert rename.success
    assert "Input" in calc["A1"].formula

    structural = xlpp.StructuralEditOptions()
    structural.transactional = True
    result = wb.insert_rows("Input", 1, 1, structural)
    assert result.success

    cache = wb.synchronize_chart_caches(xlpp.ChartCacheSyncOptions())
    assert cache.success
    wb.reset_chart_cache_dependency_tracking()
    assert wb.tracked_chart_cache_dependency_count == 0

    validation = wb.validate()
    assert validation.ok
    validation_options = xlpp.WorkbookValidationOptions()
    validation_options.validate_charts = True
    assert validation_options.validate_charts is True

    external = wb.inspect_external_data()
    assert external.has_external_workbooks is False
    assert external.has_connections is False
    assert external.has_query_tables is False

    ws = wb.get_worksheet("Input")
    ws.vba_code_name = "InputSheet"
    props = xlpp.VbaProjectProperties()
    props.name = "ParityMacros"
    props.description = "Python VBA parity"
    props.help_file = "parity.chm"
    props.help_context_id = 9
    props.constants = "PythonBinding = 1"
    wb.vba_project_properties = props
    wb.set_vba_document_module_text("InputSheet", "Private Sub Worksheet_Activate()\r\nEnd Sub\r\n")
    wb.set_vba_class_module_text("ParityClass", "Public Function Value() As Long\r\nValue = 1\r\nEnd Function\r\n", True, True)
    wb.set_vba_module_text("Module1", "Sub Hello()\r\nEnd Sub\r\n")
    assert wb.has_vba_project and wb.vba_source_editable
    assert wb.vba_project_properties.name == "ParityMacros"
    assert wb.vba_project_properties.constants == "PythonBinding = 1"
    assert "Sub Hello()" in wb.vba_module_text("Module1")
    assert any(module.name == "Module1" for module in wb.vba_modules)
    cls = next(module for module in wb.vba_modules if module.name == "ParityClass")
    assert cls.type == xlpp.VbaModuleType.CLASS and cls.read_only and cls.private_module
    assert isinstance(wb.vba_project_bytes, bytes) and len(wb.vba_project_bytes) > 512
    assert not wb.has_vba_signature
    assert wb.remove_vba_module("Module1")


def test_memory_io_and_save_safety_options(tmp_path):
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Memory")
    ws["A1"].value = "roundtrip"

    save_options = xlpp.SaveOptions()
    save_options.atomic_write = True
    save_options.validate_before_save = True
    payload = wb.save_bytes(save_options)
    assert isinstance(payload, bytes) and payload
    assert wb.sheet_count == 1
    assert wb.strict_namespaces is False
    assert not wb.diagnostics.had_errors()

    path = tmp_path / "native_parity_in_place.xlsx"
    wb.save_in_place(str(path), save_options)
    in_place = xlpp.Workbook()
    in_place.load_in_place(str(path), xlpp.LoadOptions())
    assert in_place.sheet_count == 1
    assert in_place.get_worksheet("Memory")["A1"].value == "roundtrip"

    progress = []
    load_options = xlpp.LoadOptions()
    load_options.progress = lambda done, total: progress.append((done, total))
    loaded = xlpp.Workbook()
    loaded.load_bytes(payload, load_options)
    assert loaded["Memory"]["A1"].value == "roundtrip"
    assert progress


def test_workbook_collection_and_preservation_views():
    wb = xlpp.Workbook()
    first = wb.add_worksheet("First")
    wb.add_worksheet("Second")
    assert wb.date1904 is False
    wb.date1904 = True
    assert wb.date1904 is True
    assert [sheet.name for sheet in wb.worksheets] == ["First", "Second"]
    assert [sheet.name for sheet in wb] == ["First", "Second"]
    assert wb.preserved_relationships == []


def test_python_exposes_full_worksheet_model_views():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Model")
    ws["A1"].value = 10
    ws["B1"].value = 20
    ws.row_dimension(1).height = 24
    ws.column_dimension(1).width = 18
    assert 1 in ws.row_dimensions
    assert 1 in ws.column_dimensions
    assert ws.rows and ws.rows[0].number == 1
    assert len(ws.cells) >= 2
    assert ws.iter_rows(1, 1, 1, 2)[0] == [10.0, 20.0]
    extents = xlpp.WorksheetExtents(1, 1, 1, 2)
    assert extents.max_column == 2
    assert xlpp.Row(ws, 1).number == 1
    assert ws.loaded_image_count == 0
    assert ws.loaded_chart_count == 0
    assert ws.appended_image_count == 0
    assert ws.appended_chart_count == 0
    ws.mark_dirty()
    assert ws.dirty()
    ws.clear_dirty()
    assert not ws.dirty()


def test_python_exposes_remaining_worksheet_native_api():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("NativeApi")
    ws.print_titles_rows = "1:2"
    ws.print_titles_cols = "A:B"
    assert ws.print_titles_rows == "1:2"
    assert ws.print_titles_cols == "A:B"

    pivot = xlpp.PivotTable("Imported")
    added = ws.add_loaded_pivot_table(pivot)
    assert added.name == "Imported"
    assert ws.loaded_pivot_count == 1
    assert ws.generated_pivot_start == 1

    assert ws.remove_chart_axis_gridlines("missing", 1, True) is False
    assert ws.remove_chart_data_table("missing") is False
    assert ws.set_chart_series_trendline_line_format(
        "missing", 0, 0, xlpp.ChartLineFormat()
    ) is False
    edit = xlpp.StructuralEdit("NativeApi", xlpp.StructuralEditKind.INSERT_ROWS, 1, 1)
    report = ws.apply_structural_edit(edit)
    assert isinstance(report, xlpp.WorksheetStructuralEditReport)
    assert report.cells_moved >= 0
    assert hasattr(ws, "set_chart_title_rich_text")
    assert hasattr(ws, "set_chart_axis_title_rich_text")
    assert hasattr(ws, "set_chart_plot_high_low_lines")
    assert hasattr(ws, "replace_image")


def test_python_exposes_chart_format_models():
    line = xlpp.ChartLineFormat()
    line.present = True
    line.width_points = 2.5
    line.color = xlpp.ChartColor()
    line.color.kind = xlpp.ChartColorKind.SRGB
    line.color.value = "FF0000"

    cache = xlpp.ChartSeriesCache()
    cache.present = True
    cache.point_count = 2
    cache.points = [xlpp.ChartCachePoint()]
    cache.points[0].index = 0
    cache.points[0].value = "10"
    assert cache.effective_point_count == 2
    assert cache.valid()

    labels = xlpp.ChartDataLabels()
    labels.present = True
    labels.show_value = True
    assert labels.show_value

    layout = xlpp.ChartManualLayout()
    layout.present = True
    layout.has_width = True
    layout.width = 0.5
    assert layout.width == 0.5


def test_python_exposes_chart_and_series_native_models():
    chart = xlpp.Chart(xlpp.ChartType.LINE)
    chart.title = "Revenue"
    rich = xlpp.ChartRichText()
    rich.present = True
    run = xlpp.ChartTextRun()
    run.text = "Revenue"
    run.bold = True
    rich.runs = [run]
    chart.title_rich_text = rich
    assert chart.title == "Revenue"
    assert chart.primary_plot_or_none() is None
    assert chart.primary_plot_or_null() is None
    assert xlpp.Chart.is_modern_type(xlpp.ChartType.HISTOGRAM) is True

    series = xlpp.ChartSeries("2026")
    series.values_reference = "Data!$B$2:$B$4"
    series.categories_reference = "Data!$A$2:$A$4"
    trendline = xlpp.ChartTrendline()
    trendline.type = xlpp.ChartTrendlineType.LINEAR
    series.trendlines = [trendline]
    chart.add_series(series)
    assert chart.series[0].title == "2026"
    assert chart.series[0].trendlines[0].type == xlpp.ChartTrendlineType.LINEAR

    axis = xlpp.ChartAxis()
    axis.id = 101
    axis.kind = xlpp.ChartAxisKind.VALUE
    axis.scaling.minimum = 0
    chart.set_axes([axis])
    chart.set_stable_id("chart-1")
    chart.set_source_chart_part("/xl/charts/chart1.xml")
    chart.set_imported(True)
    assert chart.stable_id == "chart-1"
    assert chart.imported is True
    chart.add_plot(xlpp.ChartType.LINE, 0, 1)
    assert chart.plots[0].series_count == 1


def test_python_exposes_drawing_anchor_metadata():
    marker = xlpp.DrawingMarker()
    marker.row = 4
    marker.column = 3
    marker.row_offset_emu = 120
    anchor = xlpp.DrawingAnchorInfo()
    anchor.type = xlpp.DrawingAnchorType.TWO_CELL
    anchor.from_marker = marker
    anchor.width_emu = 9525
    image = xlpp.Image("C4", b"png", "png")
    image.anchor_info = anchor
    image.stable_id = "image-1"
    image.imported = True
    assert image.anchor_info.from_marker.row == 4
    assert image.stable_id == "image-1"
    assert image.imported


def test_python_exposes_cell_rich_text_and_row_cells():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Rich")
    cell = ws["A1"]
    rich = xlpp.RichText()
    run = xlpp.RichTextRun("Hello")
    run.bold = True
    rich.add_run(run)
    cell.set_rich_text(rich)
    assert cell.has_rich_text
    assert cell.rich_text_value.plain_text == "Hello"
    assert len(ws.row(1).cells) >= 1


def test_python_exposes_style_hash_and_equality():
    color_a = xlpp.Color("FF112233")
    color_b = xlpp.Color("FF112233")
    assert color_a == color_b
    assert color_a.hash() == color_b.hash()

    style = xlpp.Style()
    clone = xlpp.Style()
    assert style == clone
    assert style.hash() == clone.hash()


def test_python_exposes_reference_and_datetime_helpers():
    assert xlpp.CellReference.parse("$C$7") == xlpp.CellReference(7, 3)
    assert xlpp.is_valid_cell_coordinate(7, 3)
    assert xlpp.make_cell_key(7, 3) == (7 << 20) | 3
    assert xlpp.is_valid_worksheet_name("Input_Data")
    assert xlpp.worksheet_names_equal("Sheet", "sheet")
    assert xlpp.MAX_EXCEL_ROWS == 1048576
    assert xlpp.legacy_protection_password_hash("secret")
    renamed = xlpp.rename_worksheet_references("=Sheet!A1", "Sheet", "Input")
    assert renamed.changed and "Input!A1" in renamed.value
    invalidated = xlpp.invalidate_worksheet_references("=Sheet!A1", "Sheet")
    assert invalidated.changed and "#REF!" in invalidated.value
    value = xlpp.DateTime(2026, 8, 9, 10, 11, 12.5)
    assert value.to_iso8601_date() == "2026-08-09"
    assert value == xlpp.DateTime(2026, 8, 9, 10, 11, 12.5)
    assert value != xlpp.DateTime(2026, 8, 9)
    value.hour = 12
    assert value.hour == 12


def test_python_exposes_cell_position_and_date_overloads():
    cell = xlpp.Cell()
    cell.set_date(2026, 8, 9)
    assert cell.date().year == 2026


def test_python_exposes_table_column_constructors():
    column = xlpp.TableColumn(3, "Amount")
    assert column.id == 3
    assert column.name == "Amount"
    assert xlpp.DateTime().year == 1900
    assert xlpp.ChartSeries().title == ""
    assert xlpp.DefinedName().value == ""
    assert xlpp.ConditionalRule().type == xlpp.ConditionalRuleType.FORMULA


def test_python_exposes_streaming_default_constructors():
    cell = xlpp.StreamingCell()
    assert cell.address == ""
    assert xlpp.StreamingWorksheetReader()
    writer = xlpp.StreamingWorksheetWriter()
    assert writer.row_count == 0


def test_python_exposes_native_report_and_inspection_constructors():
    calculation = xlpp.CalculationReport()
    calculation.formula_cells_visited = 2
    assert calculation.formula_cells_visited == 2
    translation = xlpp.ReferenceTranslationResult()
    translation.value = "=A1"
    assert translation.value == "=A1"
    structural = xlpp.StructuralEditReport()
    structural.worksheets_visited = 1
    assert structural.worksheets_visited == 1
    assert xlpp.WorksheetRenameReport().worksheets_visited == 0
    assert xlpp.ChartCacheSyncReport().charts_visited == 0
    validation = xlpp.WorkbookValidationReport()
    validation.issues = [xlpp.WorkbookValidationIssue()]
    assert len(validation.issues) == 1
    encryption = xlpp.OfficeEncryptionInfo()
    encryption.encrypted = True
    assert encryption.encrypted is True
    dependency = xlpp.FormulaDependency()
    dependency.symbol = "A1"
    assert dependency.symbol == "A1"
    assert xlpp.FormulaDependencyReport().formula_cells == []
    assert xlpp.FormulaDependencyGraph().edges == []
    extents = xlpp.WorksheetExtents()
    extents.max_row = 10
    assert extents.max_row == 10

    assert xlpp.AutoFilter().enabled is False
    assert xlpp.ConditionalFormattingCollection().empty()
    assert xlpp.DataValidationCollection().empty()
    assert xlpp.ExternalDataInspection().has_external_workbooks is False
    assert xlpp.DataModelInspection().present is False


def test_python_exposes_error_codes_and_calculation_mode():
    assert xlpp.ErrorCode.INVALID_ARGUMENT != xlpp.ErrorCode.UNKNOWN
    properties = xlpp.CalcProperties()
    properties.calculation_mode = xlpp.CalculationMode.MANUAL
    assert properties.calculation_mode == xlpp.CalculationMode.MANUAL
    assert properties.calc_mode == "manual"


def test_scoped_defined_name_parity():
    wb = xlpp.Workbook()
    wb.add_worksheet("S1")
    wb.add_worksheet("S2")
    n1 = xlpp.DefinedName("Rate", "S1!$A$1")
    n1.local_sheet_id = 0
    wb.add_defined_name(n1)
    n2 = xlpp.DefinedName("Rate", "S2!$A$1")
    n2.local_sheet_id = 1
    wb.add_defined_name(n2)
    assert wb.defined_name("Rate", 0).local_sheet_id == 0
    assert wb.defined_name("Rate", 1).local_sheet_id == 1
    assert wb.defined_name("Rate", 0).value == "S1!$A$1"
    assert wb.defined_name("Rate", 1).value == "S2!$A$1"


def test_stable_child_handles_survive_collection_growth():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Stable")

    style = wb.add_named_style(xlpp.NamedStyle("PinnedStyle", xlpp.Style()))
    for i in range(192):
        wb.add_named_style(xlpp.NamedStyle(f"Style{i}", xlpp.Style()))
    style.name = "PinnedStyleUpdated"
    assert style.name == "PinnedStyleUpdated"

    defined = wb.add_defined_name(xlpp.DefinedName("PinnedName", "Stable!$A$1"))
    for i in range(192):
        wb.add_defined_name(xlpp.DefinedName(f"Name{i}", "Stable!$A$1"))
    defined.value = "Stable!$B$2"
    assert defined.value == "Stable!$B$2"

    table = ws.add_table("PinnedTable", "A1:B2")
    column = table.add_column("PinnedColumn")
    for i in range(128):
        table.add_column(f"Column{i}")
    column.name = "PinnedColumnUpdated"
    assert column.name == "PinnedColumnUpdated"

    for i in range(128):
        ws.add_table(f"Table{i}", "A1:B2")
    table.display_name = "PinnedTableUpdated"
    assert table.display_name == "PinnedTableUpdated"

    chart = xlpp.Chart(xlpp.ChartType.LINE)
    series = chart.add_series(xlpp.ChartSeries("PinnedSeries"))
    for i in range(128):
        chart.add_series(xlpp.ChartSeries(f"Series{i}"))
    series.title = "PinnedSeriesUpdated"
    assert series.title == "PinnedSeriesUpdated"
