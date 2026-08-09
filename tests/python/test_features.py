from datetime import date, datetime

import pytest

import xlpp


def test_cell_types_formula_and_round_trip(tmp_path):
    path = tmp_path / "features.xlsx"
    book = xlpp.Workbook()
    sheet = book.add_worksheet("Data")

    sheet["A1"].value = "Name"
    sheet["B1"].value = 42
    sheet["C1"].value = 3.5
    sheet["D1"].value = True
    sheet["E1"].value = date(2024, 1, 15)
    sheet["F1"].value = datetime(2024, 1, 15, 12, 30)
    sheet["G1"].value = None
    sheet["H1"].set_formula("=B1+C1")
    sheet["I1"].set_dynamic_array_formula("=SORT(A1:A3)", "I1:I3")

    assert sheet["A1"].is_string()
    assert sheet["B1"].is_numeric()
    assert sheet.try_cell_rc(1, 2).address == "B1"
    assert sheet.get_cell((1, 2)).address == "B1"
    assert sheet.rows_list()[0].number == 1
    assert sheet["D1"].is_bool()
    assert sheet["E1"].is_date()
    assert sheet["H1"].has_formula()

    book.save(str(path))
    loaded = xlpp.Workbook()
    loaded.load(str(path))
    result = loaded["Data"]
    assert result["A1"].value == "Name"
    assert result["B1"].value == 42.0
    assert result["D1"].value is True
    assert result["H1"].has_formula()


def test_styles_layout_metadata_and_relationships(tmp_path):
    path = tmp_path / "features.xlsx"
    book = xlpp.Workbook()
    sheet = book.add_worksheet("Report")
    cell = sheet["A1"]
    cell.value = "Report"
    cell.font().bold = True
    cell.font().size = 16
    cell.font().color().set_argb("FFFF0000")
    cell.fill().pattern_type = "solid"
    cell.fill().foreground().set_argb("FFCCFFCC")
    cell.border().bottom().style = "thin"
    cell.alignment().horizontal = "center"
    cell.set_number_format("#,##0.00")
    cell.set_hyperlink(xlpp.Hyperlink("https://example.com"))
    cell.set_comment(xlpp.Comment("Check this value", "QA"))

    sheet.merge_cells("A1:C1")
    sheet.freeze_panes("A2")
    sheet.set_print_area("A1:C10")
    sheet.print_titles_rows = "1:1"
    sheet.print_titles_cols = "A:A"
    sheet.auto_filter().reference = "A1:C10"
    sheet.page_setup().orientation = xlpp.PageOrientation.LANDSCAPE
    sheet.page_margins().left = 0.5
    sheet.protection().enabled = True
    table = sheet.add_table("A2:C3", "ReportTable")
    table.add_column("Name")

    book.properties.title = "Feature test"
    book.properties.creator = "XLPP tests"
    book.add_defined_name("ReportRange", "Report!$A$1:$C$10")
    assert sheet.is_merged("A1")
    assert cell.has_hyperlink()
    assert cell.has_comment()

    book.save(str(path))
    loaded = xlpp.Workbook()
    loaded.load(str(path))
    assert loaded.properties.title == "Feature test"
    assert loaded["Report"].is_merged("A1")


def test_bulk_numpy_and_records(tmp_path):
    np = pytest.importorskip("numpy")
    path = tmp_path / "bulk.xlsx"
    book = xlpp.Workbook()
    sheet = book.add_worksheet("Bulk")
    values = np.arange(100, dtype=np.float64).reshape(20, 5)
    sheet.write_array(values)
    assert sheet.to_array(1, 1, 20, 5).shape == (20, 5)
    assert sheet.to_array(1, 1, 20, 5)[19, 4] == 99

    sheet.from_records([["Alice", 30], ["Bob", 25]], ["Name", "Age"])
    records = sheet.to_records()
    assert records[0] == ["Name", "Age"]
    assert records[1][0] == "Alice"
    book.save(str(path))


def test_workbook_operations(tmp_path):
    path = tmp_path / "operations.xlsx"
    book = xlpp.Workbook()
    first = book.add_worksheet("First")
    first.append(["a", 1])
    book.add_worksheet("Second")
    assert len(book) == 2
    copied = book.copy_worksheet(first, "Copy")
    assert copied.name == "Copy"
    book.remove_worksheet("Second")
    assert book.sheet_names == ["First", "Copy"]
    assert xlpp.xlfn("SORT") == "_xlfn.SORT"
    book.save(str(path))


def test_bytes_round_trip():
    book = xlpp.Workbook()
    book.add_worksheet("Bytes")["A1"].value = "in-memory"
    payload = book.save_bytes()
    assert isinstance(payload, bytes)
    loaded = xlpp.Workbook()
    loaded.load_bytes(payload)
    assert loaded["Bytes"]["A1"].value == "in-memory"
    assert loaded.diagnostics.warning_count == 0
    assert loaded.diagnostics.error_count == 0


def test_extended_python_api():
    for name in ("DateTime", "to_excel_serial", "FormulaType", "StreamingWorkbookReader"):
        if not hasattr(xlpp, name):
            pytest.skip("Python extension must be rebuilt for the extended binding API")
    dt = xlpp.DateTime(2024, 1, 15, 12, 30, 0)
    assert xlpp.CellReference.column_name(27) == "AA"
    assert xlpp.CellReference.column_index("AA") == 27
    serial = xlpp.to_excel_serial(dt)
    restored = xlpp.from_excel_serial(serial)
    assert restored.year == 2024
    assert xlpp.parse_iso8601("2024-01-15T12:30:00").month == 1
    assert xlpp.cell_error_to_string(xlpp.CellError.DIVISION_BY_ZERO) == "#DIV/0!"

    formula = xlpp.FormulaMetadata()
    formula.type = xlpp.FormulaType.DYNAMIC_ARRAY
    formula.reference = "A1:A3"
    formula.always_calculate_array = True
    assert formula.reference == "A1:A3"

    book = xlpp.Workbook()
    book.protection.lock_structure = True
    book.calc_properties.full_calc_on_load = True
    book.custom_properties.add(xlpp.CustomProperty("Build", "test"))
    sheet = book.add_worksheet("Extended")
    sheet.sheet_view.zoom_scale = 90
    sheet.protection.enabled = True
    sheet.protection.select_locked_cells = False
    sheet["A1"].value = 1
    sheet.row_dimension(1).height = 24
    sheet.column_dimension_by_name("A").width = 18
    sheet.page_setup().fit_to_page = True
    sheet.page_margins().header = 0.4
    assert sheet.extents() == (1, 1, 1, 1)
    assert sheet.range_rc(1, 1, 1, 1).address == "A1"
    assert list(sheet.range_rc(1, 1, 1, 1))[0].address == "A1"
    assert xlpp.chart_type_name(xlpp.ChartType.BAR) == "barChart"
    sheet["A1"].set_date(dt)
    assert sheet["A1"].date.year == 2024
    assert sheet["A1"].numeric_value_or(-1) == xlpp.to_excel_serial(dt)
    sheet["C1"].set_error(xlpp.CellError.VALUE)
    assert sheet["C1"].error == xlpp.CellError.VALUE
    sheet["D1"].set_date_parts(2024, 2, 1)
    assert sheet["D1"].date.month == 2
    link = xlpp.Hyperlink("Sheet2!A1")
    link.external = False
    sheet["B1"].set_hyperlink(link)
    assert sheet["B1"].hyperlink().external is False

    rule = xlpp.ConditionalRule.data_bar("FF638EC6")
    sheet.conditional_formatting().add_rule("A1:A3", rule)
    icon_rule = xlpp.ConditionalRule.icon_set("3Arrows")
    icon_rule.icon_set_config().style = xlpp.IconSetStyle.THREE_ARROWS
    sheet.conditional_formatting().add_rule("B1:B3", icon_rule)
    validation = xlpp.DataValidation.list("A1:A3", '"A,B"')
    validation.allow_blank = True
    validation.show_error_message = True
    sheet.data_validations().add(validation)
    assert not sheet.data_validations().empty
    assert not sheet.conditional_formatting().empty
