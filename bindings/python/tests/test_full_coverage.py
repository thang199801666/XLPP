import xlpp


def test_cell_error_and_date_api(tmp_path):
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("S")
    ws["A1"].set_error(xlpp.CellError.DIVISION_BY_ZERO)
    assert ws["A1"].is_error()
    assert ws["A1"].value == "#DIV/0!"
    assert xlpp.cell_error_to_string(xlpp.CellError.NAME) == "#NAME?"
    ws["B1"].set_date(xlpp.DateTime(2024, 3, 15))
    assert ws["B1"].is_date()
    d = ws["B1"].date()
    assert d.year == 2024 and d.month == 3 and d.day == 15
    ws["C1"].set_formula("=A1*2")
    ws["D1"].set_array_formula("SUM(A1:A10)", "D1")
    ws["E1"].set_shared_formula("=A1+B1", 0, "E1")
    ws["F1"].set_dynamic_array_formula("_xlfn.SORT(A1:A10)", "F1")
    assert ws["F1"].formula.startswith("_xlfn.SORT")
    assert ws["F1"].formula_metadata.type == xlpp.FormulaType.DYNAMIC_ARRAY
    assert ws["G1"].has_value() is False


def test_cell_hyperlink_comment_style(tmp_path):
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("S")
    c = ws["A1"]
    hl = xlpp.Hyperlink("https://example.com")
    hl.display = "Example"
    hl.tooltip = "Tip"
    hl.external = True
    c.set_hyperlink(hl)
    assert c.has_hyperlink()
    assert c.hyperlink().target == "https://example.com"
    c.clear_hyperlink()
    assert not c.has_hyperlink()
    c.set_comment(xlpp.Comment("Note", "Author"))
    assert c.has_comment()
    assert c.comment().text == "Note"
    c.clear_comment()
    assert not c.has_comment()
    c.set_number_format("0.00")
    assert c.number_format == "0.00"
    s = c.style()
    s.num_fmt_id = 4
    assert s.num_fmt_id == 4
    s.locked = False
    assert not s.locked
    assert s.is_default() is False
    a = c.alignment()
    a.shrink_to_fit = True
    a.text_rotation = 45
    a.indent = 2
    assert a.shrink_to_fit is True
    assert a.text_rotation == 45
    assert a.indent == 2


def test_richtext(tmp_path):
    rt = xlpp.RichText()
    run = xlpp.RichTextRun("Hello")
    run.bold = True
    run.size = 14
    run.color = "FFFF0000"
    rt.add_run(run)
    rt.add_run(xlpp.RichTextRun("World"))
    assert rt.plain_text() == "HelloWorld"
    assert not rt.empty()
    assert len(rt.runs) == 2
    assert rt.runs[0].bold is True


def test_worksheet_dimensions_and_print(tmp_path):
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("S")
    ws["A1"].value = 1
    rd = ws.row_dimension(1)
    rd.height = 30
    assert rd.height == 30
    cd = ws.column_dimension(1)
    cd.width = 20
    assert cd.width == 20
    ws.set_print_area("A1:D20")
    assert ws.print_area == "A1:D20"
    ws.print_titles_rows = "$1:$1"
    assert ws.print_titles_rows == "$1:$1"
    ps = ws.page_setup()
    ps.paper_size = xlpp.PaperSize.A4
    ps.fit_to_page = True
    ps.black_and_white = True
    assert ps.paper_size == xlpp.PaperSize.A4
    assert ps.fit_to_page is True
    po = ws.print_options()
    po.horizontal_centered = True
    po.grid_lines = True
    assert po.horizontal_centered is True
    hf = ws.header_footer()
    hf.odd_header = "&CHead"
    hf.different_odd_even = True
    assert hf.different_odd_even is True


def test_protection_and_sheetview(tmp_path):
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("S")
    p = ws.protection()
    p.enabled = True
    p.insert_rows = False
    p.sort = False
    p.set_password("sheet-secret")
    assert p.has_password() is True
    p.clear_password()
    assert p.has_password() is False
    assert p.enabled is True
    assert p.insert_rows is False
    wbp = wb.protection()
    wbp.lock_structure = True
    wbp.set_password("workbook-secret")
    assert wbp.has_password() is True
    wbp.clear_password()
    assert wbp.has_password() is False
    assert wbp.lock_structure is True
    sv = ws.sheet_view()
    sv.zoom_scale = 120
    sv.show_grid_lines = False
    sv.right_to_left = True
    assert sv.zoom_scale == 120
    assert sv.show_grid_lines is False
    assert sv.right_to_left is True


def test_autofilter_datavalidation_cf(tmp_path):
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("S")
    af = ws.auto_filter()
    af.set_reference("A1:B2")
    assert af.enabled is True
    col = af.column(0)
    col.add_value("foo")
    col.and_mode = True
    col.include_blank = True
    assert col.and_mode is True
    assert 0 in af.columns
    af.sort_state().case_sensitive = True
    assert af.sort_state_value.case_sensitive is True
    dynamic = xlpp.DynamicFilter()
    dynamic.type = xlpp.DynamicFilterType.TODAY
    col.set_dynamic_filter(dynamic)
    assert col.dynamic_filter.type == xlpp.DynamicFilterType.TODAY
    col.clear_dynamic_filter()
    assert col.dynamic_filter is None
    dv = ws.data_validations.add(xlpp.DataValidationType.WHOLE, "A1:A10")
    dv.op = xlpp.DataValidationOperator.BETWEEN
    dv.formula1 = "1"
    dv.formula2 = "100"
    dv.error_style = xlpp.DataValidationErrorStyle.WARNING
    dv.show_drop_down = True
    assert dv.error_style == xlpp.DataValidationErrorStyle.WARNING
    assert dv.show_drop_down is True
    cf = ws.conditional_formatting.add_rule("B1:B3", xlpp.ConditionalRuleType.FORMULA, "B1>0")
    assert cf.type == xlpp.ConditionalRuleType.FORMULA
    assert len(ws.conditional_formatting.entries) == 1
    path = tmp_path / "af.xlsx"
    wb.save(str(path))
    wb2 = xlpp.Workbook()
    wb2.load(str(path))
    assert wb2["S"].auto_filter().enabled is True


def test_streaming_writer_reader(tmp_path):
    path = tmp_path / "stream.xlsx"
    writer = xlpp.StreamingWorkbookWriter(str(path))
    sheet = writer.add_worksheet("Big")
    for i in range(100):
        sheet.append([f"row{i}", i])
    assert sheet.row_count == 100
    writer.close()
    assert writer.closed is True
    options = xlpp.StreamingReaderOptions()
    options.max_file_bytes = 1024 * 1024
    options.max_entries = 100
    reader = xlpp.StreamingWorkbookReader(str(path), options)
    assert reader.worksheet_names() == ["Big"]
    ws = reader.worksheet("Big")
    rows = list(ws)
    assert len(rows) == 100
    assert rows[0][1][0].address == "A1"
    assert rows[0][1][0].value == "row0"


def test_workbook_bytes_and_customprops(tmp_path):
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("S")
    ws["A1"].value = 42
    data = wb.save_bytes()
    assert isinstance(data, bytes)
    wb2 = xlpp.Workbook()
    wb2.load_bytes(data)
    assert wb2["S"]["A1"].value == 42.0
    cp = xlpp.CustomProperty("myKey", "myValue")
    assert cp.name == "myKey"
    assert cp.type == "lpwstr"
    wb.custom_properties().add(cp)
    assert wb.custom_properties().items[0].value == "myValue"
    wb.calc_properties().full_calc_on_load = True
    assert wb.calc_properties().full_calc_on_load is True


def test_cell_range_and_row(tmp_path):
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("S")
    for r in range(1, 4):
        for c in range(1, 4):
            ws.cell(r, c).value = r * 10 + c
    rng = ws.range("A1:C3")
    assert rng.address() == "A1:C3"
    assert rng.row_count == 3
    assert rng.column_count == 3
    assert rng.min_row == 1
    assert rng.max_row == 3
    assert rng.cell(1, 1).value == 11.0
    assert len(rng.values()) == 9
    assert len(rng.cells()) == 9
    row = ws.row(1)
    assert row.number == 1
    assert row.values()[0] == 11.0


def test_defined_name_and_docprops(tmp_path):
    wb = xlpp.Workbook()
    wb.add_worksheet("S")
    dn = wb.add_defined_name(xlpp.DefinedName("MyRange", "S!$A$1:$A$10"))
    assert dn.name == "MyRange"
    dn.hidden = True
    dn.comment = "note"
    assert dn.hidden is True
    assert dn.comment == "note"
    assert wb.defined_name("MyRange") is not None
    props = wb.properties
    props.category = "Cat"
    props.last_modified_by = "User"
    assert props.category == "Cat"
    assert props.last_modified_by == "User"
    ns = wb.add_named_style(xlpp.NamedStyle("MyStyle"))
    ns.style().font().bold = True
    ns.style().font().size = 14
    assert wb.named_style("MyStyle") is not None
    assert len(wb.named_styles) == 1


def test_excel_serial_and_parse():
    d = xlpp.from_excel_serial(45366.0)
    assert d.year == 2024
    assert d.month == 3
    assert d.day == 15
    assert xlpp.to_excel_serial(d) == 45366.0
    assert xlpp.CellReference.column_name(27) == "AA"
    assert xlpp.CellReference.column_index("AA") == 27
    parsed = xlpp.parse_iso8601("2024-03-15T10:30:00")
    assert parsed is not None
    assert parsed.hour == 10
