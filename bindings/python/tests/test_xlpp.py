"""Comprehensive tests for the xlpp Python binding.

Runs against the compiled extension in bindings/python (see conftest.py).
Each test exercises the public Python API and round-trips through real
.xlsx files where relevant.
"""

from datetime import date, datetime

import pytest

import xlpp


# ---------------------------------------------------------------------------
# Module basics
# ---------------------------------------------------------------------------

def test_version():
    assert isinstance(xlpp.__version__, str)
    assert xlpp.__version__.count(".") == 2


def test_xlfn_helper():
    assert xlpp.xlfn("SORT") == "_xlfn.SORT"
    assert xlpp.xlfn("_xlfn.XLOOKUP") == "_xlfn.XLOOKUP"
    assert xlpp.xlfn("FILTER(A1:A5,\"x\")") == "_xlfn.FILTER(A1:A5,\"x\")"
    assert xlpp.xlfn("") == ""


def test_compression_enum():
    assert xlpp.CompressionLevel.STORE.value == 0
    assert xlpp.CompressionLevel.DEFAULT.value == 2
    assert xlpp.CompressionLevel.BEST.value == 3


def test_page_orientation_enum():
    assert xlpp.PageOrientation.DEFAULT.value == 0
    assert xlpp.PageOrientation.PORTRAIT.value == 1
    assert xlpp.PageOrientation.LANDSCAPE.value == 2


# ---------------------------------------------------------------------------
# CellReference
# ---------------------------------------------------------------------------

def test_cell_reference():
    ref = xlpp.CellReference(3, 2)
    assert ref.row == 3
    assert ref.column == 2
    assert ref.address() == "B3"
    parsed = xlpp.CellReference.parse("$AA$42")
    assert parsed.row == 42
    assert parsed.column == 27
    assert parsed.address() == "AA42"


# ---------------------------------------------------------------------------
# DateTime
# ---------------------------------------------------------------------------

def test_datetime_props():
    dt = xlpp.DateTime(2024, 1, 15, 13, 30, 45.25)
    assert dt.year == 2024
    assert dt.month == 1
    assert dt.day == 15
    assert dt.hour == 13
    assert dt.minute == 30
    assert dt.second_int == 45
    assert dt.millisecond == 250
    assert "2024-01-15T13:30:45.250" in str(dt)


# ---------------------------------------------------------------------------
# Workbook / worksheet lifecycle
# ---------------------------------------------------------------------------

def test_workbook_lifecycle():
    wb = xlpp.Workbook()
    assert len(wb) == 0
    ws = wb.add_worksheet("Data")
    assert len(wb) == 1
    assert wb.sheet_names == ["Data"]
    assert wb.worksheet("Data") is ws
    assert wb[0] is ws
    assert wb["Data"] is ws
    assert wb.active is ws
    assert wb.index(ws) == 0
    assert "Data" in repr(ws)
    assert "sheets=1" in repr(wb)


def test_workbook_iteration():
    wb = xlpp.Workbook()
    wb.add_worksheet("A")
    wb.add_worksheet("B")
    assert [w.name for w in wb] == ["A", "B"]


def test_add_duplicate_worksheet_raises():
    wb = xlpp.Workbook()
    wb.add_worksheet("Dup")
    with pytest.raises(ValueError):
        wb.add_worksheet("Dup")


def test_remove_worksheet():
    wb = xlpp.Workbook()
    wb.add_worksheet("First")
    wb.add_worksheet("Second")
    assert wb.remove_worksheet("First") is True
    assert wb.sheet_names == ["Second"]
    assert wb.remove_worksheet("Nope") is False


def test_worksheet_not_found_raises():
    wb = xlpp.Workbook()
    with pytest.raises(ValueError):
        _ = wb["Missing"]


def test_copy_worksheet():
    wb = xlpp.Workbook()
    src = wb.add_worksheet("Source")
    src.cell("A1").value = "original"
    copy = wb.copy_worksheet(src, "Copy")
    assert copy.name == "Copy"
    assert copy.cell("A1").value == "original"
    copy.cell("A1").value = "changed"
    assert src.cell("A1").value == "original"


def test_workbook_clear():
    wb = xlpp.Workbook()
    wb.add_worksheet("A")
    wb.add_worksheet("B")
    wb.clear()
    assert len(wb) == 0


# ---------------------------------------------------------------------------
# Cell values
# ---------------------------------------------------------------------------

def test_cell_value_types():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Values")
    ws["A1"].value = "text"
    ws["A2"].value = 42
    ws["A3"].value = 3.5
    ws["A4"].value = True
    ws["A5"].value = None
    ws["A6"].value = date(2024, 1, 15)
    ws["A7"].value = datetime(2024, 1, 15, 13, 30, 45)

    assert ws["A1"].value == "text"
    assert ws["A2"].value == 42.0
    assert ws["A3"].value == 3.5
    assert ws["A4"].value is True
    assert ws["A5"].value is None
    assert ws["A6"].value == datetime(2024, 1, 15)
    assert ws["A7"].value == datetime(2024, 1, 15, 13, 30, 45)

    assert ws["A1"].is_string()
    assert ws["A2"].is_numeric()
    assert ws["A3"].is_numeric()
    assert ws["A4"].is_bool()
    assert ws["A5"].empty()
    assert ws["A6"].is_date()
    assert ws["A7"].is_date()

    assert ws["A1"].value_type() == "string"
    assert ws["A2"].value_type() == "numeric"
    assert ws["A4"].value_type() == "bool"
    assert ws["A5"].value_type() == "empty"
    assert ws["A6"].value_type() == "date"


def test_cell_coordinates_and_address():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Coords")
    c = ws.cell(5, 3)
    assert c.address == "C5"
    assert c.row == 5
    assert c.column == 3
    assert ws.cell((2, 1)).address == "A2"
    with pytest.raises(ValueError):
        ws.cell(0, 1)


def test_cell_clear_and_formula():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("F")
    c = ws["A1"]
    c.value = 42
    c.set_formula("=B1*2")
    assert c.has_formula()
    assert c.formula == "=B1*2"
    c.clear()
    assert c.empty()
    assert not c.has_formula()


def test_dynamic_array_formula():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("DA")
    ws["A1"].set_dynamic_array_formula("_xlfn.SORT(A2:A5)", "A1")
    assert ws["A1"].has_formula()
    assert ws["A1"].formula == "_xlfn.SORT(A2:A5)"


def test_cell_offset():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Off")
    ref = ws["C5"].offset(2, 1)
    assert ref.address() == "D7"


def test_try_cell():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("TC")
    ws["A1"].value = "x"
    assert ws.try_cell("A1") is not None
    assert ws.try_cell("Z99") is None


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

def test_font_styling():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Style")
    f = ws["A1"].font()
    f.name = "Arial"
    f.size = 14.0
    f.bold = True
    f.italic = True
    f.underline = True
    f.strike = True
    f.color().set_argb("FFFF0000")
    assert f.name == "Arial"
    assert f.size == 14.0
    assert f.bold is True
    assert f.italic is True
    assert f.underline is True
    assert f.strike is True
    assert f.color().argb == "FFFF0000"


def test_fill_and_border():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Style2")
    fill = ws["A1"].fill()
    fill.pattern_type = "solid"
    fill.foreground().set_argb("FFFFFF00")
    assert fill.pattern_type == "solid"
    assert fill.foreground().argb == "FFFFFF00"

    b = ws["A1"].border()
    b.left().style = "thin"
    b.left().color().set_argb("FF000000")
    assert b.left().style == "thin"
    assert b.left().color().argb == "FF000000"


def test_alignment_and_number_format():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Style3")
    a = ws["A1"].alignment()
    a.horizontal = "center"
    a.vertical = "center"
    a.wrap_text = True
    assert a.horizontal == "center"
    assert a.vertical == "center"
    assert a.wrap_text is True

    ws["A1"].set_number_format("0.00")
    assert ws["A1"].number_format == "0.00"


def test_style_proxy():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Style4")
    st = ws["A1"].style()
    st.number_format = "#,##0.00"
    assert st.number_format == "#,##0.00"
    assert st.font().bold is False


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------

def test_merge_cells():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Merge")
    ws.merge_cells("A1:B2")
    assert ws.is_merged("A1")
    assert ws.is_merged("B2")
    assert ws.merged_ranges == ["A1:B2"]
    ws.unmerge_cells("A1:B2")
    assert not ws.is_merged("A1")
    assert ws.merged_ranges == []


def test_freeze_panes():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Freeze")
    ws.freeze_panes("B2")
    assert ws.frozen_pane == "B2"
    ws.clear_freeze_panes()
    assert ws.frozen_pane is None


def test_dimensions():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Dim")
    assert ws.empty
    assert ws.max_row == 1
    assert ws.max_column == 1
    ws["C5"].value = 1
    ws["A2"].value = "a"
    assert ws.max_row == 5
    assert ws.max_column == 3
    assert ws.row_count == 5
    assert ws.col_count == 3
    assert ws.dimensions() == "A2:C5"
    assert not ws.empty


def test_insert_delete_rows_columns():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Edit")
    ws["A1"].value = "one"
    ws["B2"].value = 10.0
    ws.insert_rows(2)
    assert ws["A1"].value == "one"
    assert ws["B3"].value == 10.0
    ws.delete_rows(2)
    assert ws["B2"].value == 10.0
    ws.insert_columns(2, 2)
    assert ws["D2"].value == 10.0
    ws.delete_columns(2, 2)
    assert ws["B2"].value == 10.0


def test_print_area():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Print")
    ws.set_print_area("A1:D10")
    assert ws.print_area == "A1:D10"


# ---------------------------------------------------------------------------
# Bulk / records / iteration
# ---------------------------------------------------------------------------

def test_append():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Append")
    ws.append(["Name", "Value"])
    ws.append(["Alice", 100])
    ws.append(["Bob", True, date(2024, 5, 1)])
    assert ws["A1"].value == "Name"
    assert ws["B2"].value == 100.0
    assert ws["B3"].value is True
    assert ws["C3"].value == datetime(2024, 5, 1)
    assert ws.max_row == 3
    assert ws.max_column == 3


def test_append_strings():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("AppendS")
    ws.append_strings(["a", "b", "c"])
    ws.append_strings([1, 2, 3])
    assert ws["A1"].value == "a"
    assert ws["C1"].value == "c"
    assert ws["A2"].value == 1.0


def test_from_records():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Records")
    ws.from_records([["a", 1], ["b", 2]], columns=["key", "val"], header=True)
    assert ws["A1"].value == "key"
    assert ws["B1"].value == "val"
    assert ws["A2"].value == "a"
    assert ws["B2"].value == 1.0
    assert ws["A3"].value == "b"


def test_to_records():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Rec2")
    ws["A1"].value = "key"
    ws["B1"].value = "val"
    ws["A2"].value = "x"
    ws["B2"].value = 2.0
    records = ws.to_records(include_header=True)
    assert records[0] == ["key", "val"]
    assert records[1] == ["x", 2.0]
    no_header = ws.to_records(include_header=False)
    assert no_header[0] == ["x", 2.0]


def test_worksheet_row_iteration():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Rows")
    ws["A1"].value = "a"
    ws["B1"].value = "b"
    ws["A2"].value = "c"
    rows = list(ws)
    assert len(rows) == 2
    assert rows[0].number == 1
    assert rows[1].number == 2
    cells = list(rows[0])
    assert len(cells) == 2
    assert cells[0].value == "a"
    assert cells[1].value == "b"


# ---------------------------------------------------------------------------
# Hyperlinks / comments / properties
# ---------------------------------------------------------------------------

def test_hyperlink():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Links")
    link = xlpp.Hyperlink("https://example.com")
    link.display = "Example"
    link.tooltip = "Open"
    ws["A1"].set_hyperlink(link)
    assert ws["A1"].has_hyperlink()
    assert ws["A1"].hyperlink().target == "https://example.com"
    assert ws["A1"].hyperlink().display == "Example"


def test_comment():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Comments")
    c = xlpp.Comment("Review this", "Alice")
    ws["A1"].set_comment(c)
    assert ws["A1"].has_comment()
    assert ws["A1"].comment().text == "Review this"
    assert ws["A1"].comment().author == "Alice"


def test_document_properties():
    wb = xlpp.Workbook()
    wb.add_worksheet("S")
    props = wb.properties
    props.title = "My Report"
    props.creator = "Tester"
    props.subject = "Unit"
    props.description = "A description"
    props.keywords = "a,b"
    assert props.title == "My Report"
    assert props.creator == "Tester"


def test_defined_name():
    wb = xlpp.Workbook()
    wb.add_worksheet("S")
    name = wb.add_defined_name(xlpp.DefinedName("MyRange", "'S'!$A$1:$B$2"))
    assert name.name == "MyRange"
    assert name.value == "'S'!$A$1:$B$2"


def test_date1904_flag():
    wb = xlpp.Workbook()
    wb.add_worksheet("S")
    assert wb.date_1904 is False
    wb.date_1904 = True
    assert wb.date_1904 is True


# ---------------------------------------------------------------------------
# Tables / filters / page setup
# ---------------------------------------------------------------------------

def test_table():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Table")
    ws.append(["Col1", "Col2"])
    ws.append([1, 2])
    t = ws.add_table("MyTable", "A1:B2")
    assert t.name == "MyTable"
    assert t.reference == "A1:B2"
    t.add_column("Col1")
    assert t.display_name == "MyTable"


def test_auto_filter():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Filter")
    af = ws.auto_filter()
    af.reference = "A1:D20"
    assert af.reference == "A1:D20"
    col = af.column(1)
    col.add_value("Open")
    assert col.values == ["Open"]
    ss = af.sort_state()
    ss.set_reference("A2:D20")
    ss.add_condition("C2:C20", True)


def test_page_setup_and_margins():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Page")
    ps = ws.page_setup()
    ps.orientation = xlpp.PageOrientation.LANDSCAPE
    ps.scale = 85
    assert ps.orientation == xlpp.PageOrientation.LANDSCAPE
    assert ps.scale == 85
    m = ws.page_margins()
    m.left = 0.25
    m.right = 0.5
    m.top = 0.75
    m.bottom = 1.0
    assert m.left == 0.25
    assert m.bottom == 1.0


# ---------------------------------------------------------------------------
# Round-trips through real files
# ---------------------------------------------------------------------------

def test_roundtrip_values(tmp_path):
    path = tmp_path / "values.xlsx"
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Data")
    ws["A1"].value = "hello"
    ws["B1"].value = 42.5
    ws["C1"].value = True
    ws["D1"].value = date(2024, 1, 15)
    ws["E1"].set_formula("B1*2")
    wb.properties.title = "Roundtrip"
    wb.save(str(path))

    wb2 = xlpp.Workbook()
    wb2.load(str(path))
    ws2 = wb2["Data"]
    assert ws2["A1"].value == "hello"
    assert ws2["B1"].value == 42.5
    assert ws2["C1"].value is True
    assert ws2["D1"].value == datetime(2024, 1, 15)
    assert ws2["E1"].formula == "B1*2"
    assert wb2.properties.title == "Roundtrip"


def test_roundtrip_styles(tmp_path):
    path = tmp_path / "styles.xlsx"
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Styled")
    ws["A1"].value = "Bold text"
    f = ws["A1"].font()
    f.bold = True
    f.size = 16.0
    f.color().set_argb("FF0000FF")
    ws["A1"].set_number_format("0.00")
    ws["B2"].value = 0
    ws["B2"].fill().pattern_type = "solid"
    ws["B2"].fill().foreground().set_argb("FFFFFF00")
    ws["B2"].border().top().style = "medium"
    wb.save(str(path))

    wb2 = xlpp.Workbook()
    wb2.load(str(path))
    cell = wb2["Styled"]["A1"]
    assert cell.value == "Bold text"
    assert cell.font().bold is True
    assert cell.font().size == 16.0
    assert cell.font().color().argb == "FF0000FF"
    assert cell.number_format == "0.00"
    assert wb2["Styled"]["B2"].fill().pattern_type == "solid"
    assert wb2["Styled"]["B2"].border().top().style == "medium"


def test_roundtrip_layout(tmp_path):
    path = tmp_path / "layout.xlsx"
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Layout")
    ws["A1"].value = "Title"
    ws.merge_cells("A1:C1")
    ws.freeze_panes("A2")
    wb.save(str(path))

    wb2 = xlpp.Workbook()
    wb2.load(str(path))
    ws2 = wb2["Layout"]
    assert ws2.merged_ranges == ["A1:C1"]
    assert ws2.frozen_pane == "A2"
    assert ws2["A1"].value == "Title"


def test_roundtrip_hyperlink_comment(tmp_path):
    path = tmp_path / "links.xlsx"
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("Links")
    ws["A1"].value = "Click"
    ws["A1"].set_hyperlink(xlpp.Hyperlink("https://example.com"))
    ws["A1"].set_comment(xlpp.Comment("note", "Alice"))
    wb.save(str(path))

    wb2 = xlpp.Workbook()
    wb2.load(str(path))
    ws2 = wb2["Links"]
    assert ws2["A1"].has_hyperlink()
    assert ws2["A1"].hyperlink().target == "https://example.com"
    assert ws2["A1"].has_comment()
    assert ws2["A1"].comment().text == "note"
    assert ws2["A1"].comment().author == "Alice"


def test_roundtrip_multiple_sheets(tmp_path):
    path = tmp_path / "multi.xlsx"
    wb = xlpp.Workbook()
    wb.add_worksheet("One").cell("A1").value = 1
    wb.add_worksheet("Two").cell("A1").value = 2
    wb.add_worksheet("Three").cell("A1").value = 3
    wb.save(str(path))

    wb2 = xlpp.Workbook()
    wb2.load(str(path))
    assert wb2.sheet_names == ["One", "Two", "Three"]
    assert wb2["One"]["A1"].value == 1.0
    assert wb2["Three"]["A1"].value == 3.0


def test_roundtrip_date1904(tmp_path):
    path = tmp_path / "date1904.xlsx"
    wb = xlpp.Workbook()
    wb.date_1904 = True
    ws = wb.add_worksheet("Dates")
    ws["A1"].value = date(2024, 1, 15)
    wb.save(str(path))

    wb2 = xlpp.Workbook()
    wb2.load(str(path))
    assert wb2.date_1904 is True
    assert wb2["Dates"]["A1"].value == datetime(2024, 1, 15)


def test_save_options(tmp_path):
    path = tmp_path / "opts.xlsx"
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("S")
    for i in range(1, 31):
        ws.append([f"row{i}", i])
    opts = xlpp.SaveOptions()
    opts.compression_level = xlpp.CompressionLevel.BEST
    opts.parallel_workers = 4
    wb.save(str(path), opts)

    wb2 = xlpp.Workbook()
    wb2.load(str(path))
    assert wb2["S"]["A30"].value == "row30"
    assert wb2["S"]["B30"].value == 30.0


def test_load_options_lenient(tmp_path):
    path = tmp_path / "lenient.xlsx"
    wb = xlpp.Workbook()
    wb.add_worksheet("S").cell("A1").value = "x"
    wb.save(str(path))

    opts = xlpp.LoadOptions()
    opts.lenient = True
    wb2 = xlpp.Workbook()
    wb2.load(str(path), opts)
    assert wb2["S"]["A1"].value == "x"


def test_missing_file_raises(tmp_path):
    wb = xlpp.Workbook()
    with pytest.raises(RuntimeError):
        wb.load(str(tmp_path / "does_not_exist.xlsx"))


# ---------------------------------------------------------------------------
# NumPy integration (skipped when numpy is unavailable)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("transpose", [False, True])
def test_write_array(transpose):
    np = pytest.importorskip("numpy")
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("NP")
    arr = np.array([[1.0, 2.0], [3.0, 4.0]])
    ws.write_array(arr, row=1, col=1, transpose=transpose)
    if transpose:
        assert ws["A1"].value == 1.0
        assert ws["B1"].value == 3.0
        assert ws["A2"].value == 2.0
    else:
        assert ws["A1"].value == 1.0
        assert ws["B1"].value == 2.0
        assert ws["A2"].value == 3.0
        assert ws["B2"].value == 4.0


def test_to_array():
    np = pytest.importorskip("numpy")
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("NP2")
    ws["A1"].value = 1.0
    ws["B1"].value = 2.0
    ws["A2"].value = 3.0
    ws["B2"].value = 4.0
    arr = ws.to_array()
    assert arr.shape == (2, 2)
    assert arr[0][0] == 1.0
    assert arr[1][1] == 4.0


def test_numpy_roundtrip(tmp_path):
    np = pytest.importorskip("numpy")
    path = tmp_path / "numpy.xlsx"
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("NP3")
    data = np.arange(12.0).reshape(3, 4)
    ws.write_array(data, row=1, col=1)
    wb.save(str(path))

    wb2 = xlpp.Workbook()
    wb2.load(str(path))
    got = wb2["NP3"].to_array()
    assert np.array_equal(got, data)


# ---------------------------------------------------------------------------
# Exception mapping
# ---------------------------------------------------------------------------

def test_bad_cell_key_raises():
    wb = xlpp.Workbook()
    ws = wb.add_worksheet("S")
    with pytest.raises(ValueError):
        ws.cell("not-a-valid-address!")


def test_index_error_for_bad_index():
    wb = xlpp.Workbook()
    with pytest.raises(IndexError):
        _ = wb[5]
