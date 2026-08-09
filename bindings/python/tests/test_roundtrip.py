from datetime import date, datetime

import pytest

import xlpp


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


def test_write_read_file_with_library_features(tmp_path):
    path = tmp_path / "library_features.xlsx"
    workbook = xlpp.Workbook()
    worksheet = workbook.add_worksheet("Report")
    worksheet.append(["Product", "Quantity", "Price"])
    worksheet.append(["Keyboard", 2, 25.5])
    worksheet.append(["Mouse", 3, 12.0])
    worksheet["D1"].value = "Total"
    worksheet["D2"].set_formula("B2*C2")
    worksheet["D3"].set_formula("B3*C3")
    worksheet["A1"].font().bold = True
    worksheet["D1"].font().bold = True
    workbook.properties.title = "Sales report"
    workbook.save(str(path))
    loaded = xlpp.Workbook()
    loaded.load(str(path))
    report = loaded["Report"]
    assert report.to_records(include_header=True) == [
        ["Product", "Quantity", "Price", "Total"],
        ["Keyboard", 2.0, 25.5, None],
        ["Mouse", 3.0, 12.0, None],
    ]
    assert report["D2"].formula == "B2*C2"
    assert report["D3"].formula == "B3*C3"
    assert report["A1"].font().bold is True
    assert loaded.properties.title == "Sales report"


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
    wb.add_worksheet("Dates")["A1"].value = date(2024, 1, 15)
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
    with pytest.raises(RuntimeError):
        xlpp.Workbook().load(str(tmp_path / "does_not_exist.xlsx"))
