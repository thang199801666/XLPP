import xlpp


def main():
    assert xlpp.__version__ == "1.1.0"
    assert xlpp.CellReference.column_name(27) == "AA"
    assert xlpp.CellReference.column_index("AA") == 27
    assert hasattr(xlpp, "FormulaType")
    assert hasattr(xlpp, "CellError")
    assert hasattr(xlpp, "StreamingWorkbookReader")
    assert hasattr(xlpp, "DataValidation")
    assert hasattr(xlpp, "ConditionalRule")

    workbook = xlpp.Workbook()
    sheet = workbook.add_worksheet("Smoke")
    sheet["A1"].value = 42
    sheet["B1"].set_formula("=A1*2")
    sheet["A1"].font().bold = True
    sheet.merge_cells("A2:B2")
    sheet.data_validations().add(xlpp.DataValidation.list("A1:A3", '"Open,Closed"'))
    sheet.conditional_formatting().add_rule("A1:A3", xlpp.ConditionalRule.data_bar())

    payload = workbook.save_bytes()
    loaded = xlpp.Workbook()
    loaded.load_bytes(payload)
    assert loaded["Smoke"]["A1"].value == 42.0
    assert loaded["Smoke"]["B1"].has_formula()


if __name__ == "__main__":
    main()
