import pytest

import xlpp


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
    assert np.array_equal(wb2["NP3"].to_array(), data)
