"""Test writing a chart to an Excel workbook."""

from pathlib import Path
import xlpp


def test_chart_write_and_read(tmp_path: Path):
    workbook = xlpp.Workbook()
    worksheet = workbook.add_worksheet("ChartData")
    worksheet.append(["Month", "Sales"])
    worksheet.append(["Jan", 120])
    worksheet.append(["Feb", 175])
    worksheet.append(["Mar", 150])

    chart = xlpp.Chart(xlpp.ChartType.BAR)
    chart.title = "Monthly sales"
    chart.x_axis_title = "Month"
    chart.y_axis_title = "Sales"
    series = chart.add_series(xlpp.ChartSeries("Sales"))
    series.categories_reference = "'ChartData'!$A$2:$A$4"
    series.values_reference = "'ChartData'!$B$2:$B$4"
    worksheet.add_chart(chart)

    path = tmp_path / "chart.xlsx"
    workbook.save(str(path))
    assert path.exists()
    assert worksheet.chart_count == 1

    loaded = xlpp.Workbook()
    loaded.load(str(path))
    assert loaded["ChartData"].chart_count == 1
