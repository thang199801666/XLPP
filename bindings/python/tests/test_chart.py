"""Test writing a chart to an Excel workbook."""


import os
import sys
from pathlib import Path

sys.path.insert(0, r"D:\Temp")
os.add_dll_directory(r"D:\Temp")
import xlpp

def test_chart_write_and_read():
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
    path = Path(r"D:\Code\CPlusPlus\XLPP\bindings\python\tests") / "chart.xlsx"
    workbook.save(str(path))
    assert path.exists()
    assert worksheet.chart_count == 1

test_chart_write_and_read()
