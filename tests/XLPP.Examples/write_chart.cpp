#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteChart() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("ChartData");
    sheet.append({std::string("Month"), std::string("Sales")});
    sheet.append({std::string("Jan"), 120.0});
    sheet.append({std::string("Feb"), 175.0});
    sheet.append({std::string("Mar"), 150.0});
    xlpp::Chart chart(xlpp::Chart::Type::Bar);
    chart.setTitle("Monthly sales");
    auto& series = chart.addSeries(xlpp::ChartSeries("Sales"));
    series.setCategoriesReference("'ChartData'!$A$2:$A$4");
    series.setValuesReference("'ChartData'!$B$2:$B$4");
    sheet.addChart(chart);
    workbook.save(xlpp_numbered_tests::outputPath("25_chart.xlsx"));
    std::cout << "Saved: 25_chart.xlsx\n";
}
