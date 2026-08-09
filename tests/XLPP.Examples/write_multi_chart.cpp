#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>
void testWriteMultiChart() {
    xlpp::Workbook wb; auto& ws=wb.addWorksheet("Charts"); ws.append({std::string("Category"),std::string("A"),std::string("B")});
    ws.append({std::string("Q1"),10.0,15.0}); ws.append({std::string("Q2"),20.0,18.0}); ws.append({std::string("Q3"),14.0,25.0});
    for (auto type : {xlpp::Chart::Type::Bar, xlpp::Chart::Type::Line, xlpp::Chart::Type::Pie}) { xlpp::Chart ch(type); ch.setTitle(type==xlpp::Chart::Type::Bar?"Bar":type==xlpp::Chart::Type::Line?"Line":"Pie"); auto& s=ch.addSeries(xlpp::ChartSeries("Series A")); s.setCategoriesReference("'Charts'!$A$2:$A$4"); s.setValuesReference("'Charts'!$B$2:$B$4"); ws.addChart(ch); }
    wb.save(xlpp_numbered_tests::outputPath("32_multi_chart.xlsx")); std::cout<<"Saved: 32_multi_chart.xlsx\n";
}
