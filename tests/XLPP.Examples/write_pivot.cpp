#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

void testWritePivot() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("PivotData");
    sheet.append({std::string("Quarter"), std::string("Amount")});
    sheet.append({std::string("Q1"), 10.0});
    sheet.append({std::string("Q2"), 20.0});

    xlpp::PivotTable pivot("SalesPivot");
    pivot.setLocation("D2");
    pivot.cache().setSourceData("'PivotData'!$A$1:$B$3");
    // Cache fields and records are inferred from the source range.
    pivot.addRowField("Quarter");
    pivot.addDataField("Amount", "sum");
    sheet.addPivotTable(std::move(pivot));

    workbook.save(xlpp_numbered_tests::outputPath("33_pivot.xlsx"));
    std::cout << "Saved: 33_pivot.xlsx\n";
}
