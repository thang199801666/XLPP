#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteFilter() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Filter");
    sheet.append({std::string("Name"), std::string("Status")});
    sheet.append({std::string("Alpha"), std::string("Open")});
    sheet.append({std::string("Beta"), std::string("Closed")});
    sheet.autoFilter().setReference("A1:B3");
    workbook.save(xlpp_numbered_tests::outputPath("09_filter.xlsx"));
    std::cout << "Saved: 09_filter.xlsx\n";
}
