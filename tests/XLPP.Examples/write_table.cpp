#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteTable() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Table");
    sheet.append({std::string("Name"), std::string("Status")});
    sheet.append({std::string("Alpha"), std::string("Open")});
    sheet.append({std::string("Beta"), std::string("Closed")});
    auto& table = sheet.addTable("StatusTable", "A1:B3");
    table.styleInfo().setName("TableStyleMedium2");
    workbook.save(xlpp_numbered_tests::outputPath("08_table.xlsx"));
    std::cout << "Saved: 08_table.xlsx\n";
}
