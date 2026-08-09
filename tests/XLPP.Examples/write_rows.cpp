#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteRows() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Name"), std::string("Quantity"), std::string("Price")});
    sheet.append({std::string("Keyboard"), 2.0, 25.5});
    sheet.append({std::string("Mouse"), 3.0, 12.0});
    workbook.save(xlpp_numbered_tests::outputPath("02_rows.xlsx"));
    std::cout << "Saved: 02_rows.xlsx\n";
}
