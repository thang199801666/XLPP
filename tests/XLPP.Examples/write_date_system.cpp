#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteDateSystem() {
    xlpp::Workbook workbook;
    workbook.setDate1904(true);
    auto& sheet = workbook.addWorksheet("Dates");
    sheet.cell("A1").setValue(xlpp::DateTime(2024, 3, 15));
    workbook.save(xlpp_numbered_tests::outputPath("21_date_system.xlsx"));
    std::cout << "Saved: 21_date_system.xlsx\n";
}
