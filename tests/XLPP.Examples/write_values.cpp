#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteValues() {
    const auto output = xlpp_numbered_tests::outputPath("00_values.xlsx");
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.cell("A1").setValue("Name");
    sheet.cell("B1").setValue("Value");
    sheet.cell("A2").setValue("XLPP");
    sheet.cell("B2").setValue(42.0);
    sheet.cell("A1").font().setBold(true);
    workbook.save(output);
    std::cout << "Saved: 00_values.xlsx\n";
}
