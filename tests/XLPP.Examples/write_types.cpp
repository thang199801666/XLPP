#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteTypes() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Types");
    sheet.cell("A1").setValue("Text");
    sheet.cell("A2").setValue(123.45);
    sheet.cell("A3").setValue(true);
    sheet.cell("A4").setValue(xlpp::DateTime(2024, 1, 15));
    sheet.cell("A5").setError(xlpp::CellError::Value);
    workbook.save(xlpp_numbered_tests::outputPath("05_types.xlsx"));
    std::cout << "Saved: 05_types.xlsx\n";
}
