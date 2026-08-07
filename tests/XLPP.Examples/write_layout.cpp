#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteLayout() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Layout");
    sheet.cell("A1").setValue("Merged title");
    sheet.mergeCells("A1:D1");
    sheet.freezePanes("A2");
    sheet.setPrintArea("A1:D20");
    workbook.save(xlpp_numbered_tests::outputPath("07_layout.xlsx"));
    std::cout << "Saved: 07_layout.xlsx\n";
}
