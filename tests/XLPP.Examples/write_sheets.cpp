#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteSheets() {
    xlpp::Workbook workbook;
    workbook.addWorksheet("First").cell("A1").setValue("First sheet");
    workbook.addWorksheet("Second").cell("A1").setValue("Second sheet");
    workbook.addWorksheet("Third").cell("A1").setValue("Third sheet");
    workbook.save(xlpp_numbered_tests::outputPath("06_sheets.xlsx"));
    std::cout << "Saved: 06_sheets.xlsx\n";
}
