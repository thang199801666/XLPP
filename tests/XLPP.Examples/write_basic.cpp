#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteBasic() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Sheet1");
    sheet.cell("A1").setValue("Hello XLPP");
    sheet.cell("B1").setValue(42.0);
    workbook.save(xlpp_numbered_tests::outputPath("01_basic.xlsx"));
    std::cout << "Saved: 01_basic.xlsx\n";
}
