#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteCopySheet() {
    xlpp::Workbook workbook;
    auto& source = workbook.addWorksheet("Source");
    source.cell("A1").setValue("Copied value");
    workbook.copyWorksheet(source, "Copy");
    workbook.removeWorksheet("Source");
    workbook.save(xlpp_numbered_tests::outputPath("23_copy_sheet.xlsx"));
    std::cout << "Saved: 23_copy_sheet.xlsx\n";
}
