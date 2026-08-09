#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteFormula() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Formula");
    sheet.cell("A1").setValue(10.0);
    sheet.cell("B1").setValue(5.0);
    sheet.cell("C1").setFormula("=A1+B1");
    sheet.cell("D1").setFormula("=A1*B1");
    workbook.save(xlpp_numbered_tests::outputPath("03_formula.xlsx"));
    std::cout << "Saved: 03_formula.xlsx\n";
}
