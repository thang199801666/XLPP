#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteArrayFormula() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("ArrayFormula");
    sheet.cell("A1").setValue(1.0);
    sheet.cell("A2").setValue(2.0);
    sheet.cell("B1").setArrayFormula("SUM(A1:A2)", "B1:B1");
    sheet.cell("C1").setDynamicArrayFormula("SORT(A1:A2)", "C1");
    workbook.save(xlpp_numbered_tests::outputPath("22_array_formula.xlsx"));
    std::cout << "Saved: 22_array_formula.xlsx\n";
}
