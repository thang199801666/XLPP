#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

void testWriteFormulaMetadata() {
    xlpp::Workbook wb;
    auto& ws = wb.addWorksheet("FormulaMetadata");
    ws.append({std::string("Input A"), std::string("Input B"), std::string("Shared result"), std::string("Dynamic array")});
    ws.cell("A2").setValue(1.0);
    ws.cell("B2").setValue(2.0);
    ws.cell("A3").setValue(10.0);
    ws.cell("B3").setValue(20.0);

    // Shared formula master and follower use one shared index and a valid master range.
    ws.cell("C2").setSharedFormula("SUM(A2:B2)", 7, "C2:C3");
    ws.cell("C2").setValue(3.0);
    ws.cell("C3").setSharedFormula("", 7);
    ws.cell("C3").setValue(30.0);

    // TRANSPOSE of a horizontal 1x2 range produces a vertical 2x1 array.
    ws.cell("A5").setArrayFormula("TRANSPOSE(A2:B2)", "A5:A6");
    // SORT of a horizontal 1x2 range spills horizontally into two cells.
    ws.cell("D2").setDynamicArrayFormula("SORT(A2:B2)", "D2:E2");

    wb.save(xlpp_numbered_tests::outputPath("34_formula_metadata.xlsx"));
    std::cout << "Saved: 34_formula_metadata.xlsx\n";
}
