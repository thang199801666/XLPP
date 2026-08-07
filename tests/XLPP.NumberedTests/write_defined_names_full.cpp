#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

void testWriteDefinedNamesFull() {
    xlpp::Workbook wb;
    auto& ws = wb.addWorksheet("Names");
    ws.append({std::string("Amount A"), std::string("Amount B"), std::string("Named formula")});
    ws.cell("A2").setValue(100.0);
    ws.cell("B2").setValue(200.0);

    xlpp::DefinedName global("GlobalAmount", "'Names'!$A$2:$B$2");
    global.setComment("Global range");
    wb.addDefinedName(std::move(global));

    xlpp::DefinedName local("LocalAmount", "'Names'!$B$2");
    local.setLocalSheetId(0);
    local.setHidden(true);
    wb.addDefinedName(std::move(local));

    ws.cell("C2").setFormula("SUM(GlobalAmount)");
    wb.save(xlpp_numbered_tests::outputPath("29_defined_names_full.xlsx"));
    std::cout << "Saved: 29_defined_names_full.xlsx\n";
}
