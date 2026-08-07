#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>
#include <stdexcept>

void testWriteRoundtrip() {
    const auto path = xlpp_numbered_tests::outputPath("14_roundtrip.xlsx");
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Item"), std::string("Value"), std::string("Double")});
    sheet.append({std::string("A"), 21.0, std::monostate{}});
    sheet.cell("C2").setFormula("B2*2");
    sheet.cell("A1").font().setBold(true);
    sheet.cell("A1").fill().setPatternType("solid");
    sheet.cell("A1").fill().foregroundColor().setArgb("FF4472C4");
    sheet.cell("A1").font().color().setArgb("FFFFFFFF");
    sheet.columnDimension("A").width = 18.0;
    workbook.save(path);

    xlpp::Workbook loaded;
    loaded.load(path);
    auto* loadedSheet = loaded.worksheet("Data");
    if (!loadedSheet || loadedSheet->cell("C2").formula() != "B2*2")
        throw std::runtime_error("Initial round-trip failed");
    loadedSheet->cell("B3").setValue(7.0);
    loadedSheet->cell("C3").setFormula("B3*2");
    loadedSheet->cell("A3").setValue("Added after load");
    loaded.save(path);

    xlpp::Workbook verified;
    verified.load(path);
    auto* verifiedSheet = verified.worksheet("Data");
    if (!verifiedSheet || verifiedSheet->cell("C3").formula() != "B3*2" ||
        verifiedSheet->cell("A3").stringValueOr("") != "Added after load")
        throw std::runtime_error("Load-modify-save round-trip failed");
    std::cout << "Saved, loaded, modified, and reloaded: 14_roundtrip.xlsx\n";
}
