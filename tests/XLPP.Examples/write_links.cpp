#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteLinks() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Links");
    sheet.cell("A1").setValue("Open website");
    sheet.cell("A1").setHyperlink(xlpp::Hyperlink("https://example.com"));
    sheet.cell("A2").setValue("Review note");
    sheet.cell("A2").setComment(xlpp::Comment("Check this cell", "Tester"));
    workbook.save(xlpp_numbered_tests::outputPath("10_links.xlsx"));
    std::cout << "Saved: 10_links.xlsx\n";
}
