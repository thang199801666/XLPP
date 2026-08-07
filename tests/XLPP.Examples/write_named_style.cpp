#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

void testWriteNamedStyle() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("NamedStyle");

    xlpp::Style header;
    header.font().setBold(true);
    header.font().setSize(12.0);
    header.font().color().setArgb("FFFFFFFF");
    header.fill().setPatternType("solid");
    header.fill().foregroundColor().setArgb("FF2F75B5");
    header.alignment().setHorizontal("center");
    header.alignment().setVertical("center");
    for (auto* side : {&header.border().left(), &header.border().right(), &header.border().top(), &header.border().bottom()}) {
        side->setStyle("thin");
        side->color().setArgb("FF1F1F1F");
    }
    workbook.addNamedStyle({"Header", header});

    for (std::size_t column = 1; column <= 4; ++column) {
        auto& cell = sheet.cell(1, column);
        cell.setValue(std::string("Column ") + std::to_string(column));
        workbook.applyNamedStyle(cell, "Header");
        sheet.columnDimension(column).width = 16.0;
    }
    sheet.rowDimension(1).height = 24.0;
    workbook.save(xlpp_numbered_tests::outputPath("16_named_style.xlsx"));
    std::cout << "Saved: 16_named_style.xlsx\n";
}
