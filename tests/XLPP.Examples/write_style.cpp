#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

namespace {
void setThinBorder(xlpp::Cell& cell, const std::string& color) {
    for (auto* side : {&cell.border().left(), &cell.border().right(), &cell.border().top(), &cell.border().bottom()}) {
        side->setStyle("thin");
        side->color().setArgb(color);
    }
}
}

void testWriteStyle() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Styles");
    sheet.append({std::string("Feature"), std::string("Result")});

    auto& fontCell = sheet.cell("B2");
    sheet.cell("A2").setValue("Font color");
    fontCell.setValue("Red, bold, italic");
    fontCell.font().setBold(true);
    fontCell.font().setItalic(true);
    fontCell.font().setSize(14.0);
    fontCell.font().color().setArgb("FFFF0000");

    auto& fillCell = sheet.cell("B3");
    sheet.cell("A3").setValue("Cell fill");
    fillCell.setValue("Yellow background");
    fillCell.fill().setPatternType("solid");
    fillCell.fill().foregroundColor().setArgb("FFFFFF00");

    auto& alignCell = sheet.cell("B4");
    sheet.cell("A4").setValue("Alignment");
    alignCell.setValue("Centered and wrapped text");
    alignCell.alignment().setHorizontal("center");
    alignCell.alignment().setVertical("center");
    alignCell.alignment().setWrapText(true);

    auto& borderCell = sheet.cell("B5");
    sheet.cell("A5").setValue("Border");
    borderCell.setValue("Four blue borders");
    setThinBorder(borderCell, "FF0070C0");

    auto& numberCell = sheet.cell("B6");
    sheet.cell("A6").setValue("Number format");
    numberCell.setValue(1234.567);
    numberCell.setNumberFormat("#,##0.00");

    sheet.columnDimension("A").width = 22.0;
    sheet.columnDimension("B").width = 34.0;
    sheet.rowDimension(4).height = 36.0;
    sheet.cell("A1").font().setBold(true);
    sheet.cell("B1").font().setBold(true);
    workbook.save(xlpp_numbered_tests::outputPath("04_style.xlsx"));
    std::cout << "Saved: 04_style.xlsx\n";
}
