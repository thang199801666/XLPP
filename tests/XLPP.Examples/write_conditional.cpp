#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

void testWriteConditional() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Conditional");
    sheet.append({std::string("Item"), std::string("Value")});
    sheet.append({std::string("Profit"), 25.0});
    sheet.append({std::string("Loss"), -12.0});
    sheet.append({std::string("Break even"), 0.0});
    sheet.columnDimension("A").width = 18.0;
    sheet.columnDimension("B").width = 14.0;

    auto negative = xlpp::ConditionalRule::formula("B2<0");
    negative.setStopIfTrue(true);
    negative.differentialStyle().font().color().setArgb("FF9C0006");
    negative.differentialStyle().fill().setPatternType("solid");
    negative.differentialStyle().fill().foregroundColor().setArgb("FFFFC7CE");
    sheet.conditionalFormatting().addRule("B2:B4", std::move(negative));

    auto positive = xlpp::ConditionalRule::formula("B2>0");
    positive.differentialStyle().font().color().setArgb("FF006100");
    positive.differentialStyle().fill().setPatternType("solid");
    positive.differentialStyle().fill().foregroundColor().setArgb("FFC6EFCE");
    sheet.conditionalFormatting().addRule("B2:B4", std::move(positive));

    workbook.save(xlpp_numbered_tests::outputPath("12_conditional.xlsx"));
    std::cout << "Saved: 12_conditional.xlsx\n";
}
