#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteRange() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Range");
    for (std::size_t row = 1; row <= 5; ++row)
        for (std::size_t column = 1; column <= 3; ++column)
            sheet.cell(row, column).setValue(static_cast<double>(row * 10 + column));
    sheet.range("A1:C5").forEach([](xlpp::Cell& cell) { cell.setNumberFormat("0.00"); });
    workbook.save(xlpp_numbered_tests::outputPath("15_range.xlsx"));
    std::cout << "Saved: 15_range.xlsx\n";
}
