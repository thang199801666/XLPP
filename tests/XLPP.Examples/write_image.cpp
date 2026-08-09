#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteImage(const std::filesystem::path& imagePath) {
    const auto output = xlpp_numbered_tests::outputPath("36_image.xlsx");
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Images");
    sheet.addImage(xlpp::Image::fromFile(imagePath, "D2"));
    workbook.save(output);
    std::cout << "Saved: " << output << '\n';
}
