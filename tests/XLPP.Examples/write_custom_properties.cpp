#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteCustomProperties() {
    xlpp::Workbook workbook;
    workbook.addWorksheet("Custom");
    workbook.customProperties().add(xlpp::CustomProperty("Build", "Example"));
    workbook.customProperties().add(xlpp::CustomProperty("Version", 1));
    workbook.save(xlpp_numbered_tests::outputPath("20_custom_properties.xlsx"));
    std::cout << "Saved: 20_custom_properties.xlsx\n";
}
