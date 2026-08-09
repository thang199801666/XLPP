#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteLoadOptions() {
    const auto path = xlpp_numbered_tests::outputPath("24_load_options.xlsx");
    xlpp::Workbook workbook;
    workbook.addWorksheet("Load").cell("A1").setValue("Lenient load");
    workbook.save(path);
    xlpp::LoadOptions options;
    options.lenient = true;
    xlpp::Workbook loaded;
    loaded.load(path, options);
    std::cout << "Saved and loaded: 24_load_options.xlsx\n";
}
