#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteStreaming() {
    const auto path = xlpp_numbered_tests::outputPath("19_streaming.xlsx");
    xlpp::StreamingWorkbookWriter writer(path);
    auto& sheet = writer.addWorksheet("Rows");
    for (int row = 1; row <= 100; ++row)
        sheet.append({std::string("Row"), static_cast<double>(row)});
    writer.close();
    std::cout << "Saved: 19_streaming.xlsx\n";
}
