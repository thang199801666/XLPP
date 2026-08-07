#include <XLPP/XLPP.h>
#include "TestOutput.h"

#include <filesystem>
#include <iostream>
#include <stdexcept>
#include <string>

void testWriteVbaText() {
    static const std::string source =
        "Option Explicit\n"
        "\n"
        "Public Sub XLPP_Hello()\n"
        "    With ThisWorkbook.Worksheets(\"MacroWorkbook\")\n"
        "        .Range(\"B2\").Value = \"Hello from VBA source text\"\n"
        "        .Range(\"B3\").Value = 2026\n"
        "    End With\n"
        "End Sub\n";

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("MacroWorkbook");
    sheet.cell("A1").setValue("This workbook contains VBA generated directly from source text.");
    sheet.cell("A2").setValue("Run macro: XLPP_Hello");
    workbook.setVbaModuleText("XLPPGenerated", source);

    const auto output = xlpp_numbered_tests::outputPath("37_vba_text.xlsm");
    workbook.save(output);

    // Read verification: load the package again and extract the module source from vbaProject.bin.
    xlpp::Workbook loaded;
    loaded.load(output);
    const auto loadedSource = loaded.vbaModuleText("XLPPGenerated");
    if (!loadedSource || loadedSource->find("Public Sub XLPP_Hello()") == std::string::npos ||
        loadedSource->find("Hello from VBA source text") == std::string::npos) {
        throw std::runtime_error("Generated VBA source could not be read back from the XLSM package");
    }

    std::cout << "Saved and verified VBA source text: " << output << '\n';
}
