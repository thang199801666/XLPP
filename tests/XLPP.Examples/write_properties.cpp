#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>
#include <stdexcept>

void testWriteProperties() {
    const auto path = xlpp_numbered_tests::outputPath("13_properties.xlsx");
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Properties");
    sheet.append({std::string("Property"), std::string("Value")});
    sheet.append({std::string("Title"), std::string("XLPP properties example")});
    sheet.append({std::string("Creator"), std::string("XLPP tests")});
    sheet.append({std::string("Subject"), std::string("Workbook properties")});
    sheet.append({std::string("Category"), std::string("Examples")});
    sheet.columnDimension("A").width = 20.0;
    sheet.columnDimension("B").width = 34.0;

    workbook.properties().setTitle("XLPP properties example");
    workbook.properties().setCreator("XLPP tests");
    workbook.properties().setSubject("Workbook properties");
    workbook.properties().setDescription("Core document properties written and loaded by XL++");
    workbook.properties().setKeywords("XLPP, OOXML, properties");
    workbook.properties().setCategory("Examples");
    workbook.properties().setLastModifiedBy("XLPP round-trip test");
    workbook.save(path);

    xlpp::Workbook loaded;
    loaded.load(path);
    if (loaded.properties().title() != "XLPP properties example" ||
        loaded.properties().lastModifiedBy() != "XLPP round-trip test")
        throw std::runtime_error("Document properties did not round-trip");
    std::cout << "Saved and verified: 13_properties.xlsx\n";
}
