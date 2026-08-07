#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

void testWriteDocumentPropertiesFull() {
    xlpp::Workbook wb;
    auto& ws = wb.addWorksheet("Metadata");
    auto& p = wb.properties();
    p.setTitle("XLPP feature workbook");
    p.setSubject("Complete metadata");
    p.setCreator("XLPP tests");
    p.setDescription("Core and extended document properties");
    p.setKeywords("xlsx,c++,xlpp");
    p.setCategory("Testing");
    p.setLastModifiedBy("XLPP CI");

    ws.append({std::string("Property"), std::string("Expected value")});
    ws.append({std::string("Title"), std::string("XLPP feature workbook")});
    ws.append({std::string("Subject"), std::string("Complete metadata")});
    ws.append({std::string("Creator"), std::string("XLPP tests")});
    ws.append({std::string("Description"), std::string("Core and extended document properties")});
    ws.append({std::string("Keywords"), std::string("xlsx,c++,xlpp")});
    ws.append({std::string("Category"), std::string("Testing")});
    ws.append({std::string("Last modified by"), std::string("XLPP CI")});

    wb.save(xlpp_numbered_tests::outputPath("31_document_properties_full.xlsx"));
    std::cout << "Saved: 31_document_properties_full.xlsx\n";
}
