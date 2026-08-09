#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

void testWriteSheetView() {
    xlpp::Workbook wb;
    auto& ws = wb.addWorksheet("View");
    ws.append({std::string("Sheet view feature"), std::string("Expected value")});
    ws.append({std::string("Zoom scale"), 125.0});
    ws.append({std::string("Normal zoom"), 90.0});
    ws.append({std::string("Grid lines"), std::string("Hidden")});
    ws.append({std::string("Direction"), std::string("Right-to-left")});
    ws.append({std::string("Frozen pane"), std::string("C4")});

    auto& view = ws.sheetView();
    view.setTabColor("FF00AA55");
    view.setZoomScale(125);
    view.setZoomScaleNormal(90);
    view.setShowGridLines(false);
    view.setRightToLeft(true);
    view.setShowOutlineSymbols(false);
    ws.freezePanes("C4");

    wb.save(xlpp_numbered_tests::outputPath("28_sheet_view.xlsx"));
    std::cout << "Saved: 28_sheet_view.xlsx\n";
}
