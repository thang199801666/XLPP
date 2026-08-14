#include <XLPP/XLPP.h>
#include <filesystem>
#include <iostream>
#include <vector>

int main() {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_consumer_out.xltx";
    {
        xlpp::Workbook wb;
        auto& ws = wb.addWorksheet("Data");
        ws.cell("A1").setValue(42.0);
        ws.cell("B1").setValue("hello");
        ws.cell("C1").setValue(true);
        xlpp::Chart dashboard(xlpp::Chart::Type::Bar);
        dashboard.setTitle("Consumer Dashboard");
        auto& dashboardSheet = wb.addChartsheet("Dashboard", std::move(dashboard));
        dashboardSheet.setPrinterSettingsData(std::string("\x10\x00\x20\x30", 4));
        wb.moveWorkbookSheet(1, 0);
        wb.setTemplate(true);
        wb.setActiveWorkbookSheet("Dashboard");
        wb.setWorkbookSheetVisibility(1, xlpp::WorkbookSheetVisibility::Hidden);
        wb.save(path);
    }
    {
        xlpp::Workbook loaded;
        loaded.load(path);
        if (loaded.sheetCount() != 1 || loaded.chartsheetCount() != 1) return 1;
        const auto* dashboard = loaded.chartsheet("Dashboard");
        if (!dashboard || !dashboard->printerSettingsData() || dashboard->printerSettingsData()->size() != 4) return 1;
        if (!loaded.isTemplate() || loaded.activeWorkbookSheetIndex() != 0) return 1;
        if (loaded.workbookSheetVisibility(1) != xlpp::WorkbookSheetVisibility::Hidden) return 1;
        if (loaded.workbookSheetNames() != std::vector<std::string>{"Dashboard", "Data"}) return 1;
        const auto& sheet = loaded.worksheet("Data");
        if (sheet == nullptr) return 1;
        if (sheet->cell("A1").numericValueOr(-1) != 42.0) return 2;
        if (sheet->cell("B1").stringValueOr("") != "hello") return 3;
        if (!sheet->cell("C1").isBoolean()) return 4;
    }
    std::filesystem::remove(path);
    std::cout << "consumer OK" << std::endl;
    return 0;
}
