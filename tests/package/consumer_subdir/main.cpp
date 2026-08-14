#include <XLPP/XLPP.h>
#include <filesystem>
#include <iostream>

int main() {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_consumer_subdir_out.xltx";
    {
        xlpp::Workbook wb;
        auto& ws = wb.addWorksheet("Sheet");
        ws.cell("A1").setValue("from subdir");
        ws.cell("B1").setValue(7.5);
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
        if (loaded.chartsheetCount() != 1 || loaded.workbookSheetNames().front() != "Dashboard") return 1;
        const auto* dashboard = loaded.chartsheet("Dashboard");
        if (!dashboard || !dashboard->printerSettingsData() || dashboard->printerSettingsData()->size() != 4) return 1;
        if (!loaded.isTemplate() || loaded.activeWorkbookSheetIndex() != 0) return 1;
        if (loaded.workbookSheetVisibility(1) != xlpp::WorkbookSheetVisibility::Hidden) return 1;
        const auto& sheet = loaded.worksheet("Sheet");
        if (sheet == nullptr) return 1;
        if (sheet->cell("A1").stringValueOr("") != "from subdir") return 1;
        if (sheet->cell("B1").numericValueOr(-1) != 7.5) return 2;
    }
    std::filesystem::remove(path);
    std::cout << "subdir consumer OK" << std::endl;
    return 0;
}
