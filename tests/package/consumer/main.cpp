#include <XLPP/XLPP.h>
#include <filesystem>
#include <iostream>

int main() {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_consumer_out.xlsx";
    {
        xlpp::Workbook wb;
        auto& ws = wb.addWorksheet("Data");
        ws.cell("A1").setValue(42.0);
        ws.cell("B1").setValue("hello");
        ws.cell("C1").setValue(true);
        wb.save(path);
    }
    {
        xlpp::Workbook loaded;
        loaded.load(path);
        if (loaded.sheetCount() != 1) return 1;
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
