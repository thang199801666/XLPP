#include <XLPP/XLPP.h>
#include <filesystem>
#include <iostream>

int main() {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_consumer_subdir_out.xlsx";
    {
        xlpp::Workbook wb;
        auto& ws = wb.addWorksheet("Sheet");
        ws.cell("A1").setValue("from subdir");
        ws.cell("B1").setValue(7.5);
        wb.save(path);
    }
    {
        xlpp::Workbook loaded;
        loaded.load(path);
        const auto& sheet = loaded.worksheet("Sheet");
        if (sheet == nullptr) return 1;
        if (sheet->cell("A1").stringValueOr("") != "from subdir") return 1;
        if (sheet->cell("B1").numericValueOr(-1) != 7.5) return 2;
    }
    std::filesystem::remove(path);
    std::cout << "subdir consumer OK" << std::endl;
    return 0;
}
