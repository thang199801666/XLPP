// Comparable write benchmark for XLPP and libxlsxwriter.
#include <XLPP/XLPP.h>
#include <xlsxwriter.h>

#include <chrono>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <stdexcept>
#include <string>

using Clock = std::chrono::steady_clock;
using xlpp::CellValue;

static double elapsed_ms(Clock::time_point start, Clock::time_point end) {
    return std::chrono::duration<double, std::milli>(end - start).count();
}

static std::string text(int row, int col) {
    if (col == 0) return "Item-" + std::to_string(row);
    if (col == 1) return std::to_string(row + 1);
    if (col == 2) return std::to_string((row * 7919 % 10000000) / 100.0);
    return "Lorem ipsum dolor sit amet";
}

static double write_xlpp(const std::filesystem::path& path, int rows, int cols) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    auto start = Clock::now();
    for (int row = 0; row < rows; ++row) {
        std::vector<CellValue> values;
        values.reserve(cols);
        for (int col = 0; col < cols; ++col) {
            if (col == 0 || col >= 3) values.emplace_back(text(row, col));
            else if (col == 1) values.emplace_back(static_cast<double>(row + 1));
            else values.emplace_back((row * 7919 % 10000000) / 100.0);
        }
        sheet.append(values);
    }
    sheet.cell("A1").font().setBold(true);
    sheet.freezePanes("A2");
    workbook.save(path);
    return elapsed_ms(start, Clock::now());
}

static double write_xlsxwriter(const std::filesystem::path& path, int rows, int cols) {
    auto start = Clock::now();
    lxw_workbook* workbook = workbook_new(path.string().c_str());
    if (!workbook) throw std::runtime_error("libxlsxwriter could not create workbook");
    lxw_worksheet* sheet = workbook_add_worksheet(workbook, "Data");
    lxw_format* bold = workbook_add_format(workbook);
    format_set_bold(bold);
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            if (col == 1 || col == 2) {
                const double number = col == 1
                    ? static_cast<double>(row + 1)
                    : (row * 7919 % 10000000) / 100.0;
                worksheet_write_number(sheet, row, col, number, row == 0 ? bold : nullptr);
                continue;
            }
            const auto value = text(row, col);
            worksheet_write_string(sheet, row, col, value.c_str(), row == 0 ? bold : nullptr);
        }
    }
    worksheet_freeze_panes(sheet, 1, 0);
    if (workbook_close(workbook) != LXW_ERROR_NO_ERROR)
        throw std::runtime_error("libxlsxwriter could not close workbook");
    return elapsed_ms(start, Clock::now());
}

int main() {
    constexpr int rows = 10000;
    constexpr int cols = 15;
    const auto directory = std::filesystem::temp_directory_path();
    const auto xlpp_path = directory / "xlpp_external_benchmark.xlsx";
    const auto xlsxwriter_path = directory / "xlsxwriter_benchmark.xlsx";

    const double xlpp_ms = write_xlpp(xlpp_path, rows, cols);
    const double xlsxwriter_ms = write_xlsxwriter(xlsxwriter_path, rows, cols);
    std::cout << "BENCHMARK,c++,XLPP,write," << std::fixed << std::setprecision(2) << xlpp_ms
              << "," << std::filesystem::file_size(xlpp_path) << "\n";
    std::cout << "BENCHMARK,c++,libxlsxwriter,write," << std::fixed << std::setprecision(2)
              << xlsxwriter_ms << "," << std::filesystem::file_size(xlsxwriter_path) << "\n";
    std::filesystem::remove(xlpp_path);
    std::filesystem::remove(xlsxwriter_path);
}
