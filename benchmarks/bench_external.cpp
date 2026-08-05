// Comparable write benchmark for XLPP and libxlsxwriter.
#include <XLPP/XLPP.h>
#include <XLPP/Streaming.h>

#include <chrono>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <stdexcept>
#include <string>

using Clock = std::chrono::steady_clock;
using xlpp::CellValue;

// The vcpkg libxlsxwriter package omits its private queue/tree headers. These
// are the public C ABI declarations needed by this write-only benchmark.
extern "C" {
struct lxw_workbook;
struct lxw_worksheet;
struct lxw_format;
lxw_workbook* workbook_new(const char* filename);
lxw_worksheet* workbook_add_worksheet(lxw_workbook*, const char* name);
lxw_format* workbook_add_format(lxw_workbook*);
void format_set_bold(lxw_format*);
int worksheet_write_string(lxw_worksheet*, int row, int col, const char*, lxw_format*);
int worksheet_write_number(lxw_worksheet*, int row, int col, double, lxw_format*);
int worksheet_write_formula(lxw_worksheet*, int row, int col, const char*, double, lxw_format*);
int worksheet_freeze_panes(lxw_worksheet*, int row, int col);
int workbook_close(lxw_workbook*);
}

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
    if (workbook_close(workbook) != 0)
        throw std::runtime_error("libxlsxwriter could not close workbook");
    return elapsed_ms(start, Clock::now());
}

static double write_xlpp_scenario(const std::filesystem::path& path, const char* scenario,
                                  int rows, int cols) {
    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data");
    auto& lookup = workbook.addWorksheet("Lookup");
    auto start = Clock::now();
    for (int row = 1; row <= 500; ++row) {
        lookup.cell(static_cast<std::size_t>(row), 1).setValue("Item-" + std::to_string(row - 1));
        lookup.cell(static_cast<std::size_t>(row), 2).setValue(row * 1.25);
    }
    for (int row = 1; row <= rows; ++row) {
        for (int col = 1; col <= cols; ++col) {
            auto& cell = data.cell(static_cast<std::size_t>(row), static_cast<std::size_t>(col));
            if (std::string(scenario) == "formula" && col == cols)
                cell.setFormula("=VLOOKUP(A" + std::to_string(row) +
                                ",Lookup!$A$1:$B$500,2,FALSE)");
            else if (col == 1 || col >= 4) cell.setValue(text(row - 1, col - 1));
            else if (col == 2) cell.setValue(static_cast<double>(row));
            else cell.setValue((row * 7919 % 10000000) / 100.0);
        }
    }
    data.cell("A1").font().setBold(true);
    data.freezePanes("A2");
    workbook.save(path);
    return elapsed_ms(start, Clock::now());
}

static double write_xlsxwriter_scenario(const std::filesystem::path& path, const char* scenario,
                                        int rows, int cols) {
    auto start = Clock::now();
    lxw_workbook* workbook = workbook_new(path.string().c_str());
    lxw_worksheet* data = workbook_add_worksheet(workbook, "Data");
    lxw_worksheet* lookup = workbook_add_worksheet(workbook, "Lookup");
    for (int row = 0; row < 500; ++row) {
        worksheet_write_string(lookup, row, 0, ("Item-" + std::to_string(row)).c_str(), nullptr);
        worksheet_write_number(lookup, row, 1, (row + 1) * 1.25, nullptr);
    }
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            if (std::string(scenario) == "formula" && col == cols - 1) {
                const auto formula = "=VLOOKUP(A" + std::to_string(row + 1) +
                    ",Lookup!$A$1:$B$500,2,FALSE)";
                worksheet_write_formula(data, row, col, formula.c_str(), 0, nullptr);
            } else if (col == 1 || col == 2) {
                worksheet_write_number(data, row, col,
                    col == 1 ? row + 1 : (row * 7919 % 10000000) / 100.0, nullptr);
            } else {
                const auto cell = text(row, col);
                worksheet_write_string(data, row, col, cell.c_str(), nullptr);
            }
        }
    }
    worksheet_freeze_panes(data, 1, 0);
    if (workbook_close(workbook) != 0) throw std::runtime_error("scenario save failed");
    return elapsed_ms(start, Clock::now());
}

static void run_scenario(const char* scenario, int rows, int cols) {
    const auto directory = std::filesystem::temp_directory_path();
    const auto xlpp_path = directory / (std::string("xlpp-") + scenario + ".xlsx");
    const auto xlsxwriter_path = directory / (std::string("xlsxwriter-") + scenario + ".xlsx");
    const auto xlpp_ms = write_xlpp_scenario(xlpp_path, scenario, rows, cols);
    const auto xlsxwriter_ms = write_xlsxwriter_scenario(xlsxwriter_path, scenario, rows, cols);
    xlpp::Workbook loaded;
    const auto read_start = Clock::now();
    loaded.load(xlpp_path);
    const auto xlpp_read_ms = elapsed_ms(read_start, Clock::now());
    std::cout << "BENCHMARK,c++,XLPP," << scenario << "_write," << std::fixed
              << std::setprecision(2) << xlpp_ms << "," << std::filesystem::file_size(xlpp_path) << "\n";
    std::cout << "BENCHMARK,c++,XLPP," << scenario << "_read," << std::fixed
              << std::setprecision(2) << xlpp_read_ms << ",0\n";
    std::cout << "BENCHMARK,c++,libxlsxwriter," << scenario << "_write," << std::fixed
              << std::setprecision(2) << xlsxwriter_ms << ","
              << std::filesystem::file_size(xlsxwriter_path) << "\n";
    std::filesystem::remove(xlpp_path);
    std::filesystem::remove(xlsxwriter_path);
}

static void run_streaming_read(int rows) {
    const auto path = std::filesystem::temp_directory_path() /
        ("xlpp-streaming-" + std::to_string(rows) + ".xlsx");
    {
        xlpp::StreamingWorkbookWriter writer(path, xlpp::SharedStringMode::Disabled);
        auto& sheet = writer.addWorksheet("Data");
        for (int row = 0; row < rows; ++row) {
            std::vector<CellValue> values;
            values.reserve(10);
            for (int col = 0; col < 10; ++col)
                values.emplace_back(col == 0 ? text(row, col) : std::to_string(row + col));
            sheet.append(values);
        }
        writer.close();
    }
    const auto size = std::filesystem::file_size(path);
    const auto start = Clock::now();
    xlpp::StreamingWorkbookReader reader(path);
    std::size_t cells = 0;
    reader.forEachRow("Data", [&cells](std::size_t, const xlpp::StreamingRow& row) {
        cells += row.size();
        return true;
    });
    const auto read_ms = elapsed_ms(start, Clock::now());
    std::cout << "BENCHMARK,c++,XLPP,streaming_" << rows << "_read," << std::fixed
              << std::setprecision(2) << read_ms << "," << size << "\n";
    std::filesystem::remove(path);
}

int main() {
    constexpr int rows = 10000;
    constexpr int cols = 15;
    const auto directory = std::filesystem::temp_directory_path();
    const auto xlpp_path = directory / "xlpp_external_benchmark.xlsx";
    const auto xlsxwriter_path = directory / "xlsxwriter_benchmark.xlsx";

    // Warm up each serializer so one-time initialization is not charged to a case.
    const auto warmup_xlpp = directory / "xlpp-warmup.xlsx";
    const auto warmup_xlsxwriter = directory / "xlsxwriter-warmup.xlsx";
    write_xlpp(warmup_xlpp, 20, 5);
    write_xlsxwriter(warmup_xlsxwriter, 20, 5);
    std::filesystem::remove(warmup_xlpp);
    std::filesystem::remove(warmup_xlsxwriter);

    const double xlpp_ms = write_xlpp(xlpp_path, rows, cols);
    const double xlsxwriter_ms = write_xlsxwriter(xlsxwriter_path, rows, cols);
    std::cout << "BENCHMARK,c++,XLPP,write," << std::fixed << std::setprecision(2) << xlpp_ms
              << "," << std::filesystem::file_size(xlpp_path) << "\n";
    std::cout << "BENCHMARK,c++,libxlsxwriter,write," << std::fixed << std::setprecision(2)
              << xlsxwriter_ms << "," << std::filesystem::file_size(xlsxwriter_path) << "\n";
    const auto read_start = Clock::now();
    xlpp::Workbook loaded;
    loaded.load(xlpp_path);
    std::cout << "BENCHMARK,c++,XLPP,read," << std::fixed << std::setprecision(2)
              << elapsed_ms(read_start, Clock::now()) << ",0\n";
    std::filesystem::remove(xlpp_path);
    std::filesystem::remove(xlsxwriter_path);
    run_scenario("lookup", rows, cols);
    run_scenario("formula", rows, cols);
    run_streaming_read(100000);
    run_streaming_read(500000);
    run_streaming_read(1000000);
}
