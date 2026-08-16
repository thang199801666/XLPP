#include <XLPP/XLPP.h>
#include <chrono>
#include <cstddef>
#include <iostream>
#include <string>

int main(int argc, char** argv) {
    const std::size_t rows = argc > 1 ? static_cast<std::size_t>(std::stoull(argv[1])) : 500;
    const std::size_t cols = argc > 2 ? static_cast<std::size_t>(std::stoull(argv[2])) : 200;
    const std::string out = argc > 3 ? argv[3] : std::string("bench_xlpp_cpp.xlsx");
    const auto cellCount = rows * cols;

    auto t0 = std::chrono::steady_clock::now();
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    for (std::size_t r = 0; r < rows; ++r) {
        for (std::size_t c = 0; c < cols; ++c) {
            switch (c % 4) {
                case 0: sheet.cell(r + 1, c + 1).setValue(static_cast<double>(r * 1000 + c)); break;
                case 1: sheet.cell(r + 1, c + 1).setValue(static_cast<double>(r * 1000 + c) / 7.0); break;
                case 2: sheet.cell(r + 1, c + 1).setValue("text-" + std::to_string(r) + "-" + std::to_string(c)); break;
                default: sheet.cell(r + 1, c + 1).setValue(true); break;
            }
        }
    }
    auto tBuild = std::chrono::steady_clock::now();

    workbook.save(out);
    auto tSave = std::chrono::steady_clock::now();

    const auto buildMs = std::chrono::duration<double, std::milli>(tBuild - t0).count();
    const auto saveMs = std::chrono::duration<double, std::milli>(tSave - tBuild).count();
    const auto totalMs = std::chrono::duration<double, std::milli>(tSave - t0).count();
    std::cout << "xlpp C++ writer: build=" << buildMs << "ms save=" << saveMs
              << "ms total=" << totalMs << "ms  (" << cellCount << " cells, "
              << static_cast<double>(cellCount) / (totalMs / 1000.0) / 1e6 << "M cells/s)\n";
    return 0;
}
