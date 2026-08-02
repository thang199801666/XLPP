// XL++ Native Performance Benchmark — Public API end-to-end
#include <XLPP/XLPP.h>
#include <iostream>
#include <chrono>
#include <iomanip>
#include <filesystem>
#include <random>

using namespace xlpp;
using Clock = std::chrono::high_resolution_clock;
using Ms = std::chrono::duration<double, std::milli>;
static double ms(Clock::duration d) { return Ms(d).count(); }

static auto genData(int rows, int cols) {
    std::vector<std::vector<CellValue>> data;
    data.reserve(rows);
    std::mt19937 rng(42);
    std::uniform_real_distribution<double> dist(0, 100000);
    for (int r = 0; r < rows; ++r) {
        std::vector<CellValue> row;
        row.reserve(cols);
        for (int c = 0; c < cols; ++c) {
            if (c == 0) row.push_back("Item-" + std::to_string(r));
            else if (c == 1) row.push_back(static_cast<double>(r + 1));
            else if (c == 2) row.push_back(std::round(dist(rng) * 100) / 100);
            else row.push_back("Lorem ipsum dolor sit amet");
        }
        data.push_back(std::move(row));
    }
    return data;
}

void bench(int rows, int cols, const char* label) {
    std::cout << "\n=== " << label << " (" << rows << "x" << cols << "=" << rows*cols << " cells) ===\n";
    auto data = genData(rows, cols);
    auto tmp = std::filesystem::temp_directory_path() / "xlpp_bench.xlsx";

    // Build
    auto t0 = Clock::now();
    Workbook wb;
    auto& ws = wb.addWorksheet("Data");
    for (const auto& row : data) ws.append(row);
    ws.cell("A1").font().setBold(true);
    ws.cell("A1").font().setSize(14);
    ws.freezePanes("A2");
    auto t1 = Clock::now();
    std::cout << "  Build:  " << std::fixed << std::setw(8) << std::setprecision(1) << ms(t1-t0) << " ms\n";

    // Save sequential
    SaveOptions seqOpt;
    auto t2 = Clock::now();
    wb.save(tmp, seqOpt);
    auto t3 = Clock::now();
    auto sz = std::filesystem::file_size(tmp);
    std::cout << "  Save:   " << std::fixed << std::setw(8) << std::setprecision(1) << ms(t3-t2) << " ms  (" << sz << " bytes)\n";

    // Save parallel (4 threads)
    if (rows >= 5000) {
        SaveOptions parOpt;
        parOpt.parallelWorkers = 4;
        parOpt.parallelSheets = true;
        auto tp = Clock::now();
        wb.save(tmp, parOpt);
        auto tq = Clock::now();
        std::cout << "  Save(4):" << std::fixed << std::setw(8) << std::setprecision(1) << ms(tq-tp) << " ms\n";
    }

    // Load
    auto t4 = Clock::now();
    Workbook loaded;
    loaded.load(tmp);
    auto t5 = Clock::now();
    std::cout << "  Load:   " << std::fixed << std::setw(8) << std::setprecision(1) << ms(t5-t4) << " ms\n";

    // Verify
    auto* lws = loaded.worksheet("Data");
    if (lws)
    std::cout << "  Verify: " << lws->maxRow() << " rows, " << lws->maxColumn() << " cols, A1='"
              << lws->cell("A1").stringValueOr("?") << "'\n";

    std::filesystem::remove(tmp);
}

int main() {
    std::cout << "XL++ Native Performance Benchmarks\n";
    std::cout << "===================================\n";

    bench(1000, 10,  "1K");
    bench(10000, 15, "10K");
    bench(50000, 10, "50K");
    bench(100000, 10, "100K");

    std::cout << "\nDone.\n";
    return 0;
}
