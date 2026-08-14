// XL++ Core Hot-Path Benchmark
//
// Focuses on operations that should remain close to O(number of inserted
// cells), plus repeated worksheet geometry queries which are expected to be
// O(1) once the extents cache is warm.
#include <XLPP/XLPP.h>

#include <chrono>
#include <cstddef>
#include <filesystem>
#include <iomanip>
#include <iostream>
#include <string>
#include <vector>
#include <cstdint>

#ifdef _WIN32
#  define NOMINMAX
#  include <windows.h>
#  include <psapi.h>
#else
#  include <sys/resource.h>
#endif

using Clock = std::chrono::steady_clock;
using Milliseconds = std::chrono::duration<double, std::milli>;

namespace {

double elapsedMs(Clock::time_point begin, Clock::time_point end) {
    return Milliseconds(end - begin).count();
}

std::uint64_t peakResidentBytes() noexcept {
#ifdef _WIN32
    PROCESS_MEMORY_COUNTERS_EX counters{};
    counters.cb = sizeof(counters);
    if (GetProcessMemoryInfo(GetCurrentProcess(),
                             reinterpret_cast<PROCESS_MEMORY_COUNTERS*>(&counters),
                             sizeof(counters))) {
        return static_cast<std::uint64_t>(counters.PeakWorkingSetSize);
    }
    return 0;
#else
    rusage usage{};
    if (getrusage(RUSAGE_SELF, &usage) != 0) return 0;
#  ifdef __APPLE__
    return static_cast<std::uint64_t>(usage.ru_maxrss);
#  else
    return static_cast<std::uint64_t>(usage.ru_maxrss) * 1024ull;
#  endif
#endif
}

std::size_t parseRows(int argc, char** argv) {
    constexpr std::size_t defaultRows = 30000;
    if (argc < 2) return defaultRows;
    try {
        const auto parsed = std::stoull(argv[1]);
        return parsed == 0 ? defaultRows : static_cast<std::size_t>(parsed);
    } catch (...) {
        return defaultRows;
    }
}

} // namespace

int main(int argc, char** argv) {
    const std::size_t rows = parseRows(argc, argv);
    constexpr std::size_t columns = 10;
    constexpr std::size_t geometryQueries = 20;

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");

    std::vector<xlpp::CellValue> row;
    row.reserve(columns);

    const auto buildStart = Clock::now();
    for (std::size_t r = 0; r < rows; ++r) {
        row.clear();
        for (std::size_t c = 0; c < columns; ++c) {
            if (c < 6) {
                row.emplace_back("Row-" + std::to_string(r) + "-Col-" + std::to_string(c));
            } else {
                row.emplace_back(static_cast<double>(r * columns + c));
            }
        }
        sheet.append(row);
    }
    const auto buildEnd = Clock::now();

    // First query may populate the cache. Repeated calls are intentionally
    // included so an accidental return to full cell-map scans is visible.
    std::string dimensions;
    const auto geometryStart = Clock::now();
    for (std::size_t i = 0; i < geometryQueries; ++i) {
        dimensions = sheet.dimensions();
        (void)sheet.maxRow();
        (void)sheet.maxColumn();
    }
    const auto geometryEnd = Clock::now();

    constexpr std::size_t trackingQueries = 5;
    std::size_t trackedChangeTotal = 0;
    const auto trackingStart = Clock::now();
    for (std::size_t i = 0; i < trackingQueries; ++i)
        trackedChangeTotal += sheet.trackedCellChangeCount();
    const auto trackingEnd = Clock::now();

    xlpp::SaveOptions options;
    options.compressionLevel = xlpp::CompressionLevel::Store;
    const auto output = std::filesystem::temp_directory_path() / "xlpp_core_hotpath_bench.xlsx";
    const auto saveStart = Clock::now();
    workbook.save(output, options);
    const auto saveEnd = Clock::now();

    const auto bytes = std::filesystem::file_size(output);
    std::filesystem::remove(output);

    const auto cells = rows * columns;
    std::cout << std::fixed << std::setprecision(3);
    std::cout << "METRIC,C++,XLPP,sizeof_cell_bytes," << sizeof(xlpp::Cell) << '\n';
    std::cout << "METRIC,C++,XLPP,sizeof_style_bytes," << sizeof(xlpp::Style) << '\n';
    std::cout << "METRIC,C++,XLPP,peak_rss_bytes," << peakResidentBytes() << '\n';
    std::cout << "BENCHMARK,C++,XLPP,core_bulk_build," << elapsedMs(buildStart, buildEnd)
              << ',' << cells << '\n';
    std::cout << "BENCHMARK,C++,XLPP,geometry_queries," << elapsedMs(geometryStart, geometryEnd)
              << ',' << geometryQueries << '\n';
    std::cout << "BENCHMARK,C++,XLPP,tracked_change_count," << elapsedMs(trackingStart, trackingEnd)
              << ',' << trackingQueries << '\n';
    std::cout << "BENCHMARK,C++,XLPP,save_store," << elapsedMs(saveStart, saveEnd)
              << ',' << bytes << '\n';
    std::cout << "VERIFY," << sheet.maxRow() << ',' << sheet.maxColumn() << ',' << dimensions
              << ",tracked=" << trackedChangeTotal / trackingQueries << '\n';
    return 0;
}
