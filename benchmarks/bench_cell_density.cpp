// XL++ Cell Density Benchmark
//
// Measures construction/save cost for default cells versus densely styled
// cells. This guards lazy model storage against regressions in both the common
// path and the deliberately adversarial 100%-styled path.
#include <XLPP/XLPP.h>

#include <chrono>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <iomanip>
#include <iostream>
#include <string>

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
                             sizeof(counters)))
        return static_cast<std::uint64_t>(counters.PeakWorkingSetSize);
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
    constexpr std::size_t fallback = 20000;
    if (argc < 2) return fallback;
    try {
        const auto value = std::stoull(argv[1]);
        return value == 0 ? fallback : static_cast<std::size_t>(value);
    } catch (...) {
        return fallback;
    }
}

bool parseStyled(int argc, char** argv) {
    if (argc < 3) return false;
    const std::string mode = argv[2];
    return mode == "styled" || mode == "1" || mode == "true";
}

} // namespace

int main(int argc, char** argv) {
    const std::size_t rows = parseRows(argc, argv);
    constexpr std::size_t columns = 10;
    const bool styled = parseStyled(argc, argv);

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Density");
    xlpp::Style denseStyle;
    if (styled) {
        denseStyle.font().setBold(true);
        denseStyle.fill().setPatternType("solid");
        denseStyle.fill().foregroundColor().setArgb("FFD9EAF7");
        denseStyle.alignment().setHorizontal("center");
        denseStyle.setNumberFormat("#,##0.00");
    }

    const auto buildStart = Clock::now();
    for (std::size_t r = 1; r <= rows; ++r) {
        for (std::size_t c = 1; c <= columns; ++c) {
            auto& cell = sheet.cell(r, c);
            cell.setValue(static_cast<double>(r * c) / 3.0);
            if (styled) cell.style() = denseStyle;
        }
    }
    const auto buildEnd = Clock::now();

    xlpp::SaveOptions options;
    options.compressionLevel = xlpp::CompressionLevel::Store;
    const auto output = std::filesystem::temp_directory_path() /
        (styled ? "xlpp_cell_density_styled.xlsx" : "xlpp_cell_density_default.xlsx");
    const auto saveStart = Clock::now();
    workbook.save(output, options);
    const auto saveEnd = Clock::now();
    const auto bytes = std::filesystem::file_size(output);
    std::filesystem::remove(output);

    std::cout << std::fixed << std::setprecision(3);
    std::cout << "METRIC,C++,XLPP,sizeof_cell_bytes," << sizeof(xlpp::Cell) << '\n';
    std::cout << "METRIC,C++,XLPP,peak_rss_bytes," << peakResidentBytes() << '\n';
    std::cout << "BENCHMARK,C++,XLPP,cell_density_build," << elapsedMs(buildStart, buildEnd)
              << ',' << rows * columns << ',' << (styled ? "styled" : "default") << '\n';
    std::cout << "BENCHMARK,C++,XLPP,cell_density_save," << elapsedMs(saveStart, saveEnd)
              << ',' << bytes << ',' << (styled ? "styled" : "default") << '\n';
    return 0;
}
