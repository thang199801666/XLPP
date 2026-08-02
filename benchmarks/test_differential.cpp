// Test differential save: load, modify one sheet, save
#include <XLPP/XLPP.h>
#include <iostream>
#include <chrono>
#include <filesystem>

using namespace xlpp;
using Clock = std::chrono::high_resolution_clock;
using Ms = std::chrono::duration<double, std::milli>;
static double ms(Clock::duration d) { return Ms(d).count(); }

int main() {
    auto tmp = std::filesystem::temp_directory_path() / "differential_test.xlsx";

    // Create workbook with 2 sheets, each 50K rows
    {
        Workbook wb;
        auto& s1 = wb.addWorksheet("Sheet1");
        for (int r = 1; r <= 50000; ++r) {
            s1.append({std::string("A") + std::to_string(r), double(r), std::string("data")});
        }
        auto& s2 = wb.addWorksheet("Sheet2");
        for (int r = 1; r <= 50000; ++r) {
            s2.append({std::string("B") + std::to_string(r), double(r), std::string("info")});
        }
        wb.save(tmp);
        std::cout << "Initial save: 100K rows\n";
    }

    // Load, modify ONLY Sheet2, save — Sheet1 should be skipped
    {
        Workbook wb;
        wb.load(tmp);

        // Verify both sheets are clean after load
        std::cout << "After load: Sheet1 dirty=" << wb.worksheet("Sheet1")->dirty()
                  << " Sheet2 dirty=" << wb.worksheet("Sheet2")->dirty() << "\n";

        // Save 1: nothing modified (all clean) — should reuse cache
        auto t0 = Clock::now();
        wb.save(tmp);
        auto t1 = Clock::now();
        std::cout << "Save 1 (no changes):   " << ms(t1 - t0) << " ms\n";

        // Save 2: modify Sheet2 only
        wb.worksheet("Sheet2")->cell("A1").setValue("MODIFIED");
        auto t2 = Clock::now();
        wb.save(tmp);
        auto t3 = Clock::now();
        std::cout << "Save 2 (1 cell change): " << ms(t3 - t2) << " ms\n";
        std::cout << "  Sheet1 dirty=" << wb.worksheet("Sheet1")->dirty()
                  << " Sheet2 dirty=" << wb.worksheet("Sheet2")->dirty() << "\n";
    }

    std::filesystem::remove(tmp);
    return 0;
}
