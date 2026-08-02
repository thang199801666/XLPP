// Test Excel 365 dynamic array formulas
#include <XLPP/XLPP.h>
#include <iostream>
#include <filesystem>
#include <cassert>

using namespace xlpp;

int main() {
    // xlfn() prefix helper
    assert(xlfn("SORT") == "_xlfn.SORT");
    assert(xlfn("_xlfn.FILTER") == "_xlfn.FILTER");
    std::cout << "xlfn(): OK\n";

    auto tmp = std::filesystem::temp_directory_path() / "xlpp_dynarray.xlsx";

    {
        Workbook wb;
        auto& ws = wb.addWorksheet("DA");

        for (int r = 1; r <= 10; ++r)
            ws.cell(r, 1).setValue(static_cast<double>(r));

        // Dynamic array formula with spill range
        auto& f = ws.cell("B1");
        f.setDynamicArrayFormula(xlfn("SORT") + "(A1:A10)", "B1:B10");
        assert(f.formulaMetadata().type() == FormulaType::DynamicArray);
        assert(f.formulaMetadata().alwaysCalculateArray());

        // Spill reference
        ws.cell("C1").setFormula("=B1#");

        wb.save(tmp);
        std::cout << "Saved dynamic array workbook\n";
    }

    {
        Workbook wb;
        wb.load(tmp);
        auto& ws = *wb.worksheet("DA");
        const auto& f = ws.cell("B1");
        std::cout << "B1 formula: " << f.formula() << "\n";
        std::cout << "B1 type: " << static_cast<int>(f.formulaMetadata().type())
                  << " (3=DynamicArray)\n";
        std::cout << "B1 ref: " << f.formulaMetadata().reference() << "\n";
        std::cout << "B1 aca: " << f.formulaMetadata().alwaysCalculateArray() << "\n";
        std::cout << "C1 formula: " << ws.cell("C1").formula() << "\n";

        assert(f.formula() == "_xlfn.SORT(A1:A10)");
        assert(f.formulaMetadata().type() == FormulaType::DynamicArray);
        assert(f.formulaMetadata().reference() == "B1:B10");
        assert(f.formulaMetadata().alwaysCalculateArray());
        assert(ws.cell("C1").formula() == "=B1#");
        std::cout << "Round-trip OK\n";
    }
    std::filesystem::remove(tmp);
    return 0;
}
