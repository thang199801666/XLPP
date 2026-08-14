#include <XLPP/XLPP.h>

#include <filesystem>
#include <iostream>
#include <string>

namespace {
int fail(const std::string& message) {
    std::cerr << "P1P regression failure: " << message << '\n';
    return 1;
}
}

int main() {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_p1p_lazy_formula_metadata.xlsx";
    std::error_code ignored;
    std::filesystem::remove(path, ignored);

    xlpp::Workbook source;
    auto& sheet = source.addWorksheet("FormulaDensity");
    for (std::size_t row = 1; row <= 5000; ++row)
        sheet.cell(row, 1).setFormula("1+1");
    sheet.cell("B1").setSharedFormula("A1+1", 3, "B1:B2");
    sheet.cell("B2").setArrayFormula("A1:A2*2", "B2:B3");
    source.save(path);

    xlpp::Workbook loaded;
    loaded.load(path);
    auto* loadedSheet = loaded.worksheet("FormulaDensity");
    if (!loadedSheet) return fail("worksheet was not loaded");

    for (std::size_t row = 1; row <= 5000; ++row) {
        const auto& cell = loadedSheet->cell(row, 1);
        if (cell.hasFormulaMetadata())
            return fail("ordinary formula unexpectedly materialized FormulaMetadata at row " + std::to_string(row));
    }

    const auto& shared = loadedSheet->cell("B1");
    if (!shared.hasFormulaMetadata() || shared.formulaMetadata().type() != xlpp::FormulaType::Shared ||
        shared.formulaMetadata().sharedIndex().value_or(0u) != 3u)
        return fail("shared formula metadata did not round-trip");

    const auto& arrayFormula = loadedSheet->cell("B2");
    if (!arrayFormula.hasFormulaMetadata() || arrayFormula.formulaMetadata().type() != xlpp::FormulaType::Array)
        return fail("array formula metadata did not round-trip");

    auto& rewritten = loadedSheet->cell("B1");
    rewritten.setFormula("2+2");
    if (rewritten.hasFormulaMetadata())
        return fail("normal formula retained stale shared metadata");

    std::filesystem::remove(path, ignored);
    std::cout << "P1P lazy formula metadata regression OK\n";
    return 0;
}
