#include <XLPP/XLPP.h>
#include <filesystem>
#include <iostream>

int main(int argc, char** argv) {
    try {
        const auto output = argc > 1 ? std::filesystem::path(argv[1])
                                     : std::filesystem::path("xlpp_p0x_sample.xlsx");
        xlpp::Workbook workbook;
        auto& sheet = workbook.addWorksheet("Data");
        sheet.append({std::string("Item"), std::string("Amount")});
        sheet.append({std::string("A"), 10.0});
        sheet.append({std::string("B"), 20.0});
        sheet.cell("D1").setFormula("SUM(B2:B3)");

        // Workbook-level structural editing rewrites dependent formulas and
        // workbook objects transactionally.
        workbook.insertRows("Data", 2, 1);
        sheet.cell("A2").setValue("Inserted");
        sheet.cell("B2").setValue(5.0);

        auto calculation = workbook.calculateFormulas();
        workbook.save(output);
        std::cout << "Saved " << output << " with calculated D1="
                  << sheet.cell("D1").numericValueOr(0.0)
                  << ", formula evaluations=" << calculation.formulaCellsEvaluated << '\n';
        return 0;
    } catch (const std::exception& error) {
        std::cerr << error.what() << '\n';
        return 1;
    }
}
