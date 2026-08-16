#include <XLPP/XLPP.h>
#include <iostream>

static void trySingle(const char* formula) {
    xlpp::Workbook wb;
    auto& s = wb.addWorksheet("S1");
    s.cell("A1").setValue(10.0);
    s.cell("A2").setValue(5.0);
    s.cell("X1").setFormula(formula);
    const auto rep = wb.calculate();
    const auto* cell = s.tryCell("X1");
    std::cout << "  [" << formula << "] -> ";
    if (!cell) { std::cout << "missing\n"; return; }
    if (const auto* d = std::get_if<double>(&cell->value())) std::cout << *d;
    else if (const auto* str = std::get_if<std::string>(&cell->value())) std::cout << *str;
    else if (const auto* b = std::get_if<bool>(&cell->value())) std::cout << (*b ? "true" : "false");
    else if (const auto* e = std::get_if<xlpp::CellError>(&cell->value())) std::cout << xlpp::toString(*e);
    else std::cout << "(empty)";
    std::cout << " (err=" << rep.evaluationErrors << ")\n";
}

int main() {
    trySingle("=A1+A2");
    trySingle("=SUM(A1:A2)");
    trySingle("=IF(B1>10,\"big\",\"small\")");
    trySingle("=IF(A1>10,\"big\",\"small\")");
    trySingle("=A1&\"-\"&A2");
    trySingle("=ROUND(A1/7,2)");
    trySingle("=ROUND(A1,2)");
    trySingle("=MIN(A1,A2,3)");
    std::cout << "--- full model ---\n";
    xlpp::Workbook wb;
    auto& s1 = wb.addWorksheet("Sheet1");
    auto& s2 = wb.addWorksheet("Sheet2");
    s1.cell("A1").setValue(10.0);
    s1.cell("A2").setValue(5.0);
    s1.cell("B1").setFormula("=A1+A2");
    s1.cell("B2").setFormula("=SUM(A1:A2)");
    s1.cell("B3").setFormula("=IF(B1>10,\"big\",\"small\")");
    s1.cell("B4").setFormula("=A1&\"-\"&A2");
    s1.cell("B5").setFormula("=ROUND(A1/7,2)");
    s1.cell("B6").setFormula("='Sheet2'!B2");
    s1.cell("B7").setFormula("=1/0");
    s1.cell("B8").setFormula("=B8+1");
    s1.cell("B9").setFormula("=MIN(A1,A2,3)");
    s2.cell("B2").setValue(100.0);
    std::cout << "calling calculate" << std::endl;
    const auto report = wb.calculate();
    std::cout << "done evaluated=" << report.formulaCellsEvaluated << " errors=" << report.evaluationErrors
              << " circular=" << report.circularReferences << std::endl;
    for (const char* addr : {"B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9"}) {
        const auto* cell = s1.tryCell(addr);
        std::cout << addr << ": ";
        if (!cell) { std::cout << "missing\n"; continue; }
        if (const auto* d = std::get_if<double>(&cell->value())) std::cout << *d;
        else if (const auto* str = std::get_if<std::string>(&cell->value())) std::cout << *str;
        else if (const auto* b = std::get_if<bool>(&cell->value())) std::cout << (*b ? "true" : "false");
        else if (const auto* e = std::get_if<xlpp::CellError>(&cell->value())) std::cout << xlpp::toString(*e);
        else std::cout << "(empty)";
        std::cout << "\n";
    }
    return 0;
}
