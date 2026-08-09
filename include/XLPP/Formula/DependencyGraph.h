#pragma once
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

class Workbook;

enum class FormulaDependencyKind {
    CellOrRange,
    DefinedName,
    Table,
    ExternalReference,
    VolatileReference
};

struct FormulaDependency {
    std::string dependentSheet;
    std::string dependentCell;
    FormulaDependencyKind kind{FormulaDependencyKind::CellOrRange};
    std::string precedentSheet;
    std::string precedentReference;
    std::string symbol;
};

struct FormulaDependencyReport {
    std::size_t formulaCells{0};
    std::size_t edges{0};
    std::size_t cellOrRangeEdges{0};
    std::size_t definedNameEdges{0};
    std::size_t tableEdges{0};
    std::size_t externalEdges{0};
    std::size_t volatileReferences{0};
    std::size_t unresolvedSymbols{0};
    std::vector<std::string> warnings;
};

class FormulaDependencyGraph {
public:
    const std::vector<FormulaDependency>& edges() const noexcept { return edges_; }
    const FormulaDependencyReport& report() const noexcept { return report_; }

    std::vector<FormulaDependency> precedentsOf(const std::string& sheet, const std::string& cell) const;
    std::vector<FormulaDependency> dependentsOf(const std::string& sheet, const std::string& cell) const;
    bool dependsOn(const std::string& dependentSheet, const std::string& dependentCell,
                   const std::string& precedentSheet, const std::string& precedentCell) const;

private:
    friend FormulaDependencyGraph buildFormulaDependencyGraph(const Workbook& workbook);
    std::vector<FormulaDependency> edges_;
    FormulaDependencyReport report_;
};

FormulaDependencyGraph buildFormulaDependencyGraph(const Workbook& workbook);

} // namespace xlpp
