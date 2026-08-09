#pragma once
#include <XLPP/Cell/Cell.h>
#include <functional>
#include <optional>

#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

struct CalculationCell {
    std::string sheet;
    std::string cell;
};

// Controls XL++'s in-process formula calculation engine. The engine updates
// each formula cell's cached value; the formula text itself is never changed.
struct CalculationOptions {
    // Recalculate formula dependencies recursively before evaluating callers.
    bool recursiveDependencies{true};
    // Update the cached value stored in formula cells. When false, calculation
    // is diagnostic-only and the workbook remains unchanged.
    bool updateCachedValues{true};
    // Evaluate volatile NOW()/TODAY() functions. If false they are reported as
    // unsupported so deterministic callers can opt out explicitly.
    bool evaluateVolatileFunctions{true};
    // Materialize multi-cell dynamic-array results into their spill range.
    bool spillDynamicArrays{true};
    // Resolve circular numeric models by fixed-point iteration, using the prior
    // iteration's cached values for back-edges in the dependency graph.
    bool iterativeCalculation{false};
    // Maximum fixed-point passes when iterative calculation is enabled.
    std::size_t maxIterations{100};
    // Numeric convergence tolerance, matching Excel's maximum-change concept.
    double maxChange{0.001};
    // Resolve a cell from an external workbook reference such as
    // '[Book.xlsx]Sheet 1'!A1. The callback receives workbook token, sheet name
    // and canonical A1 address. Returning nullopt leaves the reference unresolved.
    std::function<std::optional<CellValue>(const std::string&, const std::string&, const std::string&)> externalReferenceResolver{};
    // Maximum recursive formula/reference depth before reporting #NUM!.
    std::size_t maxDepth{512};
    // Optional dirty roots for dependency-driven partial recalculation. When
    // empty, every formula cell is visited. When populated, XL++ recalculates
    // only formulas transitively dependent on these changed cells plus formula
    // roots themselves when a dirty root is a formula cell.
    std::vector<CalculationCell> changedCells{};
};

struct CalculationReport {
    std::size_t formulaCellsVisited{0};
    std::size_t formulaCellsEvaluated{0};
    std::size_t cachedValuesUpdated{0};
    std::size_t dependencyEvaluations{0};
    std::size_t definedNamesResolved{0};
    std::size_t circularReferences{0};
    std::size_t unsupportedFormulas{0};
    std::size_t evaluationErrors{0};
    std::size_t dynamicArraysSpilled{0};
    std::size_t spillCellsUpdated{0};
    std::size_t spillConflicts{0};
    std::size_t structuredReferencesResolved{0};
    std::size_t iterativeIterations{0};
    std::size_t iterativeConvergenceFailures{0};
    std::size_t externalReferencesResolved{0};
    std::size_t unresolvedExternalReferences{0};
    std::size_t dirtyRoots{0};
    std::size_t dirtyFormulaCellsSelected{0};
    std::vector<std::string> warnings;

    bool success() const noexcept {
        return circularReferences == 0 && unsupportedFormulas == 0 && evaluationErrors == 0 && spillConflicts == 0
            && iterativeConvergenceFailures == 0 && unresolvedExternalReferences == 0;
    }
};

} // namespace xlpp
