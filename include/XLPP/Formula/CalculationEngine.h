#pragma once
#include <XLPP/Cell/Cell.h>
#include <XLPP/Formula/Calculation.h>
#include <unordered_map>
#include <string>
#include <vector>

namespace xlpp {
class Workbook;
class Worksheet;

namespace internal {

// Recursive-descent formula parser/evaluator used by Workbook::calculate().
// Supports A1 references (sheet-qualified and external-book), ranges, the
// standard arithmetic/comparison/concat operators and a practical set of
// worksheet functions. Each formula cell's cached value is updated in place;
// the formula text is never modified.

struct CalculationEvaluator {
    const xlpp::Workbook& workbook;
    const std::string& sheetName;
    const CalculationOptions& options;
    // "sheet\x1Fcell" keys of formula cells on the current evaluation stack.
    std::vector<std::string>& stack;
    std::unordered_map<std::string, CellValue>& memo;
    std::size_t depth{0};
    std::size_t* unsupportedFormulas{nullptr};
    std::size_t* externalResolved{nullptr};
    std::size_t* unresolvedExternal{nullptr};
    std::size_t* circularReferences{nullptr};
};

// Evaluates `formula` (without the leading '=') within `sheetName`.
CellValue evaluateFormula(const CalculationEvaluator& ctx, const std::string& formula);

// Resolves a cell reference in the given worksheet, evaluating dependent
// formulas recursively. Used by the evaluator and exposed for tests.
CellValue resolveReference(const CalculationEvaluator& ctx,
                           const std::string& sheetName,
                           const std::string& cellAddress);

} // namespace internal
} // namespace xlpp
