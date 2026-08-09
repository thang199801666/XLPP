#pragma once

#include "Formula/FormulaEvaluationContext.h"

#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace xlpp::internal::formula {

std::vector<Scalar> flattened(const std::vector<EvalValue>& args);
bool wildcardMatchInsensitive(std::string_view text, std::string_view pattern);

struct FormulaFunctionCall {
    FormulaEvaluationContext& engine;
    const std::vector<EvalValue>& args;
    Worksheet& worksheet;
    std::size_t depth;

    std::optional<CellError> error() const;
    Scalar scalarArg(std::size_t index) const;
    std::optional<double> numArg(std::size_t index) const;
    std::optional<bool> boolArg(std::size_t index) const;
    std::string textArg(std::size_t index) const;
    std::vector<double> numericValues() const;
    bool matchesCriterion(const Scalar& candidate, const Scalar& criterionValue) const;
};

using FormulaFunctionResult = std::optional<EvalValue>;

FormulaFunctionResult evaluateMathStatFinancialFunctions(FormulaFunctionCall& call, const std::string& name);
FormulaFunctionResult evaluateLogicalTextDateFunctions(FormulaFunctionCall& call, const std::string& name);
FormulaFunctionResult evaluateCriteriaLookupFunctions(FormulaFunctionCall& call, const std::string& name);
FormulaFunctionResult evaluateDynamicReferenceFunctions(FormulaFunctionCall& call, const std::string& name);

} // namespace xlpp::internal::formula
