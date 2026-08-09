#pragma once
#include "Formula/CalculationValue.h"

#include <optional>
#include <string>
#include <vector>

namespace xlpp::internal::formula {

class FormulaEvaluationContext {
public:
    virtual ~FormulaEvaluationContext() = default;
    virtual EvalValue resolveReference(Worksheet& context, const std::string& sheetName,
                                       const std::string& first, const std::optional<std::string>& last,
                                       std::size_t depth) = 0;
    virtual EvalValue resolveExternalReference(const std::string& workbookToken, const std::string& sheetName,
                                               const std::string& first, const std::optional<std::string>& last) = 0;
    virtual EvalValue resolveDefinedName(Worksheet& context, const std::string& name, std::size_t depth) = 0;
    virtual EvalValue resolveStructuredReference(Worksheet& context, const std::string& expression, std::size_t depth) = 0;
    virtual EvalValue callFunction(std::string name, const std::vector<EvalValue>& args,
                                   Worksheet& context, std::size_t depth) = 0;
    virtual bool compare(const Scalar& lhs, const Scalar& rhs, const std::string& op) const = 0;
    virtual bool evaluateVolatileFunctions() const noexcept = 0;
    virtual EvalValue unsupportedFunction(const std::string& name) = 0;
    virtual bool date1904() const noexcept = 0;
};

EvalValue parseFormula(FormulaEvaluationContext& context, Worksheet& sheet,
                       std::string formula, std::size_t depth);

} // namespace xlpp::internal::formula
