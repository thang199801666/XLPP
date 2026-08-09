#pragma once
#include "Formula/FormulaEvaluationContext.h"

namespace xlpp::internal::formula {

EvalValue evaluateFormulaFunction(FormulaEvaluationContext& engine, std::string name,
                                  const std::vector<EvalValue>& args, Worksheet& worksheet,
                                  std::size_t depth);

} // namespace xlpp::internal::formula
