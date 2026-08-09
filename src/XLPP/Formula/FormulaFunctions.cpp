#include "Formula/FormulaFunctions.h"
#include "Formula/FunctionRegistry.h"

namespace xlpp::internal::formula {

EvalValue evaluateFormulaFunction(FormulaEvaluationContext& engine, std::string name,
                                  const std::vector<EvalValue>& args, Worksheet& worksheet,
                                  std::size_t depth) {
    name = normalizeFunctionName(std::move(name));
    FormulaFunctionCall call{engine, args, worksheet, depth};
    if (auto result = FunctionRegistry::defaultRegistry().evaluate(call, name)) return std::move(*result);
    return engine.unsupportedFunction(name);
}

} // namespace xlpp::internal::formula
