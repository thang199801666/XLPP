#include "Formula/FunctionRegistry.h"
#include "Formula/FormulaFunctionSupport.h"

namespace xlpp::internal::formula {

void FunctionRegistry::registerFamily(FunctionDescriptor descriptor, Evaluator evaluator) {
    entries_.push_back({std::move(descriptor), std::move(evaluator)});
}

std::optional<EvalValue> FunctionRegistry::evaluate(FormulaFunctionCall& call, const std::string& normalizedName) const {
    for (const auto& entry : entries_) {
        if (auto value = entry.evaluator(call, normalizedName)) return value;
    }
    return std::nullopt;
}

const FunctionRegistry& FunctionRegistry::defaultRegistry() {
    static const FunctionRegistry registry = [] {
        FunctionRegistry result;
        result.registerFamily({"MathStatFinancial", false, false}, evaluateMathStatFinancialFunctions);
        result.registerFamily({"LogicalTextDate", true, false}, evaluateLogicalTextDateFunctions);
        result.registerFamily({"CriteriaLookup", false, false}, evaluateCriteriaLookupFunctions);
        result.registerFamily({"DynamicReference", true, true}, evaluateDynamicReferenceFunctions);
        return result;
    }();
    return registry;
}

std::string normalizeFunctionName(std::string name) {
    while (name.rfind("_xlfn.", 0) == 0) name.erase(0, 6);
    while (name.rfind("_xlws.", 0) == 0) name.erase(0, 6);
    return upperAscii(std::move(name));
}

} // namespace xlpp::internal::formula
