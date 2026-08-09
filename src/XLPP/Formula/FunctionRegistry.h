#pragma once

#include "Formula/FormulaFunctionSupport.h"

#include <functional>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace xlpp::internal::formula {

struct FunctionDescriptor {
    std::string family;
    bool volatileFamily{false};
    bool dynamicArrayFamily{false};
};

class FunctionRegistry {
public:
    using Evaluator = std::function<std::optional<EvalValue>(FormulaFunctionCall&, const std::string&)>;

    struct Entry {
        FunctionDescriptor descriptor;
        Evaluator evaluator;
    };

    void registerFamily(FunctionDescriptor descriptor, Evaluator evaluator);
    std::optional<EvalValue> evaluate(FormulaFunctionCall& call, const std::string& normalizedName) const;
    const std::vector<Entry>& entries() const noexcept { return entries_; }

    static const FunctionRegistry& defaultRegistry();

private:
    std::vector<Entry> entries_;
};

std::string normalizeFunctionName(std::string name);

} // namespace xlpp::internal::formula
