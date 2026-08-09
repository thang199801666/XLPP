#include "TestFramework.h"
#include <algorithm>
#include <cstdlib>
#include <stdexcept>
#include <iostream>
#include <optional>
#include <string>

namespace {
std::optional<std::size_t> envIndex(const char* name) {
    const char* value = std::getenv(name);
    if (!value || !*value) return std::nullopt;
    try {
        const auto parsed = std::stoull(value);
        if (parsed == 0) throw std::invalid_argument("test indices are one-based");
        return static_cast<std::size_t>(parsed);
    } catch (...) {
        std::cerr << "Invalid " << name << "='" << value << "' (expected one-based integer)\n";
        std::exit(2);
    }
}
}

int main() {
    std::cout << std::unitbuf;
    std::vector<TestCase> tests;
    tests.reserve(256);
    registerModelWorkbookTests(tests);
    registerRegressionTests(tests);
    registerFormulaDependencyTests(tests);

    std::cout << "============================================================\n";
    std::cout << " XL++ Unit Tests - P0Z-H VBA Authoring + Preservation\n";
    std::cout << "============================================================\n";

    const auto beginEnv = envIndex("XLPP_TEST_BEGIN");
    const auto endEnv = envIndex("XLPP_TEST_END");
    const std::size_t first = beginEnv ? std::min(*beginEnv - 1, tests.size()) : 0;
    const std::size_t last = endEnv ? std::min(*endEnv, tests.size()) : tests.size();
    if (first > last) {
        std::cerr << "XLPP_TEST_BEGIN must not be greater than XLPP_TEST_END\n";
        return 2;
    }

    TestContext context;
    std::size_t passed = 0;
    const std::size_t selected = last - first;
    for (std::size_t index = first; index < last; ++index) {
        std::cout << "\n[RUN  " << index + 1 << '/' << tests.size() << "] " << tests[index].first << '\n';
        try {
            tests[index].second(context);
            ++passed;
            std::cout << "[PASS " << index + 1 << '/' << tests.size() << "] " << tests[index].first << '\n';
        } catch (const std::exception& error) {
            std::cerr << "[FAIL " << index + 1 << '/' << tests.size() << "] " << tests[index].first
                      << "\n    Reason: " << error.what() << '\n';
        } catch (...) {
            std::cerr << "[FAIL " << index + 1 << '/' << tests.size() << "] " << tests[index].first
                      << "\n    Reason: unknown exception\n";
        }
    }
    std::cout << "\n============================================================\n";
    std::cout << " Test suites passed : " << passed << '/' << selected << '\n';
    std::cout << " Checks executed    : " << context.checks() << '\n';
    std::cout << " Final result       : " << (passed == selected ? "PASS" : "FAIL") << '\n';
    std::cout << "============================================================\n";
    return passed == selected ? 0 : 1;
}
