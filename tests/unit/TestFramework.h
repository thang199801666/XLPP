#pragma once

#include <XLPP/XLPP.h>
#include <cmath>
#include <functional>
#include <iostream>
#include <sstream>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

class TestContext {
public:
    template <typename Actual, typename Expected>
    void checkEqual(const Actual& actual, const Expected& expected, const std::string& label) {
        ++checks_;
        if (actual == expected) {
            std::cout << "    [CHECK PASS] " << label << " | actual=" << printable(actual)
                      << " expected=" << printable(expected) << '\n';
            return;
        }
        std::ostringstream message;
        message << label << " | actual=" << printable(actual)
                << " expected=" << printable(expected);
        throw std::runtime_error(message.str());
    }

    void checkNear(double actual, double expected, double tolerance, const std::string& label) {
        ++checks_;
        if (std::abs(actual - expected) <= tolerance) {
            std::cout << "    [CHECK PASS] " << label << " | actual=" << actual
                      << " expected=" << expected << " tolerance=" << tolerance << '\n';
            return;
        }
        std::ostringstream message;
        message << label << " | actual=" << actual << " expected=" << expected
                << " tolerance=" << tolerance;
        throw std::runtime_error(message.str());
    }

    void checkTrue(bool condition, const std::string& label) {
        ++checks_;
        if (condition) {
            std::cout << "    [CHECK PASS] " << label << " | actual=true expected=true\n";
            return;
        }
        throw std::runtime_error(label + " | actual=false expected=true");
    }

    std::size_t checks() const noexcept { return checks_; }

private:
    template <typename T>
    static std::string printable(const T& value) {
        std::ostringstream stream;
        stream << value;
        return stream.str();
    }

    static std::string printable(const std::string& value) { return '"' + value + '"'; }
    static std::string printable(const char* value) { return printable(std::string(value)); }
    static std::string printable(bool value) { return value ? "true" : "false"; }
    static std::string printable(const xlpp::DateTime& value) { return xlpp::toIso8601(value); }

    std::size_t checks_{0};
};

using TestFunction = std::function<void(TestContext&)>;

using TestCase = std::pair<std::string, TestFunction>;

void registerModelWorkbookTests(std::vector<TestCase>& tests);
void registerFormulaDependencyTests(std::vector<TestCase>& tests);
void registerRegressionTests(std::vector<TestCase>& tests);
