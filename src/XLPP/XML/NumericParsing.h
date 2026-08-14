#pragma once

#include <charconv>
#include <cmath>
#include <cstddef>
#include <stdexcept>
#include <string>
#include <string_view>
#include <system_error>
#include <type_traits>

namespace xlpp::internal {

inline std::string_view trimAsciiWhitespace(std::string_view value) noexcept {
    while (!value.empty() && (value.front() == ' ' || value.front() == '\t' ||
                              value.front() == '\r' || value.front() == '\n'))
        value.remove_prefix(1);
    while (!value.empty() && (value.back() == ' ' || value.back() == '\t' ||
                              value.back() == '\r' || value.back() == '\n'))
        value.remove_suffix(1);
    return value;
}

template <class Integer>
bool tryParseIntegerExact(std::string_view text, Integer& value) noexcept {
    static_assert(std::is_integral_v<Integer>, "Integer parser requires an integral type");
    text = trimAsciiWhitespace(text);
    if (text.empty()) return false;
    // std::from_chars intentionally does not accept a leading '+' for integer
    // overloads. XML Schema integer lexical forms do, so accept it explicitly.
    if (text.front() == '+') {
        text.remove_prefix(1);
        if (text.empty()) return false;
    }
    Integer parsed{};
    const auto result = std::from_chars(text.data(), text.data() + text.size(), parsed, 10);
    if (result.ec != std::errc{} || result.ptr != text.data() + text.size()) return false;
    value = parsed;
    return true;
}

template <class Integer>
Integer parseIntegerExact(std::string_view text, std::string_view fieldName) {
    Integer value{};
    if (!tryParseIntegerExact(text, value))
        throw std::runtime_error("Invalid integer value for " + std::string(fieldName) + ": '" + std::string(text) + "'");
    return value;
}

inline bool tryParseDoubleExact(std::string_view text, double& value) noexcept {
    text = trimAsciiWhitespace(text);
    if (text.empty()) return false;
    if (text.front() == '+') {
        text.remove_prefix(1);
        if (text.empty()) return false;
    }
    double parsed = 0.0;
    const auto result = std::from_chars(text.data(), text.data() + text.size(), parsed, std::chars_format::general);
    if (result.ec != std::errc{} || result.ptr != text.data() + text.size() || !std::isfinite(parsed)) return false;
    value = parsed;
    return true;
}

inline double parseDoubleExact(std::string_view text, std::string_view fieldName) {
    double value = 0.0;
    if (!tryParseDoubleExact(text, value))
        throw std::runtime_error("Invalid finite numeric value for " + std::string(fieldName) + ": '" + std::string(text) + "'");
    return value;
}

} // namespace xlpp::internal
