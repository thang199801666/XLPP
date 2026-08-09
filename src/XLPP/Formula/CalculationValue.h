#pragma once
#include <XLPP/Cell/Cell.h>
#include <XLPP/Worksheet/Worksheet.h>

#include <algorithm>
#include <cctype>
#include <cstdlib>
#include <cmath>
#include <iomanip>
#include <optional>
#include <sstream>
#include <string>
#include <vector>
#include <utility>

namespace xlpp::internal::formula {

using Scalar = CellValue;

struct EvalValue {
    Scalar scalar{};
    std::vector<Scalar> range;
    bool isRange{false};
    std::size_t rows{1};
    std::size_t columns{1};
    Worksheet* referenceSheet{nullptr};
    std::size_t referenceRow{0};
    std::size_t referenceColumn{0};

    static EvalValue fromScalar(Scalar value) { EvalValue out; out.scalar = std::move(value); return out; }
    static EvalValue fromRange(std::vector<Scalar> values, std::size_t rowCount = 1, std::size_t columnCount = 0) {
        EvalValue out;
        out.range = std::move(values);
        out.isRange = true;
        out.rows = rowCount;
        out.columns = columnCount ? columnCount : (rowCount ? out.range.size() / rowCount : 0);
        return out;
    }
    EvalValue& withReference(Worksheet* sheet, std::size_t row, std::size_t column) noexcept {
        referenceSheet = sheet; referenceRow = row; referenceColumn = column; return *this;
    }
    bool hasReference() const noexcept { return referenceSheet != nullptr && referenceRow > 0 && referenceColumn > 0; }
};

inline bool isError(const Scalar& value) { return std::holds_alternative<CellError>(value); }
inline CellError errorOr(const Scalar& value, CellError fallback = CellError::Value) {
    if (const auto* error = std::get_if<CellError>(&value)) return *error;
    return fallback;
}

inline std::string upperAscii(std::string value) {
    std::transform(value.begin(), value.end(), value.begin(), [](unsigned char c) {
        return static_cast<char>(std::toupper(c));
    });
    return value;
}

inline std::string trimAscii(std::string value) {
    auto first = std::find_if_not(value.begin(), value.end(), [](unsigned char c) { return std::isspace(c) != 0; });
    auto last = std::find_if_not(value.rbegin(), value.rend(), [](unsigned char c) { return std::isspace(c) != 0; }).base();
    if (first >= last) return {};
    return std::string(first, last);
}

inline std::optional<double> parseNumberText(const std::string& text) {
    if (text.empty()) return std::nullopt;
    char* end = nullptr;
    const double value = std::strtod(text.c_str(), &end);
    if (end == text.c_str()) return std::nullopt;
    while (*end && std::isspace(static_cast<unsigned char>(*end))) ++end;
    if (*end != '\0' || !std::isfinite(value)) return std::nullopt;
    return value;
}

inline std::optional<double> numberValue(const Scalar& value, bool date1904) {
    if (const auto* v = std::get_if<double>(&value)) return *v;
    if (const auto* v = std::get_if<bool>(&value)) return *v ? 1.0 : 0.0;
    if (const auto* v = std::get_if<DateTime>(&value)) return toExcelSerial(*v, date1904);
    if (std::holds_alternative<std::monostate>(value)) return 0.0;
    if (const auto* v = std::get_if<std::string>(&value)) return parseNumberText(*v);
    return std::nullopt;
}

inline std::optional<bool> boolValue(const Scalar& value, bool date1904) {
    if (const auto* v = std::get_if<bool>(&value)) return *v;
    if (const auto n = numberValue(value, date1904)) return *n != 0.0;
    if (const auto* v = std::get_if<std::string>(&value)) {
        const auto u = upperAscii(trimAscii(*v));
        if (u == "TRUE") return true;
        if (u == "FALSE" || u.empty()) return false;
    }
    if (std::holds_alternative<std::monostate>(value)) return false;
    return std::nullopt;
}

inline std::string scalarText(const Scalar& value, bool date1904) {
    if (const auto* v = std::get_if<std::string>(&value)) return *v;
    if (const auto* v = std::get_if<bool>(&value)) return *v ? "TRUE" : "FALSE";
    if (const auto* v = std::get_if<double>(&value)) {
        std::ostringstream out; out << std::setprecision(15) << *v; return out.str();
    }
    if (const auto* v = std::get_if<DateTime>(&value)) {
        std::ostringstream out; out << std::setprecision(15) << toExcelSerial(*v, date1904); return out.str();
    }
    if (const auto* v = std::get_if<CellError>(&value)) return toString(*v);
    return {};
}

inline Scalar firstScalar(const EvalValue& value) {
    if (!value.isRange) return value.scalar;
    if (value.range.empty()) return std::monostate{};
    return value.range.front();
}

} // namespace xlpp::internal::formula
