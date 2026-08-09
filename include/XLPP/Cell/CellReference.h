#pragma once
#include <cctype>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <stdexcept>
#include <string>
#include <string_view>
#include <charconv>

namespace xlpp {

inline constexpr std::size_t MaxExcelRows = 1'048'576;
inline constexpr std::size_t MaxExcelColumns = 16'384;

inline constexpr bool isValidCellCoordinate(std::size_t row, std::size_t column) noexcept {
    return row >= 1 && row <= MaxExcelRows && column >= 1 && column <= MaxExcelColumns;
}

namespace detail {
// Precomputed powers of 26: 26^0, 26^1, 26^2, ...
// Max Excel column is XFD (16,384) which fits in 3 chars.
// 26^3 = 17,576 > 16,384, so 3 digits max.
constexpr std::size_t pow26(std::size_t n) noexcept {
    std::size_t v = 1;
    for (std::size_t i = 0; i < n; ++i) v *= 26;
    return v;
}

// Fast base-26 digit lookup: maps 'A'-'Z' and 'a'-'z' to 1-26, 0 for invalid.
inline constexpr unsigned char base26Lookup(unsigned char c) noexcept {
    if (c >= 'A' && c <= 'Z') return static_cast<unsigned char>(c - 'A' + 1);
    if (c >= 'a' && c <= 'z') return static_cast<unsigned char>(c - 'a' + 1);
    return 0;
}
} // namespace detail

// Compact cell key: (row << 20) | column. Max row 1,048,576 (20 bits), max col 16,384 (14 bits).
// std::map with this key yields row-major ordering naturally.
inline constexpr std::uint64_t makeCellKey(std::size_t row, std::size_t column) {
    if (!isValidCellCoordinate(row, column))
        throw std::out_of_range("Cell coordinate exceeds Excel worksheet bounds");
    return (static_cast<std::uint64_t>(row) << 20) | column;
}

struct CellReference {
    std::size_t row{1};
    std::size_t column{1};

    static std::string columnName(std::size_t column) {
        if (column == 0) throw std::invalid_argument("Column index is 1-based");
        if (column > MaxExcelColumns) throw std::out_of_range("Column index exceeds Excel's 16,384-column limit");
        char buffer[16];
        int pos = 15;
        buffer[pos] = '\0';
        while (column > 0) {
            const auto remainder = (column - 1) % 26;
            buffer[--pos] = static_cast<char>('A' + remainder);
            column = (column - 1) / 26;
        }
        return std::string(buffer + pos);
    }

    static std::size_t columnIndex(std::string_view name) {
        if (name.empty()) throw std::invalid_argument("Column name cannot be empty");
        std::size_t value = 0;
        for (const char raw : name) {
            const auto digit = detail::base26Lookup(static_cast<unsigned char>(raw));
            if (digit == 0) throw std::invalid_argument("Invalid column name");
            if (value > (std::numeric_limits<std::size_t>::max() - digit) / 26)
                throw std::overflow_error("Column index overflow");
            value = value * 26 + digit;
        }
        if (value > MaxExcelColumns) throw std::out_of_range("Column index exceeds Excel's 16,384-column limit");
        return value;
    }

    static CellReference parse(std::string_view address) {
        if (address.empty()) throw std::invalid_argument("Cell address cannot be empty");

        const auto invalid = [&]() -> std::invalid_argument {
            return std::invalid_argument("Invalid cell address: " + std::string(address));
        };

        std::size_t i = 0;
        if (address[i] == '$') {
            ++i;
            if (i == address.size()) throw invalid();
        }

        std::size_t columnValue = 0;
        const std::size_t columnStart = i;
        while (i < address.size()) {
            const auto digit = detail::base26Lookup(static_cast<unsigned char>(address[i]));
            if (digit == 0) break;
            if (columnValue > (std::numeric_limits<std::size_t>::max() - digit) / 26)
                throw std::overflow_error("Column index overflow");
            columnValue = columnValue * 26 + digit;
            ++i;
        }
        if (i == columnStart) throw invalid();

        if (i < address.size() && address[i] == '$') ++i;
        const std::size_t rowStart = i;
        if (rowStart == address.size()) throw invalid();
        while (i < address.size() && address[i] >= '0' && address[i] <= '9') ++i;
        if (i != address.size() || i == rowStart) throw invalid();

        std::size_t parsedRow = 0;
        const auto* data = address.data() + rowStart;
        const auto* end = address.data() + address.size();
        const auto result = std::from_chars(data, end, parsedRow);
        if (result.ec != std::errc{} || result.ptr != end || parsedRow == 0)
            throw invalid();
        if (parsedRow > MaxExcelRows || columnValue > MaxExcelColumns)
            throw std::out_of_range("Cell address exceeds Excel worksheet bounds: " + std::string(address));
        return {parsedRow, columnValue};
    }

    std::string address() const { return columnName(column) + std::to_string(row); }
};
}
