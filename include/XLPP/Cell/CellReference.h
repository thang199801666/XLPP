#pragma once
#include <cctype>
#include <cstddef>
#include <limits>
#include <stdexcept>
#include <string>
#include <string_view>
#include <charconv>

namespace xlpp {

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
inline constexpr std::uint64_t makeCellKey(std::size_t row, std::size_t column) noexcept {
    return (static_cast<std::uint64_t>(row) << 20) | column;
}

struct CellReference {
    std::size_t row{1};
    std::size_t column{1};

    static std::string columnName(std::size_t column) {
        if (column == 0) throw std::invalid_argument("Column index is 1-based");
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
        return value;
    }

    static CellReference parse(std::string_view address) {
        if (address.empty()) throw std::invalid_argument("Cell address cannot be empty");
        std::size_t columnValue = 0;
        std::size_t digitStart = 0;
        bool readingRow = false;
        bool hasColumn = false;

        for (std::size_t i = 0; i < address.size(); ++i) {
            const unsigned char ch = static_cast<unsigned char>(address[i]);
            if (ch == '$') continue;
            if (!readingRow) {
                const auto d = detail::base26Lookup(ch);
                if (d != 0) {
                    hasColumn = true;
                    if (columnValue > (std::numeric_limits<std::size_t>::max() - d) / 26)
                        throw std::overflow_error("Column index overflow");
                    columnValue = columnValue * 26 + d;
                    continue;
                }
            }
            if (ch >= '0' && ch <= '9') {
                if (!readingRow) { readingRow = true; digitStart = i; }
            } else {
                throw std::invalid_argument("Invalid cell address: " + std::string(address));
            }
        }
        if (!hasColumn || !readingRow)
            throw std::invalid_argument("Invalid cell address: " + std::string(address));
        std::size_t parsedRow = 0;
        const auto* data = address.data() + digitStart;
        const auto* end = address.data() + address.size();
        const auto result = std::from_chars(data, end, parsedRow);
        if (result.ec != std::errc{} || parsedRow == 0)
            throw std::invalid_argument("Invalid cell address: " + std::string(address));
        return {parsedRow, columnValue};
    }

    std::string address() const { return columnName(column) + std::to_string(row); }
};
}
