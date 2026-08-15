#pragma once
#include <cctype>
#include <stdexcept>
#include <string>
#include <string_view>

namespace xlpp::internal {

inline std::size_t utf8CodePointCount(std::string_view value) {
    std::size_t count = 0;
    for (std::size_t i = 0; i < value.size();) {
        const unsigned char c = static_cast<unsigned char>(value[i]);
        std::size_t width = 0;
        unsigned int scalar = 0;
        if (c < 0x80) { width = 1; scalar = c; }
        else if ((c & 0xE0) == 0xC0) { width = 2; scalar = c & 0x1F; }
        else if ((c & 0xF0) == 0xE0) { width = 3; scalar = c & 0x0F; }
        else if ((c & 0xF8) == 0xF0) { width = 4; scalar = c & 0x07; }
        else throw std::invalid_argument("Worksheet name contains invalid UTF-8");
        if (i + width > value.size()) throw std::invalid_argument("Worksheet name contains truncated UTF-8");
        for (std::size_t j = 1; j < width; ++j) {
            const unsigned char cc = static_cast<unsigned char>(value[i + j]);
            if ((cc & 0xC0) != 0x80) throw std::invalid_argument("Worksheet name contains invalid UTF-8 continuation byte");
            scalar = (scalar << 6) | (cc & 0x3F);
        }
        if ((width == 2 && scalar < 0x80) || (width == 3 && scalar < 0x800) ||
            (width == 4 && scalar < 0x10000) || scalar > 0x10FFFF ||
            (scalar >= 0xD800 && scalar <= 0xDFFF))
            throw std::invalid_argument("Worksheet name contains invalid UTF-8 scalar value");
        ++count;
        i += width;
    }
    return count;
}

inline void validateWorksheetName(std::string_view name) {
    if (name.empty()) throw std::invalid_argument("Worksheet name cannot be empty");
    if (utf8CodePointCount(name) > 31) throw std::invalid_argument("Worksheet name cannot exceed 31 Unicode characters");
    if (name.front() == '\'' || name.back() == '\'')
        throw std::invalid_argument("Worksheet name cannot begin or end with an apostrophe");
    for (const unsigned char c : name) {
        if (c < 0x20 || c == 0x7F)
            throw std::invalid_argument("Worksheet name cannot contain control characters");
        switch (c) {
        case ':': case '\\': case '/': case '?': case '*': case '[': case ']':
            throw std::invalid_argument("Worksheet name contains a character forbidden by Excel");
        default: break;
        }
    }
}

inline bool worksheetNamesEquivalent(std::string_view left, std::string_view right) noexcept {
    if (left.size() != right.size()) return false;
    for (std::size_t i = 0; i < left.size(); ++i) {
        const unsigned char a = static_cast<unsigned char>(left[i]);
        const unsigned char b = static_cast<unsigned char>(right[i]);
        if (a < 0x80 && b < 0x80) {
            if (std::tolower(a) != std::tolower(b)) return false;
        } else if (a != b) return false;
    }
    return true;
}

} // namespace xlpp::internal
