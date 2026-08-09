#pragma once
#include <cctype>
#include <cstddef>
#include <stdexcept>
#include <string>
#include <string_view>

namespace xlpp {

inline constexpr std::size_t MaxWorksheetNameCharacters = 31;

namespace detail {
inline bool decodeWorksheetNameCodePoint(std::string_view text, std::size_t& offset,
                                         char32_t& codePoint) noexcept {
    if (offset >= text.size()) return false;
    const auto first = static_cast<unsigned char>(text[offset++]);
    if (first < 0x80) { codePoint = first; return true; }
    unsigned extra = 0;
    char32_t value = 0;
    char32_t minimum = 0;
    if ((first & 0xE0u) == 0xC0u) { extra = 1; value = first & 0x1Fu; minimum = 0x80; }
    else if ((first & 0xF0u) == 0xE0u) { extra = 2; value = first & 0x0Fu; minimum = 0x800; }
    else if ((first & 0xF8u) == 0xF0u) { extra = 3; value = first & 0x07u; minimum = 0x10000; }
    else return false;
    if (offset + extra > text.size()) return false;
    for (unsigned i = 0; i < extra; ++i) {
        const auto continuation = static_cast<unsigned char>(text[offset++]);
        if ((continuation & 0xC0u) != 0x80u) return false;
        value = (value << 6u) | (continuation & 0x3Fu);
    }
    if (value < minimum || value > 0x10FFFFu || (value >= 0xD800u && value <= 0xDFFFu)) return false;
    codePoint = value;
    return true;
}
} // namespace detail

inline bool isValidWorksheetName(std::string_view name) noexcept {
    if (name.empty()) return false;
    std::size_t offset = 0;
    std::size_t characters = 0;
    while (offset < name.size()) {
        char32_t cp = 0;
        if (!detail::decodeWorksheetNameCodePoint(name, offset, cp)) return false;
        if (++characters > MaxWorksheetNameCharacters) return false;
        if (cp < 0x20) return false;
        if (cp < 0x80) {
            switch (static_cast<char>(cp)) {
                case ':': case '\\': case '/': case '?': case '*': case '[': case ']': return false;
                default: break;
            }
        }
    }
    return characters != 0;
}

inline void validateWorksheetName(std::string_view name) {
    if (!isValidWorksheetName(name))
        throw std::invalid_argument("Invalid worksheet name (must be valid UTF-8, <=31 characters, and exclude : \\ / ? * [ ])");
}

// Excel sheet identifiers are case-insensitive. This fast comparison performs
// ASCII case folding while preserving exact UTF-8 bytes for non-ASCII code
// points; it covers the overwhelmingly common OOXML identifier set without
// introducing a Unicode database dependency into the core.
inline bool worksheetNamesEqual(std::string_view a, std::string_view b) noexcept {
    if (a.size() != b.size()) return false;
    for (std::size_t i = 0; i < a.size(); ++i) {
        const auto ca = static_cast<unsigned char>(a[i]);
        const auto cb = static_cast<unsigned char>(b[i]);
        if (ca < 0x80 && cb < 0x80) {
            if (std::tolower(ca) != std::tolower(cb)) return false;
        } else if (ca != cb) return false;
    }
    return true;
}

} // namespace xlpp
