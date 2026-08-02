#pragma once
#include <charconv>
#include <cstddef>
#include <system_error>
#include <string_view>
#include "SimdScan.h"

namespace xlpp::internal {

// Allocation-free XML helpers for parsing in-memory XML fragments (whole
// elements, start tags, attribute values, text content) into string_views.
// Every view returned here aliases the caller's buffer and stays valid for as
// long as that buffer does. Values are NOT entity-decoded; call containsEntity()
// and xmlUnescape() only when decoded text is required.

namespace xmlscan_detail {
constexpr std::size_t npos = std::string_view::npos;

inline bool isNameChar(char c) noexcept {
    return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') ||
           c == '_' || c == '-' || c == ':' || c == '.';
}

// Index of '<' of an opening <name... tag found at or after `from`, or npos.
// Uses SIMD first-char filter for the name search.
inline std::size_t findOpen(std::string_view xml, std::string_view name, std::size_t from) noexcept {
    return simd::findOpenTag(xml, name, from);
}

// Index of '<' of a </name> close tag found at or after `from`, or npos.
// Uses SIMD first-char filter for the name search.
inline std::size_t findClose(std::string_view xml, std::string_view name, std::size_t from) noexcept {
    return simd::findCloseTag(xml, name, from);
}

// Index of the '>' that closes the start tag beginning at the '<' at `from`,
// skipping quoted attribute values, or npos.
// Uses SIMD dual-byte (">" | '"') bitmap scan in 16-byte chunks.
inline std::size_t scanGt(std::string_view xml, std::size_t from) noexcept {
    return simd::scanGt(xml, from);
}
} // namespace xmlscan_detail

// Iterates every complete element named `name` inside an in-memory fragment,
// self-closing tags included. Yields each element as a string_view that aliases
// the scanned buffer; no copies are made.
class XmlScanner {
public:
    explicit XmlScanner(std::string_view xml) noexcept : xml_(xml) {}

    void rewind() noexcept { pos_ = 0; }

    bool nextElement(std::string_view name, std::string_view& out) noexcept {
        const std::size_t n = name.size();
        for (;;) {
            const std::size_t p = xmlscan_detail::findOpen(xml_, name, pos_);
            if (p == xmlscan_detail::npos) return false;
            const std::size_t gt = xmlscan_detail::scanGt(xml_, p + 1 + n);
            if (gt == xmlscan_detail::npos) return false;
            if (gt > p + 1 && xml_[gt - 1] == '/') {
                out = xml_.substr(p, gt - p + 1);
                pos_ = gt + 1;
                return true;
            }
            const std::size_t c = xmlscan_detail::findClose(xml_, name, gt + 1);
            if (c == xmlscan_detail::npos) return false;
            out = xml_.substr(p, c + n + 3 - p);
            pos_ = c + n + 3;
            return true;
        }
    }

private:
    std::string_view xml_;
    std::size_t pos_{0};
};

// Value of attribute `name` in a start tag as a string_view into `tag`, or an
// empty view when the attribute is absent. Does not decode entities.
// Uses SIMD first-char filter for the attribute name search.
inline std::string_view xmlAttribute(std::string_view tag, std::string_view name) noexcept {
    return simd::xmlAttribute(tag, name);
}

// Text content of the first element named `name` inside `element`, as a
// string_view into `element`, or an empty view when absent. Self-closing and
// empty elements yield an empty view. Does not decode entities.
inline std::string_view xmlText(std::string_view element, std::string_view name) noexcept {
    const std::size_t p = xmlscan_detail::findOpen(element, name, 0);
    if (p == xmlscan_detail::npos) return {};
    const std::size_t gt = xmlscan_detail::scanGt(element, p + 1 + name.size());
    if (gt == xmlscan_detail::npos) return {};
    if (gt > p && element[gt - 1] == '/') return {};
    const std::size_t c = xmlscan_detail::findClose(element, name, gt + 1);
    if (c == xmlscan_detail::npos) return {};
    return element.substr(gt + 1, c - (gt + 1));
}

inline bool containsEntity(std::string_view s) noexcept { return s.find('&') != xmlscan_detail::npos; }

// Locale-independent, allocation-free numeric parsing (std::from_chars). Each
// parser requires the whole view to be consumed on success.
inline bool parseDouble(std::string_view s, double& out) noexcept {
    if (s.empty()) return false;
    const auto result = std::from_chars(s.data(), s.data() + s.size(), out);
    return result.ec == std::errc{} && result.ptr == s.data() + s.size();
}
inline bool parseSize(std::string_view s, std::size_t& out) noexcept {
    if (s.empty()) return false;
    const auto result = std::from_chars(s.data(), s.data() + s.size(), out);
    return result.ec == std::errc{} && result.ptr == s.data() + s.size();
}
inline bool parseInt(std::string_view s, int& out) noexcept {
    if (s.empty()) return false;
    const auto result = std::from_chars(s.data(), s.data() + s.size(), out);
    return result.ec == std::errc{} && result.ptr == s.data() + s.size();
}

} // namespace xlpp::internal
