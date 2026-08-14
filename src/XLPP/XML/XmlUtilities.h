#pragma once
#include <string>
#include <string_view>
#include <vector>
#include <sstream>

namespace xlpp::internal {

// Return the '>' that terminates an opening tag while respecting quoted
// attribute values. A plain find('>') is incorrect for valid XML such as
// <node note="1 > 0">.
inline const char* xmlOpeningTagEnd(const char* p, const char* end) noexcept {
    char quote = 0;
    for (; p < end; ++p) {
        const char ch = *p;
        if (quote != 0) {
            if (ch == quote) quote = 0;
            continue;
        }
        if (ch == '\'' || ch == '"') { quote = ch; continue; }
        if (ch == '>') return p;
    }
    return end;
}

inline bool xmlNameBoundary(char ch) noexcept {
    return ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n' || ch == '>' || ch == '/';
}

inline bool xmlTagNameMatches(const char* p, const char* end, std::string_view name, bool closing) noexcept {
    if (p >= end || *p != '<') return false;
    const std::size_t prefix = closing ? 2u : 1u;
    if (static_cast<std::size_t>(end - p) <= prefix + name.size()) return false;
    if (closing) {
        if (p[1] != '/') return false;
    } else if (p[1] == '/' || p[1] == '!' || p[1] == '?') {
        return false;
    }
    const char* q = p + prefix;
    for (std::size_t i = 0; i < name.size(); ++i)
        if (q[i] != name[i]) return false;
    return xmlNameBoundary(q[name.size()]);
}

inline const char* xmlFindSequence(const char* p, const char* end, std::string_view needle) noexcept {
    if (needle.empty()) return p;
    for (; p < end; ++p) {
        if (static_cast<std::size_t>(end - p) < needle.size()) return end;
        bool match = true;
        for (std::size_t i = 0; i < needle.size(); ++i) {
            if (p[i] != needle[i]) { match = false; break; }
        }
        if (match) return p;
    }
    return end;
}

// If p starts non-element XML markup, return the byte after it. nullptr means
// p begins normal element markup. Unterminated markup returns end so callers
// can treat the document as truncated rather than scanning markup payload as
// real elements.
inline const char* xmlSkipNonElementMarkup(const char* p, const char* end) noexcept {
    if (p >= end || *p != '<' || p + 1 >= end) return nullptr;
    if (p[1] == '?') {
        const char* close = xmlFindSequence(p + 2, end, "?>");
        return close == end ? end : close + 2;
    }
    if (p[1] != '!') return nullptr;
    if (static_cast<std::size_t>(end - p) >= 4 && p[2] == '-' && p[3] == '-') {
        const char* close = xmlFindSequence(p + 4, end, "-->");
        return close == end ? end : close + 3;
    }
    constexpr std::string_view cdata = "<![CDATA[";
    if (static_cast<std::size_t>(end - p) >= cdata.size()) {
        bool match = true;
        for (std::size_t i = 0; i < cdata.size(); ++i)
            if (p[i] != cdata[i]) { match = false; break; }
        if (match) {
            const char* close = xmlFindSequence(p + cdata.size(), end, "]]>");
            return close == end ? end : close + 3;
        }
    }
    // DOCTYPE/other declarations. Respect quotes and the internal subset so a
    // '>' inside [ ... ] cannot terminate the declaration early.
    char quote = 0;
    unsigned bracketDepth = 0;
    for (const char* q = p + 2; q < end; ++q) {
        const char ch = *q;
        if (quote != 0) {
            if (ch == quote) quote = 0;
            continue;
        }
        if (ch == '\'' || ch == '"') { quote = ch; continue; }
        if (ch == '[') { ++bracketDepth; continue; }
        if (ch == ']' && bracketDepth != 0) { --bracketDepth; continue; }
        if (ch == '>' && bracketDepth == 0) return q + 1;
    }
    return end;
}

struct XmlElementBounds {
    const char* openingEnd{nullptr};
    const char* contentBegin{nullptr};
    const char* contentEnd{nullptr};
    const char* elementEnd{nullptr};
    bool selfClosing{false};
};

inline bool xmlElementBounds(const char* start, const char* end, std::string_view name,
                             XmlElementBounds& bounds) noexcept {
    if (!xmlTagNameMatches(start, end, name, false)) return false;
    const char* openingEnd = xmlOpeningTagEnd(start + 1 + name.size(), end);
    if (openingEnd == end) return false;

    const char* beforeGt = openingEnd;
    while (beforeGt > start && (beforeGt[-1] == ' ' || beforeGt[-1] == '\t' ||
                                beforeGt[-1] == '\r' || beforeGt[-1] == '\n'))
        --beforeGt;
    if (beforeGt > start && beforeGt[-1] == '/') {
        bounds = {openingEnd, openingEnd, openingEnd, openingEnd + 1, true};
        return true;
    }

    std::size_t depth = 1;
    const char* scan = openingEnd + 1;
    while (scan < end) {
        while (scan < end && *scan != '<') ++scan;
        if (scan >= end) return false;
        if (scan + 1 >= end) return false;
        const char marker = scan[1];
        // Most child markup is neither a close/nested tag for `name` nor
        // special XML markup. Collapse that common case into one branch.
        if (marker != '/' && marker != '!' && marker != '?' &&
            (name.empty() || marker != name.front())) {
            ++scan;
            continue;
        }
        if (marker == '/' && scan + 2 < end && !name.empty() && scan[2] == name.front() &&
            xmlTagNameMatches(scan, end, name, true)) {
            // XML end-tags cannot have attributes. Avoid the quote-aware opening
            // scanner on this very hot path; only optional whitespace may
            // separate the name from '>'.
            const char* closeGt = scan + 2 + name.size();
            while (closeGt < end && (*closeGt == ' ' || *closeGt == '\t' ||
                                     *closeGt == '\r' || *closeGt == '\n'))
                ++closeGt;
            if (closeGt >= end || *closeGt != '>') return false;
            if (--depth == 0) {
                bounds = {openingEnd, openingEnd + 1, scan, closeGt + 1, false};
                return true;
            }
            scan = closeGt + 1;
            continue;
        }
        if (marker != '/' && marker != '!' && marker != '?' && !name.empty() &&
            marker == name.front() && xmlTagNameMatches(scan, end, name, false)) {
            const char* nestedGt = xmlOpeningTagEnd(scan + 1 + name.size(), end);
            if (nestedGt == end) return false;
            const char* beforeNestedGt = nestedGt;
            while (beforeNestedGt > scan && (beforeNestedGt[-1] == ' ' || beforeNestedGt[-1] == '\t' ||
                                              beforeNestedGt[-1] == '\r' || beforeNestedGt[-1] == '\n'))
                --beforeNestedGt;
            if (!(beforeNestedGt > scan && beforeNestedGt[-1] == '/')) ++depth;
            scan = nestedGt + 1;
            continue;
        }
        if (marker == '!' || marker == '?') {
            const char* next = xmlSkipNonElementMarkup(scan, end);
            if (next == nullptr || next == end) return false;
            scan = next;
            continue;
        }
        // For non-target elements, advance to continue scanning descendants.
        // XML forbids a literal '<' inside attribute values, so no quote-state
        // work is needed on this fast path.
        ++scan;
    }
    return false;
}

std::string xmlEscape(std::string_view value);
std::string xmlUnescape(std::string_view value);

// Streaming variant: writes the XML-escaped form of `s` directly into `out`
// without constructing an intermediate std::string. Used by the hot save path
// to avoid per-cell temporary allocations.
void writeXmlEscaped(std::ostringstream& out, std::string_view s);

// Zero-allocation attribute extractor (uses SIMD scan internally).
// Returns a view into `tag` or empty view when absent. Does not decode entities.
std::string_view attrView(std::string_view tag, std::string_view name);

// Convenience: entity-decoded attribute value as std::string.
std::string attribute(std::string_view tag, std::string_view name);

std::vector<std::string> tags(std::string_view xml, std::string_view tagName);
std::string tagText(std::string_view xml, std::string_view tagName);

// Zero-allocation tag enumeration: calls fn(string_view) per matching element.
// The view aliases `xml`; do not retain after xml goes out of scope.
template<typename Fn>
void tagsForEach(std::string_view xml, std::string_view tagName, Fn&& fn) {
    if (tagName.empty()) return;
    const char* const end = xml.data() + xml.size();
    const char* p = xml.data();
    while (p < end) {
        while (p < end && *p != '<') ++p;
        if (p >= end) return;
        if (p + 1 < end && (p[1] == '!' || p[1] == '?')) {
            const char* next = xmlSkipNonElementMarkup(p, end);
            if (next == nullptr || next == end) return;
            p = next;
            continue;
        }
        XmlElementBounds bounds;
        if (p + 1 < end && p[1] == tagName.front() && xmlElementBounds(p, end, tagName, bounds)) {
            fn(std::string_view(p, static_cast<std::size_t>(bounds.elementEnd - p)));
            p = bounds.elementEnd;
            continue;
        }
        ++p;
    }
}

// Zero-allocation text content extractor. Returns a view into `xml` or empty.
// Does NOT entity-decode.
inline std::string_view tagTextView(std::string_view xml, std::string_view tagName) {
    if (tagName.empty()) return {};
    const char* const end = xml.data() + xml.size();
    const char* p = xml.data();
    while (p < end) {
        while (p < end && *p != '<') ++p;
        if (p >= end) return {};
        if (p + 1 < end && (p[1] == '!' || p[1] == '?')) {
            const char* next = xmlSkipNonElementMarkup(p, end);
            if (next == nullptr || next == end) return {};
            p = next;
            continue;
        }
        XmlElementBounds bounds;
        if (p + 1 < end && p[1] == tagName.front() && xmlElementBounds(p, end, tagName, bounds)) {
            if (bounds.selfClosing) return {};
            return std::string_view(bounds.contentBegin,
                                    static_cast<std::size_t>(bounds.contentEnd - bounds.contentBegin));
        }
        ++p;
    }
    return {};
}

} // namespace xlpp::internal
