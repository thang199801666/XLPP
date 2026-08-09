#pragma once
#include <string>
#include <string_view>
#include <vector>
#include <sstream>

namespace xlpp::internal {
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
    using namespace std::string_view_literals;
    const std::size_t tnLen = tagName.size();
    const char* const data = xml.data();
    const char* const endPtr = data + xml.size();
    const char* p = data;

    for (;;) {
        while (p < endPtr && *p != '<') ++p;
        if (p >= endPtr) return;
        if (static_cast<std::size_t>(endPtr - p) < tnLen + 1) return;

        if (p[1] != '/') {
            bool nameMatch = true;
            for (std::size_t j = 0; j < tnLen; ++j) {
                if (p[1 + j] != tagName[j]) { nameMatch = false; break; }
            }
            if (nameMatch) {
                const char* after = p + 1 + tnLen;
                if (after < endPtr) {
                    char b = *after;
                    if (b == ' ' || b == '\t' || b == '\r' || b == '\n' || b == '>' || b == '/') {
                        const char* gt = after;
                        for (; gt < endPtr && *gt != '>'; ++gt) {}
                        if (gt >= endPtr) return;

                        if (gt > after && *(gt - 1) == '/') {
                            fn(std::string_view(p, static_cast<std::size_t>(gt - p + 1)));
                            p = gt + 1;
                            continue;
                        }

                        const char* findStart = gt + 1;
                        for (;;) {
                            while (findStart < endPtr && !(findStart[0] == '<' && findStart + 1 < endPtr && findStart[1] == '/')) ++findStart;
                            if (findStart >= endPtr) { p = endPtr; break; }
                            if (static_cast<std::size_t>(endPtr - findStart) < tnLen + 3) { p = endPtr; break; }

                            bool closeMatch = true;
                            for (std::size_t j = 0; j < tnLen; ++j) {
                                if (findStart[2 + j] != tagName[j]) { closeMatch = false; break; }
                            }
                            if (closeMatch && findStart[2 + tnLen] == '>') {
                                fn(std::string_view(p, static_cast<std::size_t>(findStart + tnLen + 3 - p)));
                                p = findStart + tnLen + 3;
                                break;
                            }
                            ++findStart;
                        }
                        continue;
                    }
                }
            }
        }
        ++p;
    }
}

// Zero-allocation text content extractor. Returns a view into `xml` or empty.
// Does NOT entity-decode.
inline std::string_view tagTextView(std::string_view xml, std::string_view tagName) {
    const std::size_t tnLen = tagName.size();
    const char* const data = xml.data();
    const char* const endPtr = data + xml.size();
    const char* p = data;

    for (;;) {
        while (p < endPtr && *p != '<') ++p;
        if (p >= endPtr) return {};
        if (static_cast<std::size_t>(endPtr - p) < tnLen + 1) return {};

        if (p[1] != '/') {
            bool nameMatch = true;
            for (std::size_t j = 0; j < tnLen; ++j) {
                if (p[1 + j] != tagName[j]) { nameMatch = false; break; }
            }
            if (nameMatch) {
                const char* after = p + 1 + tnLen;
                if (after < endPtr) {
                    char b = *after;
                    if (b == ' ' || b == '\t' || b == '\r' || b == '\n' || b == '>' || b == '/') {
                        const char* gt = after;
                        for (; gt < endPtr && *gt != '>'; ++gt) {}
                        if (gt >= endPtr) return {};
                        if (gt > after && *(gt - 1) == '/') return {};

                        const char* findStart = gt + 1;
                        for (;;) {
                            while (findStart < endPtr && !(findStart[0] == '<' && findStart + 1 < endPtr && findStart[1] == '/')) ++findStart;
                            if (findStart >= endPtr) return {};
                            if (static_cast<std::size_t>(endPtr - findStart) < tnLen + 3) return {};

                            bool closeMatch = true;
                            for (std::size_t j = 0; j < tnLen; ++j) {
                                if (findStart[2 + j] != tagName[j]) { closeMatch = false; break; }
                            }
                            if (closeMatch && findStart[2 + tnLen] == '>') {
                                return std::string_view(gt + 1, static_cast<std::size_t>(findStart - (gt + 1)));
                            }
                            ++findStart;
                        }
                    }
                }
            }
        }
        ++p;
    }
}

} // namespace xlpp::internal
