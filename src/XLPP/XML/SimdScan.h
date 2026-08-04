#pragma once
#include <cstddef>
#include <cstdint>
#include <string_view>

#ifdef _MSC_VER
#include <intrin.h>
#endif

// SSE2 intrinsics only exist on x86-64 targets; the header errors out on
// ARM64 (Apple Silicon) and is meaningless elsewhere. SIMD code paths are
// additionally gated on _M_AMD64 (MSVC) so non-MSVC x86-64 builds fall back
// to the scalar implementations below.
#if defined(_M_AMD64) || defined(__x86_64__)
#include <emmintrin.h> // SSE2 baseline (guaranteed on x64)
#endif

namespace xlpp::internal::simd {

// --- memchr replacement: find byte in [begin, end), 16 bytes/iter ---

inline const char* findByte(const char* p, const char* end, unsigned char target) noexcept {
#ifdef _M_AMD64
    const __m128i vt = _mm_set1_epi8(static_cast<char>(target));
    for (; p + 16 <= end; p += 16) {
        const __m128i chunk = _mm_loadu_si128(reinterpret_cast<const __m128i*>(p));
        int mask = _mm_movemask_epi8(_mm_cmpeq_epi8(chunk, vt));
        if (mask) return p + _tzcnt_u32(static_cast<unsigned>(mask));
    }
#endif
    for (; p < end; ++p) if (*p == static_cast<char>(target)) return p;
    return nullptr;
}

// --- Find first byte matching any of two candidates (a | b) ---

inline const char* findByteOr(const char* p, const char* end,
                              unsigned char ca, unsigned char cb) noexcept {
#ifdef _M_AMD64
    const __m128i va = _mm_set1_epi8(static_cast<char>(ca));
    const __m128i vb = _mm_set1_epi8(static_cast<char>(cb));
    for (; p + 16 <= end; p += 16) {
        const __m128i chunk = _mm_loadu_si128(reinterpret_cast<const __m128i*>(p));
        const __m128i ma = _mm_cmpeq_epi8(chunk, va);
        const __m128i mb = _mm_cmpeq_epi8(chunk, vb);
        int mask = _mm_movemask_epi8(_mm_or_si128(ma, mb));
        if (mask) return p + _tzcnt_u32(static_cast<unsigned>(mask));
    }
#endif
    for (; p < end; ++p) if (*p == static_cast<char>(ca) || *p == static_cast<char>(cb)) return p;
    return nullptr;
}

// --- Quoted scan: find '>' while skipping quoted "..." sections ---
// Processes 16-byte chunks with SIMD bitmaps; only falls back to scalar
// for bytes where bitmap says a special char ('>' or '"') exists.

inline std::size_t scanGt(std::string_view xml, std::size_t from) noexcept {
    const char* p = xml.data() + from;
    const char* end = xml.data() + xml.size();
    bool quoted = false;

#ifdef _M_AMD64
    const __m128i vGt = _mm_set1_epi8('>');
    const __m128i vQ = _mm_set1_epi8('"');
    for (; p + 16 <= end; ) {
        __m128i chunk = _mm_loadu_si128(reinterpret_cast<const __m128i*>(p));
        int mGt = _mm_movemask_epi8(_mm_cmpeq_epi8(chunk, vGt));
        int mQ  = _mm_movemask_epi8(_mm_cmpeq_epi8(chunk, vQ));
        unsigned bits = static_cast<unsigned>(mGt) | static_cast<unsigned>(mQ);
        if (!bits) { p += 16; continue; }
        while (bits) {
            int idx = _tzcnt_u32(bits);
            char c = p[idx];
            if (quoted) {
                if (c == '"') quoted = false;
            } else if (c == '"') {
                quoted = true;
            } else if (c == '>') {
                return static_cast<std::size_t>(p + idx - xml.data());
            }
            bits &= bits - 1;
        }
        p += 16;
    }
#endif

    for (; p < end; ++p) {
        if (quoted) { if (*p == '"') quoted = false; }
        else if (*p == '"') quoted = true;
        else if (*p == '>') return static_cast<std::size_t>(p - xml.data());
    }
    return std::string_view::npos;
}

// --- SIMD substring search: find `needle` in `haystack[start..]` ---
// First-character SIMD filter + scalar verify. Faster than std::string_view::find
// for multi-char needles by using 16-byte-chunk bitmap scans.

inline std::size_t findStr(std::string_view haystack, std::string_view needle,
                           std::size_t from) noexcept {
    const std::size_t n = needle.size();
    if (n == 0) return std::string_view::npos;
    if (from >= haystack.size()) return std::string_view::npos;

    if (n == 1) {
        const char* found = findByte(haystack.data() + from, haystack.data() + haystack.size(),
                                     static_cast<unsigned char>(needle[0]));
        return found ? static_cast<std::size_t>(found - haystack.data()) : std::string_view::npos;
    }

    const char* p = haystack.data() + from;
    const char* end = haystack.data() + haystack.size();
    const char* limit = end - n + 1;

#ifdef _M_AMD64
    const __m128i vFirst = _mm_set1_epi8(needle[0]);
    for (; p + 16 <= limit; p += 16) {
        __m128i chunk = _mm_loadu_si128(reinterpret_cast<const __m128i*>(p));
        int mask = _mm_movemask_epi8(_mm_cmpeq_epi8(chunk, vFirst));
        while (mask) {
            int idx = _tzcnt_u32(static_cast<unsigned>(mask));
            const char* candidate = p + idx;
            if (candidate + n <= end) {
                bool match = true;
                for (std::size_t j = 1; j < n; ++j) {
                    if (candidate[j] != needle[j]) { match = false; break; }
                }
                if (match) return static_cast<std::size_t>(candidate - haystack.data());
            }
            mask &= mask - 1;
        }
    }
#endif

    for (; p < limit; ++p) {
        if (*p == needle[0]) {
            bool match = true;
            for (std::size_t j = 1; j < n; ++j) {
                if (p[j] != needle[j]) { match = false; break; }
            }
            if (match) return static_cast<std::size_t>(p - haystack.data());
        }
    }
    return std::string_view::npos;
}

// --- SIMD open-tag scan: find `name` preceded by '<' and followed by boundary ---
// Returns the index of '<' (not the first char of name), matching the
// xmlscan_detail::findOpen convention.

inline std::size_t findOpenTag(std::string_view xml, std::string_view name,
                               std::size_t from) noexcept {
    const std::size_t n = name.size();
    if (n == 0 || from >= xml.size()) return std::string_view::npos;
    if (from + n + 1 > xml.size()) return std::string_view::npos;

    const char* p = xml.data() + from;
    const char* end = xml.data() + xml.size();
    const char* limit = end - n;

#ifdef _M_AMD64
    const __m128i vFirst = _mm_set1_epi8(name[0]);
    for (; p + 16 <= limit; p += 16) {
        __m128i chunk = _mm_loadu_si128(reinterpret_cast<const __m128i*>(p));
        int mask = _mm_movemask_epi8(_mm_cmpeq_epi8(chunk, vFirst));
        while (mask) {
            int idx = _tzcnt_u32(static_cast<unsigned>(mask));
            const char* candidate = p + idx;
            if (candidate > xml.data() && candidate[-1] == '<') {
                const char* after = candidate + n;
                if (after < end) {
                    bool match = true;
                    for (std::size_t j = 1; j < n; ++j) {
                        if (candidate[j] != name[j]) { match = false; break; }
                    }
                    if (match) {
                        char boundary = *after;
                        if (boundary == ' ' || boundary == '\t' || boundary == '\r' ||
                            boundary == '\n' || boundary == '>' || boundary == '/') {
                            return static_cast<std::size_t>(candidate - 1 - xml.data());
                        }
                    }
                }
            }
            mask &= mask - 1;
        }
    }
#endif

    for (; p <= limit; ++p) {
        if (*p == name[0] && p > xml.data() && p[-1] == '<') {
            const char* after = p + n;
            if (after < end) {
                bool match = true;
                for (std::size_t j = 1; j < n; ++j) {
                    if (p[j] != name[j]) { match = false; break; }
                }
                if (match) {
                    char boundary = *after;
                    if (boundary == ' ' || boundary == '\t' || boundary == '\r' ||
                        boundary == '\n' || boundary == '>' || boundary == '/') {
                        return static_cast<std::size_t>(p - 1 - xml.data());
                    }
                }
            }
        }
    }
    return std::string_view::npos;
}

// --- SIMD close-tag scan: find `name` preceded by "</" and followed by '>' ---
// Returns the index of '<' (first char of the closing tag), matching
// the xmlscan_detail::findClose convention.

inline std::size_t findCloseTag(std::string_view xml, std::string_view name,
                                std::size_t from) noexcept {
    const std::size_t n = name.size();
    if (n == 0 || from >= xml.size()) return std::string_view::npos;
    if (from + n + 3 > xml.size()) return std::string_view::npos;

    const char* p = xml.data() + from;
    const char* end = xml.data() + xml.size();
    const char* limit = end - n - 1; // need room for name + '>'

#ifdef _M_AMD64
    const __m128i vFirst = _mm_set1_epi8(name[0]);
    for (; p + 16 <= limit; p += 16) {
        __m128i chunk = _mm_loadu_si128(reinterpret_cast<const __m128i*>(p));
        int mask = _mm_movemask_epi8(_mm_cmpeq_epi8(chunk, vFirst));
        while (mask) {
            int idx = _tzcnt_u32(static_cast<unsigned>(mask));
            const char* candidate = p + idx;
            if (candidate > xml.data() + 1 && candidate[-1] == '/' && candidate[-2] == '<') {
                const char* after = candidate + n;
                if (after < end && *after == '>') {
                    bool match = true;
                    for (std::size_t j = 1; j < n; ++j) {
                        if (candidate[j] != name[j]) { match = false; break; }
                    }
                    if (match) {
                        return static_cast<std::size_t>(candidate - 2 - xml.data());
                    }
                }
            }
            mask &= mask - 1;
        }
    }
#endif

    for (; p <= limit; ++p) {
        if (*p == name[0] && p > xml.data() + 1 && p[-1] == '/' && p[-2] == '<') {
            const char* after = p + n;
            if (after < end && *after == '>') {
                bool match = true;
                for (std::size_t j = 1; j < n; ++j) {
                    if (p[j] != name[j]) { match = false; break; }
                }
                if (match) {
                    return static_cast<std::size_t>(p - 2 - xml.data());
                }
            }
        }
    }
    return std::string_view::npos;
}

// --- SIMD attribute extractor: find attrName="..." in a start tag ---
// Replaces the xmlAttribute function's hot inner loop. Returns a view into `tag`.

inline std::string_view xmlAttribute(std::string_view tag, std::string_view name) noexcept {
    const std::size_t n = name.size();
    if (n == 0 || tag.size() < n + 3) return {};

    const char* p = tag.data();
    const char* end = tag.data() + tag.size();
    const char* limit = end - n - 2; // need name + ="

#ifdef _M_AMD64
    const __m128i vFirst = _mm_set1_epi8(name[0]);
    for (; p + 16 <= limit; p += 16) {
        __m128i chunk = _mm_loadu_si128(reinterpret_cast<const __m128i*>(p));
        int mask = _mm_movemask_epi8(_mm_cmpeq_epi8(chunk, vFirst));
        while (mask) {
            int idx = _tzcnt_u32(static_cast<unsigned>(mask));
            const char* candidate = p + idx;
            const char* after = candidate + n;
            if (after + 1 < end && *after == '=' && after[1] == '"') {
                if (candidate == tag.data() ||
                    !((*(candidate - 1) >= 'A' && *(candidate - 1) <= 'Z') ||
                      (*(candidate - 1) >= 'a' && *(candidate - 1) <= 'z') ||
                      (*(candidate - 1) >= '0' && *(candidate - 1) <= '9') ||
                      *(candidate - 1) == '_' || *(candidate - 1) == '-' ||
                      *(candidate - 1) == ':' || *(candidate - 1) == '.')) {
                    bool match = true;
                    for (std::size_t j = 1; j < n; ++j) {
                        if (candidate[j] != name[j]) { match = false; break; }
                    }
                    if (match) {
                        const char* valStart = after + 2;
                        const char* valEnd = findByte(valStart, end, '"');
                        if (valEnd) return {valStart, static_cast<std::size_t>(valEnd - valStart)};
                    }
                }
            }
            mask &= mask - 1;
        }
    }
#endif

    for (; p <= limit; ++p) {
        if (*p == name[0]) {
            const char* after = p + n;
            if (after + 1 < end && *after == '=' && after[1] == '"') {
                if (p == tag.data() ||
                    !((*(p - 1) >= 'A' && *(p - 1) <= 'Z') ||
                      (*(p - 1) >= 'a' && *(p - 1) <= 'z') ||
                      (*(p - 1) >= '0' && *(p - 1) <= '9') ||
                      *(p - 1) == '_' || *(p - 1) == '-' ||
                      *(p - 1) == ':' || *(p - 1) == '.')) {
                    bool match = true;
                    for (std::size_t j = 1; j < n; ++j) {
                        if (p[j] != name[j]) { match = false; break; }
                    }
                    if (match) {
                        const char* valStart = after + 2;
                        const char* valEnd = findByte(valStart, end, '"');
                        if (valEnd) return {valStart, static_cast<std::size_t>(valEnd - valStart)};
                    }
                }
            }
        }
    }
    return {};
}

} // namespace xlpp::internal::simd
