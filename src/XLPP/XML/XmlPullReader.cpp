#include "XmlPullReader.h"
#include "XmlScanner.h"
#include <array>
#include <limits>
#include <stdexcept>
#include <utility>

namespace xlpp::internal {

XmlPullReader::XmlPullReader(Source source, std::size_t maxBufferedBytes)
    : source_(std::move(source)), maxBufferedBytes_(maxBufferedBytes) {
    if (!source_) throw std::invalid_argument("XmlPullReader requires a byte source");
}

bool XmlPullReader::refill() {
    std::array<unsigned char, 64 * 1024> chunk{};
    std::size_t capacity = chunk.size();
    if (maxBufferedBytes_) {
        if (buffer_.size() >= maxBufferedBytes_)
            throw std::runtime_error("XML element exceeds streaming buffer limit");
        capacity = std::min(capacity, maxBufferedBytes_ - buffer_.size());
    }
    const auto count = source_(chunk.data(), capacity);
    if (count > capacity)
        throw std::runtime_error("XmlPullReader source returned more bytes than requested");
    if (!count) return false;
    if (count > std::numeric_limits<std::size_t>::max() - buffer_.size())
        throw std::runtime_error("XML streaming buffer size overflow");
    buffer_.append(reinterpret_cast<const char*>(chunk.data()), count);
    return true;
}

void XmlPullReader::compact() {
    if (scan_ == 0) return;
    if (scan_ >= buffer_.size()) {
        buffer_.clear();
        scan_ = 0;
    } else if (scan_ > (1u << 20)) {
        buffer_.erase(0, scan_);
        scan_ = 0;
    }
}

std::string_view XmlPullReader::nextElement(std::string_view name) {
    if (name.empty()) throw std::invalid_argument("XmlPullReader element name cannot be empty");
    const std::size_t tnLen = name.size();
    const char* nameData = name.data();
    bool sourceExhausted = false;
    bool pendingElement = false;

    for (;;) {
        compact();
        const char* const bufData = buffer_.data();
        const std::size_t bufSize = buffer_.size();
        const char* const bufEnd = bufData + bufSize;

        // Scan for '<name'. Preserve a possible partial tag at the end so it
        // can be completed by the next refill.
        std::size_t p = scan_;
        for (;;) {
            while (p < bufSize && bufData[p] != '<') ++p;
            if (p >= bufSize) break;
            if (p + tnLen + 1 >= bufSize) { scan_ = p; break; }

            if (bufData[p + 1] == '/') { ++p; continue; }

            bool nameMatch = true;
            for (std::size_t j = 0; j < tnLen; ++j) {
                if (bufData[p + 1 + j] != nameData[j]) { nameMatch = false; break; }
            }
            if (nameMatch) {
                const char boundary = bufData[p + 1 + tnLen];
                if (boundary == ' ' || boundary == '\t' || boundary == '\r' || boundary == '\n' ||
                    boundary == '>' || boundary == '/') {
                    pendingElement = true;
                    // Quote-aware scan: an attribute value may legally contain
                    // '>', which must not terminate the opening tag.
                    const auto gtIndex = xmlscan_detail::scanGt(
                        std::string_view(bufData, bufSize), p + 1 + tnLen);
                    if (gtIndex == std::string_view::npos) { scan_ = p; break; }
                    const char* gt = bufData + gtIndex;

                    const char* beforeGt = gt;
                    while (beforeGt > bufData + p + 1 &&
                           (beforeGt[-1] == ' ' || beforeGt[-1] == '\t' ||
                            beforeGt[-1] == '\r' || beforeGt[-1] == '\n')) --beforeGt;
                    if (beforeGt > bufData + p + 1 && beforeGt[-1] == '/') {
                        std::string_view result(bufData + p, static_cast<std::size_t>(gt - (bufData + p) + 1));
                        scan_ = static_cast<std::size_t>(gt - bufData + 1);
                        return result;
                    }

                    // Find closing tag: </name>. Worksheet rows/cells cannot
                    // recursively contain the same element name; treating a
                    // missing terminator as malformed is preferable to silently
                    // returning EOF after buffering attacker-controlled input.
                    const char* findStart = gt + 1;
                    for (;;) {
                        while (findStart < bufEnd && !(findStart[0] == '<' && findStart + 1 < bufEnd && findStart[1] == '/')) ++findStart;
                        if (findStart >= bufEnd) { scan_ = p; goto needRefill; }

                        if (static_cast<std::size_t>(bufEnd - findStart) < tnLen + 3) { scan_ = p; goto needRefill; }

                        bool closeMatch = true;
                        for (std::size_t j = 0; j < tnLen; ++j) {
                            if (findStart[2 + j] != nameData[j]) { closeMatch = false; break; }
                        }
                        if (closeMatch && findStart[2 + tnLen] == '>') {
                            std::string_view result(bufData + p,
                                static_cast<std::size_t>(findStart + tnLen + 3 - (bufData + p)));
                            scan_ = static_cast<std::size_t>(findStart + tnLen + 3 - bufData);
                            return result;
                        }
                        ++findStart;
                    }
                }
            }
            ++p;
        }

        needRefill:
        if (sourceExhausted) {
            if (pendingElement)
                throw std::runtime_error("Truncated XML element <" + std::string(name) + ">");
            return {};
        }
        sourceExhausted = !refill();
    }
}

} // namespace xlpp::internal
