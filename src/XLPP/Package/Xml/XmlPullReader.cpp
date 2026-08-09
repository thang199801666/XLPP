#include "XmlPullReader.h"
#include <array>
#include <utility>

namespace xlpp::internal {

XmlPullReader::XmlPullReader(Source source) : source_(std::move(source)) {}

bool XmlPullReader::refill() {
    std::array<unsigned char, 64 * 1024> chunk{};
    const auto count = source_(chunk.data(), chunk.size());
    if (!count) return false;
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
    const std::size_t tnLen = name.size();
    const char* nameData = name.data();
    bool sourceExhausted = false;

    for (;;) {
        compact();
        const char* const bufData = buffer_.data();
        const std::size_t bufSize = buffer_.size();
        const char* const bufEnd = bufData + bufSize;

        // Scan for '<name'
        std::size_t p = scan_;
        for (;;) {
            while (p < bufSize && bufData[p] != '<') ++p;
            if (p >= bufSize) break;
            if (p + tnLen + 1 >= bufSize) break;

            if (bufData[p + 1] == '/') { ++p; continue; }

            bool nameMatch = true;
            for (std::size_t j = 0; j < tnLen; ++j) {
                if (bufData[p + 1 + j] != nameData[j]) { nameMatch = false; break; }
            }
            if (nameMatch) {
                const char boundary = bufData[p + 1 + tnLen];
                if (boundary == ' ' || boundary == '\t' || boundary == '\r' || boundary == '\n' ||
                    boundary == '>' || boundary == '/') {
                    // Found open tag
                    const char* gt = bufData + p + 1 + tnLen;
                    for (; gt < bufEnd && *gt != '>'; ++gt) {}
                    if (gt >= bufEnd) { scan_ = p; break; }

                    if (gt > bufData + p + 1 && *(gt - 1) == '/') {
                        std::string_view result(bufData + p, static_cast<std::size_t>(gt - (bufData + p) + 1));
                        scan_ = static_cast<std::size_t>(gt - bufData + 1);
                        return result;
                    }

                    // Find closing tag: </name>
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
        if (sourceExhausted) return {};
        sourceExhausted = !refill();
    }
}

} // namespace xlpp::internal
