#pragma once
#include <cstddef>
#include <functional>
#include <string>
#include <string_view>

namespace xlpp::internal {

// Incremental XML element scanner. Consumes a pull source of bytes and yields
// complete elements lazily, keeping only a bounded look-ahead buffer in memory.
class XmlPullReader {
public:
    using Source = std::function<std::size_t(unsigned char*, std::size_t)>;

    static constexpr std::size_t kDefaultMaxBufferedBytes = 64u * 1024u * 1024u;

    explicit XmlPullReader(Source source, std::size_t maxBufferedBytes = kDefaultMaxBufferedBytes);

    // Yields the next complete element with the given tag name (self-closing
    // tags included) as a string_view into the internal buffer. The returned
    // view is valid only until the next call to nextElement, which may refill
    // or compact the buffer. An empty view signals the end of the source.
    std::string_view nextElement(std::string_view name);

private:
    bool refill();
    void compact();

    Source source_;
    std::string buffer_;
    std::size_t scan_{0};
    std::size_t maxBufferedBytes_{kDefaultMaxBufferedBytes};
};

} // namespace xlpp::internal
