#pragma once
#include <cstddef>

namespace xlpp {

// Deflate compression level for package entries. `Store` writes entries
// uncompressed; `Fastest`/`Default`/`Best` map to zlib levels 1/6/9.
enum class CompressionLevel {
    Store,
    Fastest,
    Default,
    Best,
};

// Deflate strategy, mirroring the zlib strategy values. `Default` is the
// normal trade-off; the others target specialized data (e.g. text, RLE).
enum class CompressionStrategy {
    Default,
    Filtered,
    HuffmanOnly,
    Rle,
    Fixed,
};

// Tuning knobs for package writes. `parallelWorkers` of 0 disables
// parallelism (sequential, single-threaded); larger values bound the number
// of worker threads used for sheet serialization and entry compression.
// Parallel output is byte-for-byte identical to sequential output for the
// same level and strategy.
struct SaveOptions {
    CompressionLevel compressionLevel{CompressionLevel::Default};
    CompressionStrategy compressionStrategy{CompressionStrategy::Default};
    std::size_t parallelWorkers{0};
    bool parallelSheets{true};
    // Emit ISO 29500 strict OOXML namespaces (purl.oclc.org URIs) instead of
    // the transitional schemas.openxmlformats.org URIs.
    bool strictNamespace{false};
};

} // namespace xlpp
