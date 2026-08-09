#pragma once
#include <cstddef>
#include <cstdint>
#include <string>
#include <XLPP/Encryption/Encryption.h>

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
    bool parallelRows{false};
    // Emit ISO 29500 strict OOXML namespaces (purl.oclc.org URIs) instead of
    // the transitional schemas.openxmlformats.org URIs.
    bool strictNamespace{false};
    // P0V opt-in save pipeline: synchronize chart title/category/value caches
    // on a private workbook copy immediately before serialization. The caller
    // workbook is not mutated. This is disabled by default for P0U compatibility.
    bool synchronizeChartCaches{false};
    // When synchronizeChartCaches is enabled, reuse dependency fingerprints to
    // skip unchanged references. On a workbook without a prior snapshot, all
    // supported references are synchronized and registered.
    bool synchronizeChangedChartCachesOnly{true};
    // Opt-in formula calculation pipeline. Formulas are evaluated on a private
    // workbook copy before serialization, so save() remains logically const.
    // When combined with chart synchronization, formulas are calculated first.
    bool calculateFormulasBeforeSave{false};
    // Write path-based saves through a same-directory temporary file and
    // atomically replace the destination only after serialization succeeds.
    // This prevents a failed save from truncating or partially overwriting an
    // existing workbook. Stream saves are unaffected.
    bool atomicWrite{true};
    // Flush file data/metadata to stable storage around path-based saves.
    // With atomicWrite enabled (the default), XL++ fsyncs/FlushFileBuffers the
    // staging file before replace and fsyncs the containing directory on POSIX
    // after rename. This closes the power-loss durability gap of rename-only
    // atomic saves. Disable only when save latency matters more than durability.
    bool durableWrite{true};
    // Validate modeled workbook invariants before package serialization.
    // Disable only when intentionally round-tripping a non-conforming source
    // package that must be preserved for forensic/repair workflows.
    bool validateBeforeSave{true};
    // Password-based ECMA-376 encryption. A non-empty password wraps the
    // generated OOXML ZIP package in an OLE/CFB encrypted container. Agile
    // AES-256/SHA-512 remains the default; Standard AES/SHA-1 is available for
    // interoperability with older Office-compatible producers/consumers.
    std::string encryptionPassword{};
    OfficeEncryptionMode encryptionMode{OfficeEncryptionMode::AgileAes256Sha512};
    std::uint32_t encryptionSpinCount{100000};
    std::uint32_t encryptionKeyBits{256}; // Standard mode: 128, 192 or 256.
};

} // namespace xlpp
