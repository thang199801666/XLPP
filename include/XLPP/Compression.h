#pragma once
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

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

// Password-to-open package encryption mode. Agile is the modern ECMA-376
// profile. Standard is the legacy CryptoAPI/AES profile used by older Office
// versions and some third-party producers.
enum class PackageEncryptionMode {
    Agile,
    Standard,
};

// Hash algorithm advertised by Agile EncryptionInfo. Standard Encryption is
// fixed to SHA-1 by the format and rejects any other selection.
enum class PackageEncryptionHash {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
};

// Coarse format returned by encryption-profile inspection.
enum class PackageEncryptionFormat {
    None,
    Agile,
    Standard,
    Unsupported,
};

// Metadata for an Agile certificate key-encryptor. P1I inspection exposes
// the DER certificate bytes and ciphertext sizes without requiring a private
// key or attempting certificate-based decryption.
struct PackageEncryptionCertificateKeyInfo {
    std::vector<unsigned char> x509Certificate;
    std::size_t encryptedKeyBytes{0};
    std::size_t certVerifierBytes{0};
    bool validEncoding{false};
};

// Parsed outer-package encryption metadata. Inspection never needs the
// password and never decrypts the inner OOXML package.
struct PackageEncryptionInfo {
    bool encrypted{false};
    PackageEncryptionFormat format{PackageEncryptionFormat::None};
    std::uint16_t versionMajor{0};
    std::uint16_t versionMinor{0};
    std::uint32_t keyBits{0};
    std::uint32_t blockSize{0};
    std::uint32_t hashSize{0};
    std::uint32_t spinCount{0};
    PackageEncryptionHash hashAlgorithm{PackageEncryptionHash::Sha512};
    bool hasDataIntegrity{false};
    bool supportedForRead{false};
    bool supportedForWrite{false};
    std::string cipherAlgorithm;
    std::string cipherChaining;
    std::size_t keyEncryptorCount{0};
    std::size_t passwordKeyEncryptorCount{0};
    std::vector<PackageEncryptionCertificateKeyInfo> certificateKeyEncryptors;
};

// Password-to-open package encryption. When enabled XL++ wraps the complete
// OOXML ZIP package in an ECMA-376 encrypted CFB container.
//
// Defaults are modern Agile AES-256-CBC/SHA-512 with 100,000 password spins.
// Agile supports AES-128/192/256 with SHA-1/SHA-256/SHA-384/SHA-512.
// Standard uses AES-128/192/256 with the format-mandated SHA-1 and fixed
// 50,000-iteration key derivation; `spinCount`/`hashAlgorithm` are ignored for
// Standard writes.
struct PackageEncryptionOptions {
    bool enabled{false};
    std::string password;
    PackageEncryptionMode mode{PackageEncryptionMode::Agile};
    PackageEncryptionHash hashAlgorithm{PackageEncryptionHash::Sha512};
    std::uint32_t keyBits{256};
    std::uint32_t spinCount{100000};
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
    // Optional semantic model gate before serialization. Disabled by default
    // so preservation-heavy workflows can intentionally save models that
    // contain non-fatal warnings (for example #REF! after a user edit). When
    // enabled, save() rejects model-integrity errors before writing bytes.
    bool validateModelBeforeSave{false};
    // Optional strict semantic mode: when model validation is enabled, treat
    // warnings (for example #REF! references) as save-blocking conditions.
    bool rejectModelWarningsBeforeSave{false};
    // Validate the fully assembled inner OPC/OOXML package in memory before it
    // is written or encrypted. This catches relationship/content-type/owner
    // topology errors that the in-memory workbook model cannot observe.
    bool validatePackageBeforeWrite{false};
    PackageEncryptionOptions encryption{};
};

} // namespace xlpp
