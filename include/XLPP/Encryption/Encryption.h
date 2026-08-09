#pragma once

#include <cstdint>
#include <filesystem>
#include <string>

namespace xlpp {

enum class OfficeEncryptionMode {
    None,
    AgileAes256Sha512,
    StandardAesSha1,
    Unsupported
};

struct OfficeEncryptionInfo {
    bool encrypted{false};
    bool supported{false};
    OfficeEncryptionMode mode{OfficeEncryptionMode::None};
    std::uint32_t spinCount{0};
    std::uint32_t keyBits{0};
    std::string cipherAlgorithm;
    std::string hashAlgorithm;
};

// Fast signature check. This identifies an OLE/CFB container; use
// inspectOfficeEncryption() to distinguish encrypted OOXML from unrelated CFB.
bool looksLikeEncryptedOfficeFile(const std::filesystem::path& path);

// Inspect the EncryptionInfo stream without requiring a password.
OfficeEncryptionInfo inspectOfficeEncryption(const std::filesystem::path& path);

} // namespace xlpp
