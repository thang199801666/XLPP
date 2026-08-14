#pragma once

#include <XLPP/Compression.h>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <string>
#include <vector>

namespace xlpp::internal {

struct AgileEncryptionParameters {
    std::uint32_t spinCount{100000};
    std::uint32_t keyBits{256};
    PackageEncryptionHash hashAlgorithm{PackageEncryptionHash::Sha512};
};

struct StandardEncryptionParameters {
    std::uint32_t keyBits{128};
};

struct OfficeDecryptionLimits {
    std::uint32_t maxSpinCount{1000000}; // 0 -> format maximum (10,000,000)
    std::size_t maxPlainPackageBytes{0}; // 0 -> bounded only by ciphertext geometry/platform size
    bool allowStandardEncryption{true};
    bool requireAgileDataIntegrity{false};
    std::size_t maxEncryptionInfoBytes{1024u * 1024u}; // 0 -> unlimited
};

bool isEncryptedOfficeCompoundFile(const std::filesystem::path& path) noexcept;
PackageEncryptionInfo inspectOfficeEncryption(const std::vector<unsigned char>& compoundFileBytes);
PackageEncryptionInfo inspectOfficeEncryption(const std::filesystem::path& path);
std::vector<unsigned char> encryptAgileOfficePackage(
    const std::vector<unsigned char>& packageBytes,
    const std::string& password,
    const AgileEncryptionParameters& parameters = {});
std::vector<unsigned char> encryptStandardOfficePackage(
    const std::vector<unsigned char>& packageBytes,
    const std::string& password,
    const StandardEncryptionParameters& parameters = {});
std::vector<unsigned char> decryptOfficePackage(
    const std::vector<unsigned char>& compoundFileBytes,
    const std::string& password,
    const OfficeDecryptionLimits& limits = {});
std::vector<unsigned char> readBinaryFile(const std::filesystem::path& path);
void writeBinaryFile(const std::filesystem::path& path, const std::vector<unsigned char>& bytes);

} // namespace xlpp::internal
