#pragma once

#include <XLPP/Encryption/Encryption.h>
#include <cstdint>
#include <string>
#include <vector>

namespace xlpp::internal {

std::vector<unsigned char> encryptAgileOfficePackage(const std::vector<unsigned char>& package,
                                                      const std::string& passwordUtf8,
                                                      std::uint32_t spinCount);
std::vector<unsigned char> encryptStandardOfficePackage(const std::vector<unsigned char>& package,
                                                         const std::string& passwordUtf8,
                                                         std::uint32_t keyBits);
std::vector<unsigned char> decryptAgileOfficePackage(const std::vector<unsigned char>& compoundFile,
                                                      const std::string& passwordUtf8,
                                                      bool verifyIntegrity);
std::vector<unsigned char> decryptOfficePackage(const std::vector<unsigned char>& compoundFile,
                                                const std::string& passwordUtf8,
                                                bool verifyIntegrity);
OfficeEncryptionInfo inspectOfficeEncryptionBytes(const std::vector<unsigned char>& compoundFile);
bool hasCompoundFileSignature(const unsigned char* data, std::size_t size) noexcept;

} // namespace xlpp::internal
