#pragma once

#include <XLPP/Vba/VbaModule.h>
#include <cstddef>
#include <map>
#include <string>
#include <vector>

namespace xlpp::internal {

std::vector<unsigned char> buildVbaProjectBinary(const std::vector<xlpp::VbaModule>& modules,
                                                 const std::vector<std::string>& worksheetCodeNames,
                                                 const xlpp::VbaProjectInfo& projectInfo = {});
std::vector<xlpp::VbaModule> readVbaProjectBinary(const std::vector<unsigned char>& bytes);
xlpp::VbaProjectInfo readVbaProjectInfoBinary(const std::vector<unsigned char>& bytes);
bool isXlppGeneratedVbaProjectBinary(const std::vector<unsigned char>& bytes) noexcept;
std::string normalizeVbaSource(std::string source);
void validateVbaModuleName(const std::string& name);

// Generic compact CFB helpers shared by VBA and Office package encryption.
std::vector<unsigned char> buildRootCompoundFile(
    const std::map<std::string, std::vector<unsigned char>>& rootStreams);
bool isCompoundFile(const std::vector<unsigned char>& bytes) noexcept;
bool compoundFileContainsStream(const std::vector<unsigned char>& bytes, const std::string& path) noexcept;
std::vector<unsigned char> readCompoundFileStream(const std::vector<unsigned char>& bytes, const std::string& path);

} // namespace xlpp::internal
