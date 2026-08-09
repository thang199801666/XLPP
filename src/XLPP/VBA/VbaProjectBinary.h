#pragma once

#include <XLPP/Vba/VbaModule.h>
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp::internal {

std::vector<unsigned char> buildVbaProjectBinary(const std::vector<xlpp::VbaModule>& modules,
                                                 std::size_t worksheetCount);
std::vector<unsigned char> buildVbaProjectBinary(const std::vector<xlpp::VbaModule>& modules,
                                                 const std::vector<std::string>& worksheetCodeNames);
std::vector<unsigned char> buildVbaProjectBinary(const std::vector<xlpp::VbaModule>& modules,
                                                 const std::vector<std::string>& worksheetCodeNames,
                                                 const xlpp::VbaProjectProperties& properties);
std::vector<xlpp::VbaModule> readVbaProjectBinary(const std::vector<unsigned char>& bytes);
xlpp::VbaProjectProperties readVbaProjectProperties(const std::vector<unsigned char>& bytes);
bool isXlppGeneratedVbaProjectBinary(const std::vector<unsigned char>& bytes) noexcept;
std::string normalizeVbaSource(std::string source);
void validateVbaModuleName(const std::string& name);

} // namespace xlpp::internal
