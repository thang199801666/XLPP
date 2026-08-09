#pragma once
#include <filesystem>
#include <vector>
namespace xlpp::internal {
std::vector<unsigned char> readBinaryFile(const std::filesystem::path& path);
void writeBinaryFile(const std::filesystem::path& path, const std::vector<unsigned char>& bytes);
} // namespace xlpp::internal
