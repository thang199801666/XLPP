#pragma once
#include <string>
#include <vector>
namespace xlpp::internal::preservation {
std::vector<std::string> extractTagBlocks(const std::string& xml, const std::string& tag);
void eraseTagBlocks(std::string& xml, const std::string& tag);
std::string joinBlocks(const std::vector<std::string>& blocks);
void insertBefore(std::string& xml, const std::string& marker, const std::string& content);
} // namespace xlpp::internal::preservation
