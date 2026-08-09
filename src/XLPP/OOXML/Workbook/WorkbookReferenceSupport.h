#pragma once
#include <string>
namespace xlpp::internal::ooxml {
std::string quoteSheetName(const std::string& name);
std::string absoluteReference(std::string reference);
std::string qualifiedPrintReference(const std::string& sheetName, const std::string& reference);
std::string localPrintReference(std::string value);
} // namespace xlpp::internal::ooxml
