#pragma once
#include <XLPP/Worksheet/Worksheet.h>
#include <string>
namespace xlpp::internal::ooxml {
std::string commentsXml(const xlpp::Worksheet& sheet, bool strict);
std::string commentsVml(const xlpp::Worksheet& sheet);
}
