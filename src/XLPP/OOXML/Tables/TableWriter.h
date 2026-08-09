#pragma once
#include <XLPP/Worksheet/Tables.h>
#include <XLPP/Worksheet/Worksheet.h>
#include <cstddef>
#include <string>
namespace xlpp::internal::ooxml {
std::string tableXml(const xlpp::Table& table, const xlpp::Worksheet& sheet, std::size_t id, bool strict);
}
