#pragma once
#include <XLPP/Worksheet/Worksheet.h>
#include <cstddef>
#include <string>
namespace xlpp::internal::ooxml {
std::string drawingXml(const xlpp::Worksheet& sheet, bool strict);
std::string drawingRelationshipsXml(const xlpp::Worksheet& sheet, std::size_t firstMediaId, std::size_t firstChartId, bool strict);
}
