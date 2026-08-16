#pragma once
#include <cstddef>
#include <string>

namespace xlpp {
class Worksheet;
class Table;

namespace internal {

// Serializes the legacy comments part (xl/commentsN.xml).
std::string commentsXml(const xlpp::Worksheet& sheet, bool strict);

// Serializes the legacy comment shapes (VML drawing part).
std::string commentsVml(const xlpp::Worksheet& sheet);

// Serializes the spreadsheet drawing part (images and chart frames).
std::string drawingXml(const xlpp::Worksheet& sheet, bool strict);

// Serializes the drawing part's relationship list.
std::string drawingRelationshipsXml(const xlpp::Worksheet& sheet,
                                    std::size_t firstMediaId,
                                    std::size_t firstChartId,
                                    bool strict);

// Serializes a worksheet table (xl/tables/tableN.xml).
std::string tableXml(const xlpp::Table& table, const xlpp::Worksheet& sheet, std::size_t id, bool strict);

} // namespace internal
} // namespace xlpp
