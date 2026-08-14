#pragma once

#include <cstddef>
#include <string>
#include <string_view>

namespace xlpp { class Chartsheet; }

namespace xlpp::internal {

std::string serializeChartsheetXml(const Chartsheet& sheet, bool strict,
                                   std::string_view drawingRelationshipId = "rId1",
                                   std::string_view printerSettingsRelationshipId = {});
std::string serializeChartsheetRelationshipsXml(std::size_t drawingId, bool strict);
std::string serializeChartsheetDrawingXml(bool strict);
std::string serializeChartsheetDrawingRelationshipsXml(std::size_t chartId, bool strict);
void parseChartsheetModel(Chartsheet& sheet, const std::string& xml);

} // namespace xlpp::internal
