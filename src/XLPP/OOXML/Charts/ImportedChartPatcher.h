#pragma once
#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Chart/Chart.h>
#include <XLPP/Worksheet/Worksheet.h>
#include <cstddef>
#include <set>
#include <string>
#include <vector>
namespace xlpp::internal { class ZipArchive; }
namespace xlpp::internal::ooxml {
bool applyChartChangesToPreservedDrawing(xlpp::internal::ZipArchive& archive,
                                         const xlpp::Worksheet& sheet,
                                         const std::string& sourceSheetPart,
                                         const std::string& sourceSheetXml,
                                         const std::vector<xlpp::PreservedRelationship>& relationships,
                                         const std::vector<xlpp::PreservedPart>& preservedParts,
                                         std::size_t& nextChartId,
                                         std::set<std::string>& suppressedPreservedParts);
} // namespace xlpp::internal::ooxml
