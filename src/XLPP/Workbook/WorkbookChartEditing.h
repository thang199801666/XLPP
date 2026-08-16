#pragma once
#include <XLPP/Chart/Chart.h>
#include <cstddef>
#include <set>
#include <string>
#include <vector>

namespace xlpp {
class Worksheet;
struct PreservedPart;
struct PreservedRelationship;

namespace internal {
class ZipArchive;

// Chart serialization and imported-chart editing. These functions generate
// ChartML from the public chart model and patch ChartML fragments imported from
// an existing package. They carry no workbook state.

std::string chartXml(const xlpp::Chart& chart, bool strict);
std::string combinedChartXml(const xlpp::Chart& chart, bool strict);
std::string generatedChartSeriesXml(const xlpp::ChartSeries& series,
                                    xlpp::Chart::Type type,
                                    std::size_t index);

std::string chartSpaceDirectSpPr(const std::string& chartXmlText);
std::string plotAreaDirectSpPr(const std::string& plotArea);
bool patchNestedLineFormat(std::string& owner, const xlpp::ChartLineFormat& format);
std::string generatedChartWallXml(const char* localName, const xlpp::ChartWallFormat& format, bool prefixed);
std::string generatedPlotAuxiliaryXml(const xlpp::Chart::Plot& plot, bool strict);
std::string generatedDataTableXml(const xlpp::ChartDataTable& table, bool strict);
bool patchMarkerFormatInOwner(std::string& owner, const xlpp::ChartMarkerFormat& format);

// Finds the highest numeric suffix for parts matching "<prefix><n><suffix>"
// and returns it plus one (used to allocate media/chart part ids).
std::size_t nextAvailablePartId(const std::vector<xlpp::PreservedPart>& parts,
                                const std::string& prefix,
                                const std::string& suffix);

// Save-path package editing helpers (kept alongside the chart editor because
// they share its preserved-part/relationship machinery).
void suppressExclusivePartClosure(const std::string& rootPart,
                                  const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                  std::set<std::string>& suppressedPreservedParts);
bool applyChartChangesToPreservedDrawing(xlpp::internal::ZipArchive& z,
                                         const xlpp::Worksheet& sheet,
                                         const std::string& sourceSheetPart,
                                         const std::string& sourceSheetXml,
                                         const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                         const std::vector<xlpp::PreservedPart>& preservedParts,
                                         std::size_t& nextChartId,
                                         std::set<std::string>& suppressedPreservedParts);
bool applyImageChangesToPreservedDrawing(xlpp::internal::ZipArchive& z,
                                         const xlpp::Worksheet& sheet,
                                         const std::string& sourceSheetPart,
                                         const std::string& sourceSheetXml,
                                         const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                         const std::vector<xlpp::PreservedPart>& preservedParts,
                                         std::size_t& nextMediaId,
                                         std::set<std::string>& suppressedPreservedParts);
std::string rebuildWorksheetTail(std::string generated,
                                 const std::string& original,
                                 bool preserveDrawing,
                                 bool preservePivot,
                                 bool preserveTables,
                                 bool preserveComments);
std::string workbookViewsXml(const std::string& sourceWorkbookXml, std::size_t activeTab, std::size_t firstSheet);
std::string preserveWorkbookNodes(std::string generated,
                                  const std::string& original,
                                  bool preservePivotCaches);
std::string mergeWorkbookPivotCaches(const std::string& originalWorkbookXml,
                                     const std::string& generatedPivotCachesXml);
std::size_t nextAvailablePivotCacheId(const std::string& workbookXml);
std::size_t nextAvailableMediaId(const std::vector<xlpp::PreservedPart>& parts);

} // namespace internal
} // namespace xlpp
