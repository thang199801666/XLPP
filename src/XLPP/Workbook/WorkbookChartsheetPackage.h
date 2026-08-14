#pragma once

#include <cstddef>
#include <deque>
#include <functional>
#include <set>
#include <string>
#include <vector>

namespace xlpp {
class Chart;
class Chartsheet;
struct PreservedRelationship;
}

namespace xlpp::internal {
class ZipArchive;

// Writes every Chartsheet-owned package part (chartsheet XML/relationships,
// generated drawing/chart closure, and optional printerSettings payloads).
// Preservation-backed imports keep unrelated relationship siblings intact.
void writeChartsheetPackageParts(
    ZipArchive& archive,
    const std::deque<xlpp::Chartsheet>& chartsheets,
    const std::vector<std::size_t>& chartsheetPartIds,
    const std::vector<std::size_t>& chartsheetDrawingIds,
    const std::vector<std::size_t>& chartsheetChartIds,
    const std::vector<std::size_t>& chartsheetPrinterSettingsIds,
    const std::vector<xlpp::PreservedRelationship>& preservedRelationships,
    std::set<std::string>& suppressedPreservedParts,
    bool strict,
    const std::function<std::string(const xlpp::Chart&, bool)>& chartSerializer);

} // namespace xlpp::internal
