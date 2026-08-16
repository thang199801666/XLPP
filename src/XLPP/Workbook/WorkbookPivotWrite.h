#pragma once
#include <XLPP/Pivot/PivotTable.h>
#include <cstddef>
#include <deque>
#include <sstream>
#include <string>
#include <string_view>
#include <vector>

namespace xlpp {
class Cell;
class Worksheet;

namespace internal {

// Classification of a single pivot cache/record cell value for serialization.
enum class PivotValueKind { Blank, Number, Boolean, Error, String, DateTime };

struct PivotSharedItem {
    std::string value;
    PivotValueKind kind{PivotValueKind::String};
};

int resolvedPivotFieldIndex(const xlpp::PivotCache& cache, int index, const std::string& name);
PivotValueKind pivotValueKind(xlpp::PivotCacheValueKind kind);
xlpp::PivotCacheValueKind publicPivotValueKind(std::string_view nodeType);
xlpp::PivotCacheValueKind publicPivotValueKind(PivotValueKind kind);
std::vector<std::vector<PivotSharedItem>> pivotSharedItems(const xlpp::PivotCache& cache);
std::string pivotTableXml(const xlpp::PivotTable& pt, std::size_t id, bool strict);
xlpp::PivotTable effectivePivotTable(const xlpp::PivotTable& source,
                                     const std::deque<xlpp::Worksheet>& sheets,
                                     const xlpp::Worksheet& owner,
                                     std::size_t cacheId);
bool pivotFieldIsPureData(const xlpp::PivotTable& pt, std::size_t fieldIndex);
void writePivotValue(std::ostringstream& xml, const std::string& value, PivotValueKind kind);
void writePivotValue(std::ostringstream& xml, const std::string& value);
std::string pivotCacheXml(const xlpp::PivotTable& pt, bool strict);
std::string pivotCacheRecordsXml(const xlpp::PivotTable& pt, bool strict);
std::string quotePivotSheetName(const std::string& name);
std::string pivotCellText(const xlpp::Cell* cell);
bool pivotCachesEquivalent(const xlpp::PivotCache& left, const xlpp::PivotCache& right);

} // namespace internal
} // namespace xlpp
