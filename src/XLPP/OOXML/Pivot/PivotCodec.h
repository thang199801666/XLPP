#pragma once

#include <XLPP/Pivot/PivotTable.h>
#include <XLPP/Worksheet/Worksheet.h>

#include <cstddef>
#include <deque>
#include <string>

namespace xlpp { struct PreservedRelationship; }
namespace xlpp::internal { class ZipArchive; }

namespace xlpp::internal::ooxml {

std::string pivotTableXml(const xlpp::PivotTable& pivot, std::size_t id, bool strict);
std::string pivotCacheXml(const xlpp::PivotTable& pivot, bool strict);
std::string pivotCacheRecordsXml(const xlpp::PivotTable& pivot, bool strict);
std::string quotePivotSheetName(const std::string& name);
void loadImportedPivotModels(xlpp::Worksheet& sheet,
                             const xlpp::internal::ZipArchive& archive,
                             const std::string& sourceSheetPart,
                             const std::vector<xlpp::PreservedRelationship>& relationships);

xlpp::PivotTable effectivePivotTable(const xlpp::PivotTable& source,
                                     const std::deque<xlpp::Worksheet>& sheets,
                                     const xlpp::Worksheet& owner,
                                     std::size_t cacheId);

} // namespace xlpp::internal::ooxml
