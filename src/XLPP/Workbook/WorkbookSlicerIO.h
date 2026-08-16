#pragma once
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {
struct Slicer;

namespace internal {

// Serializes a slicer cache definition (xl/slicerCaches/slicerCacheN.xml).
std::string slicerCacheXml(const xlpp::Slicer& slicer);

// Serializes the per-worksheet slicer part (xl/slicers/slicerN.xml).
std::string slicerXml(const xlpp::Slicer& slicer);

// Relationship ids used for the workbook-level cache and the sheet-level
// slicer control.
std::string slicerCacheRelationshipId(std::size_t index);
std::string slicerRelationshipId(std::size_t index);

// Inserts the x14 slicerList extension before </worksheet>.
std::string insertSlicerListExt(std::string sheetXml, const std::vector<std::string>& relationshipIds,
                                bool strict);

// Inserts the x14 slicerCaches extension before </workbook>.
std::string insertWorkbookSlicerCachesExt(std::string workbookXml,
                                          const std::vector<std::string>& relationshipIds,
                                          bool strict);

} // namespace internal
} // namespace xlpp
