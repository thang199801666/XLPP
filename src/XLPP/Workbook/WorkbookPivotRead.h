#pragma once
#include <string>

namespace xlpp {
class Worksheet;

namespace internal {
class ZipArchive;

// Loads pivot tables (and their caches) referenced by a worksheet part into
// the worksheet model.
void loadPivotTables(xlpp::Worksheet& ws,
                     const std::string& sheetXml,
                     const xlpp::internal::ZipArchive& z,
                     const std::string& sheetPart);

} // namespace internal
} // namespace xlpp
