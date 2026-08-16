#pragma once
#include <vector>

namespace xlpp {
class Workbook;

namespace internal {

// Writes a classic BIFF8 (.xls) workbook into an OLE2 compound-file container.
//
// The scope mirrors XlsBinaryReader's "basic read/write" foundation: worksheet
// names, cell values (numbers, shared strings, booleans, errors), row heights,
// column widths and merged ranges are persisted. Styles, formulas, charts,
// pivots and VBA are not yet serialized. Throws std::runtime_error on failure.
void writeLegacyXls(const xlpp::Workbook& workbook, std::vector<unsigned char>& out);

} // namespace internal
} // namespace xlpp
