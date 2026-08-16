#pragma once
#include <cstddef>
#include <string_view>
#include <vector>

namespace xlpp {
class Workbook;

namespace internal {

// Classic binary workbook container detection. BIFF8 files (.xls) and
// Office-encrypted packages are both OLE2 compound files; the caller decides
// between decryption and legacy parsing by inspecting the directory entries.
inline constexpr std::string_view Ole2CompoundMagic = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";

// True when `bytes` begins with the OLE2 compound-file (CFB) header.
bool isOle2CompoundFile(std::string_view bytes) noexcept;

// Reads BIFF8 (classic .xls) worksheet cell data into `workbook`.
//
// The current scope is deliberately the "basic read/write" foundation: it
// materializes worksheet names and cell values (numbers, shared strings,
// booleans, errors, formula cached results) across all sheets. Formulas,
// styles, formatting, charts, pivots and VBA are not yet materialized; those
// are follow-on milestones. Throws std::runtime_error on malformed input.
void readLegacyXls(const std::vector<unsigned char>& bytes, xlpp::Workbook& workbook);

} // namespace internal
} // namespace xlpp
