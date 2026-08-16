#pragma once
#include <string>

namespace xlpp {
class Workbook;

namespace internal {
class ZipArchive;

// True when the OOXML package is a binary workbook (BIFF12: contains
// xl/workbook.bin rather than xl/workbook.xml).
bool isXlsbPackage(const xlpp::internal::ZipArchive& z);

// Reads a binary (BIFF12 / .xlsb) workbook: worksheet names, shared strings
// and cell values (numbers, shared strings, booleans, errors, RK and formula
// cached values). Styles, formulas and advanced objects are not yet
// materialized. Throws std::runtime_error on malformed input.
void readXlsb(const xlpp::internal::ZipArchive& z, xlpp::Workbook& workbook);

} // namespace internal
} // namespace xlpp
