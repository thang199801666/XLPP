#pragma once
#include <cstddef>
#include <filesystem>
#include <string>
#include <vector>

namespace xlpp {
class Workbook;
class Worksheet;

namespace internal {

// Binary (BIFF12 / .xlsb) writer. Serializes the workbook model into .bin
// parts (workbook.bin, worksheets/sheetN.bin, sharedStrings.bin) packaged as
// an OOXML zip. Scope mirrors XlsbBinaryReader: worksheet names, cell values
// (numbers, shared strings, booleans, errors, RK) are persisted; styles,
// formulas and advanced objects are not yet serialized.
void writeXlsbPackage(const xlpp::Workbook& workbook,
                      const std::filesystem::path& path,
                      int compressionLevel);

} // namespace internal
} // namespace xlpp
