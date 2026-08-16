#pragma once
#include "WorkbookStylesIO.h"
#include <XLPP/Cell/RichText.h>
#include <XLPP/Worksheet/Filters/AutoFilter.h>
#include <XLPP/Worksheet/ConditionalFormatting/ConditionalFormatting.h>
#include <XLPP/Worksheet/DataValidation/DataValidation.h>
#include <cstddef>
#include <optional>
#include <sstream>
#include <string>
#include <string_view>
#include <unordered_map>

namespace xlpp {
class Worksheet;

namespace internal {

// FNV-1a 64-bit: faster than std::hash<std::string> for the hot shared-string
// table (MSVC's std::hash uses FNV on bytes too, but going through a
// string_view keeps the key comparison and lookup transparent and cheap).
struct SstFnvHash {
    std::size_t operator()(std::string_view value) const noexcept {
        std::size_t hash = 1469598103934665603ULL;
        for (char ch : value) {
            hash ^= static_cast<unsigned char>(ch);
            hash *= 1099511628211ULL;
        }
        return hash;
    }
    std::size_t operator()(const std::string& value) const noexcept {
        return (*this)(std::string_view(value));
    }
    static constexpr bool is_transparent = true;
};
using SstIndex = std::unordered_map<std::string, std::size_t, SstFnvHash, std::equal_to<std::string>>;

// Conditional-formatting / data-validation / auto-filter XML helpers shared by
// the sheet writer and reader.
std::string filterOperatorName(xlpp::FilterOperator op);
xlpp::FilterOperator parseFilterOperator(const std::string& value);
std::string conditionalOperatorName(xlpp::ConditionalOperator op);
xlpp::ConditionalOperator parseConditionalOperator(const std::string& value);
void writeCfvo(std::ostringstream& xml, const xlpp::Cfvo& cfvo);
xlpp::Cfvo parseCfvo(const std::string& tag);
std::string dataValidationTypeName(xlpp::DataValidationType type);
xlpp::DataValidationType parseDataValidationType(const std::string& value);
std::string dataValidationOperatorName(xlpp::DataValidationOperator op);
xlpp::DataValidationOperator parseDataValidationOperator(const std::string& value);
std::string dataValidationErrorStyleName(xlpp::DataValidationErrorStyle style);
xlpp::DataValidationErrorStyle parseDataValidationErrorStyle(const std::string& value);

// Serializes the worksheet part (xl/worksheets/sheetN.xml): rows, cells,
// rich text, dimensions, page settings and the legacy drawing/table/pivot
// blocks. Also parses inline rich-text runs for the reader.
std::string sheetXml(const xlpp::Worksheet& sheet,
                     const StyleCatalog& styles,
                     const DxfCatalog& dxfs,
                     bool date1904,
                     bool strict,
                     const SstIndex* sstIndex = nullptr,
                     std::size_t rowWorkers = 0,
                     std::string_view vbaCodeName = {});

std::optional<xlpp::RichText> parseRichTextRuns(std::string_view container);

// One entry of the shared-string table, either plain text or rich text.
struct LoadedSharedString {
    std::string plainText;
    std::optional<xlpp::RichText> richText;
};

} // namespace internal
} // namespace xlpp
