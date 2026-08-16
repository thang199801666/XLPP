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
                     const std::unordered_map<std::string, std::size_t>* sstIndex = nullptr,
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
