#pragma once
#include <XLPP/Worksheet/Worksheet.h>
namespace xlpp::internal {
struct WorkbookDrawingAccess {
    static const auto& imageEdits(const xlpp::Worksheet& sheet) noexcept { return sheet.importedImageEdits_; }
    static const auto& chartEdits(const xlpp::Worksheet& sheet) noexcept { return sheet.importedChartEdits_; }
};
} // namespace xlpp::internal
