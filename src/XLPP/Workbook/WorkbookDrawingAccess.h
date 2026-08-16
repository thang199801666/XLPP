#pragma once
#include <XLPP/Worksheet/Worksheet.h>

namespace xlpp {
namespace internal {

// Friend accessor into worksheet private drawing-edit state. Declared a friend
// of xlpp::Worksheet so drawing/chart editing modules can read the stable-ID
// edit lists without expanding the public API.
struct WorkbookDrawingAccess {
    static const auto& imageEdits(const xlpp::Worksheet& sheet) noexcept { return sheet.importedImageEdits_; }
    static const auto& chartEdits(const xlpp::Worksheet& sheet) noexcept { return sheet.importedChartEdits_; }
};

} // namespace internal
} // namespace xlpp
