#pragma once
#include <XLPP/Cell/CellReference.h>
#include <string_view>
#include <utility>

namespace xlpp::internal {
std::pair<CellReference, CellReference> parseWorksheetRangeAddress(std::string_view address);
} // namespace xlpp::internal
