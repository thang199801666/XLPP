#include "WorksheetReferenceSupport.h"
#include <stdexcept>
#include <string>
#include <utility>

namespace xlpp::internal {
std::pair<CellReference, CellReference> parseWorksheetRangeAddress(std::string_view address) {
    const auto colon = address.find(':');
    if (colon == std::string_view::npos) {
        const auto ref = CellReference::parse(std::string(address));
        return {ref, ref};
    }
    if (address.find(':', colon + 1) != std::string_view::npos)
        throw std::invalid_argument("Invalid range address: " + std::string(address));
    auto first = CellReference::parse(std::string(address.substr(0, colon)));
    auto last = CellReference::parse(std::string(address.substr(colon + 1)));
    if (first.row > last.row) std::swap(first.row, last.row);
    if (first.column > last.column) std::swap(first.column, last.column);
    return {first, last};
}
} // namespace xlpp::internal
