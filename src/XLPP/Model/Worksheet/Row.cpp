#include <XLPP/Worksheet/Row.h>
#include <XLPP/Worksheet/Worksheet.h>

#include <stdexcept>

namespace xlpp {

Row::Row(Worksheet& sheet, std::size_t rowNumber) : sheet_(&sheet), rowNumber_(rowNumber) {
    if (rowNumber == 0) throw std::invalid_argument("Row number is 1-based");
    if (rowNumber > MaxExcelRows) throw std::out_of_range("Row number exceeds Excel's 1,048,576-row limit");
}

Cell& Row::cell(std::size_t column) {
    return sheet_->cell(rowNumber_, column);
}

const Cell* Row::tryCell(std::size_t column) const noexcept {
    return sheet_->tryCell(rowNumber_, column);
}

std::vector<Cell*> Row::cells() {
    std::vector<Cell*> result;
    const auto extents = sheet_->extents();
    for (std::size_t column = extents.minColumn; column <= extents.maxColumn; ++column) {
        auto* value = &sheet_->cell(rowNumber_, column);
        if (!value->empty()) result.push_back(value);
    }
    return result;
}

std::vector<CellValue> Row::values() const {
    std::vector<CellValue> result;
    const auto extents = sheet_->extents();
    for (std::size_t column = extents.minColumn; column <= extents.maxColumn; ++column) {
        const auto* value = sheet_->tryCell(rowNumber_, column);
        result.push_back(value ? value->value() : CellValue{});
    }
    return result;
}

} // namespace xlpp
