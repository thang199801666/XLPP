#include <XLPP/Worksheet/Range.h>
#include <XLPP/Worksheet/Worksheet.h>
#include <stdexcept>

namespace xlpp {
CellRange::CellRange(Worksheet& worksheet, std::size_t minRow, std::size_t minColumn,
                     std::size_t maxRow, std::size_t maxColumn)
    : worksheet_(&worksheet), minRow_(minRow), minColumn_(minColumn),
      maxRow_(maxRow), maxColumn_(maxColumn) {
    if (minRow == 0 || minColumn == 0 || maxRow == 0 || maxColumn == 0)
        throw std::invalid_argument("Range coordinates are 1-based");
    if (minRow > maxRow || minColumn > maxColumn)
        throw std::invalid_argument("Range minimum cannot exceed maximum");
}

std::string CellRange::address() const {
    return CellReference{minRow_, minColumn_}.address() + ":" +
           CellReference{maxRow_, maxColumn_}.address();
}

Cell& CellRange::cell(std::size_t relativeRow, std::size_t relativeColumn) {
    if (relativeRow == 0 || relativeColumn == 0 ||
        relativeRow > rowCount() || relativeColumn > columnCount())
        throw std::out_of_range("CellRange relative coordinates are outside the range");
    return worksheet_->cell(minRow_ + relativeRow - 1, minColumn_ + relativeColumn - 1);
}

std::vector<Cell*> CellRange::cells() {
    std::vector<Cell*> result;
    result.reserve(rowCount() * columnCount());
    for (std::size_t row = minRow_; row <= maxRow_; ++row)
        for (std::size_t column = minColumn_; column <= maxColumn_; ++column)
            result.push_back(&worksheet_->cell(row, column));
    return result;
}

std::vector<std::vector<Cell*>> CellRange::rows() {
    std::vector<std::vector<Cell*>> result;
    result.reserve(rowCount());
    for (std::size_t row = minRow_; row <= maxRow_; ++row) {
        std::vector<Cell*> current;
        current.reserve(columnCount());
        for (std::size_t column = minColumn_; column <= maxColumn_; ++column)
            current.push_back(&worksheet_->cell(row, column));
        result.push_back(std::move(current));
    }
    return result;
}

void CellRange::setValue(const CellValue& value) {
    for (auto* target : cells()) target->setValue(value);
}

void CellRange::clear() {
    for (auto* target : cells()) target->clear();
}

void CellRange::forEach(const std::function<void(Cell&)>& callback) {
    for (std::size_t row = minRow_; row <= maxRow_; ++row)
        for (std::size_t column = minColumn_; column <= maxColumn_; ++column)
            callback(worksheet_->cell(row, column));
}

std::vector<CellValue> CellRange::values() const {
    std::vector<CellValue> result;
    result.reserve(rowCount() * columnCount());
    for (std::size_t row = minRow_; row <= maxRow_; ++row)
        for (std::size_t column = minColumn_; column <= maxColumn_; ++column)
            result.push_back(worksheet_->cell(row, column).value());
    return result;
}

std::vector<std::string> CellRange::formulas() const {
    std::vector<std::string> result;
    result.reserve(rowCount() * columnCount());
    for (std::size_t row = minRow_; row <= maxRow_; ++row)
        for (std::size_t column = minColumn_; column <= maxColumn_; ++column)
            result.push_back(worksheet_->cell(row, column).formula());
    return result;
}
}
