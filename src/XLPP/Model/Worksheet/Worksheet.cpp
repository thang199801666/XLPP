#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Worksheet/WorksheetName.h>
#include "WorksheetReferenceSupport.h"
#include <algorithm>
#include <cctype>
#include <stdexcept>
#include <utility>

namespace xlpp {
void Worksheet::rename(std::string name) {
    validateWorksheetName(name);
    name_ = std::move(name);
    dirty_ = true;
}

void Worksheet::setVbaCodeName(std::string codeName) {
    if (codeName.empty() || codeName.size() > 31)
        throw std::invalid_argument("VBA code name must contain 1-31 characters");
    const auto first = static_cast<unsigned char>(codeName.front());
    if (!(std::isalpha(first) || first == '_'))
        throw std::invalid_argument("VBA code name must begin with a letter or underscore");
    for (const char raw : codeName) {
        const auto ch = static_cast<unsigned char>(raw);
        if (!(std::isalnum(ch) || ch == '_'))
            throw std::invalid_argument("VBA code name contains an invalid character");
    }
    std::string folded = codeName;
    std::transform(folded.begin(), folded.end(), folded.begin(), [](unsigned char ch) {
        return static_cast<char>(std::tolower(ch));
    });
    if (folded == "thisworkbook")
        throw std::invalid_argument("ThisWorkbook is reserved for the workbook document module");
    vbaCodeName_ = std::move(codeName);
    dirty_ = true;
}

Cell& Worksheet::cell(const std::string& address) {
    const auto ref = CellReference::parse(address);
    return cell(ref.row, ref.column);
}

Cell& Worksheet::cell(std::size_t row, std::size_t column) {
    if (!isValidCellCoordinate(row, column))
        throw std::out_of_range("Cell coordinate exceeds Excel worksheet bounds");
    const auto key = makeCellKey(row, column);
    auto [it, inserted] = cells_.try_emplace(key, row, column);
    dirty_ = true;
    return it->second;
}

const Cell* Worksheet::tryCell(const std::string& address) const noexcept {
    try {
        const auto ref = CellReference::parse(address);
        return tryCell(ref.row, ref.column);
    } catch (...) {
        return nullptr;
    }
}

const Cell* Worksheet::tryCell(std::size_t row, std::size_t column) const noexcept {
    try {
        const auto it = cells_.find(makeCellKey(row, column));
        return it == cells_.end() ? nullptr : &it->second;
    } catch (...) {
        return nullptr;
    }
}

CellRange Worksheet::range(const std::string& address) {
    const auto [first, last] = internal::parseWorksheetRangeAddress(address);
    return CellRange(*this, first.row, first.column, last.row, last.column);
}

CellRange Worksheet::range(std::size_t minRow, std::size_t minColumn,
                           std::size_t maxRow, std::size_t maxColumn) {
    if (minRow > maxRow) std::swap(minRow, maxRow);
    if (minColumn > maxColumn) std::swap(minColumn, maxColumn);
    return CellRange(*this, minRow, minColumn, maxRow, maxColumn);
}

std::size_t Worksheet::maxRow() const noexcept {
    std::size_t result = 1;
    for (const auto& [_, value] : cells_)
        if (value.row() > result) result = value.row();
    return result;
}

std::size_t Worksheet::maxColumn() const noexcept {
    std::size_t result = 1;
    for (const auto& [_, value] : cells_)
        if (value.column() > result) result = value.column();
    return result;
}

WorksheetExtents Worksheet::extents() const noexcept {
    if (cells_.empty()) return {1, 1, 1, 1};
    std::size_t minRow = static_cast<std::size_t>(-1);
    std::size_t minColumn = static_cast<std::size_t>(-1);
    std::size_t maxRowValue = 1;
    std::size_t maxColumnValue = 1;
    for (const auto& [_, value] : cells_) {
        if (value.row() < minRow) minRow = value.row();
        if (value.column() < minColumn) minColumn = value.column();
        if (value.row() > maxRowValue) maxRowValue = value.row();
        if (value.column() > maxColumnValue) maxColumnValue = value.column();
    }
    return {minRow, minColumn, maxRowValue, maxColumnValue};
}

std::string Worksheet::dimensions() const {
    const auto e = extents();
    return CellReference{e.minRow, e.minColumn}.address() + ":" +
           CellReference{e.maxRow, e.maxColumn}.address();
}

void Worksheet::append(const std::vector<CellValue>& values) {
    const auto targetRow = cells_.empty() ? 1 : 1 + (cells_.rbegin()->first >> 20);
    for (std::size_t column = 1; column <= values.size(); ++column)
        cell(targetRow, column).setValue(values[column - 1]);
    dirty_ = true;
}

void Worksheet::mergeCells(const std::string& rangeAddress) {
    const auto [first, last] = internal::parseWorksheetRangeAddress(rangeAddress);
    const auto canonical = CellReference{first.row, first.column}.address() + ":" +
                           CellReference{last.row, last.column}.address();
    if (first.row == last.row && first.column == last.column)
        throw std::invalid_argument("Merged range must contain at least two cells");
    if (std::find(mergedRanges_.begin(), mergedRanges_.end(), canonical) == mergedRanges_.end()) {
        mergedRanges_.push_back(canonical);
        mergedRangesParsed_.push_back({first.row, first.column, last.row, last.column});
        dirty_ = true;
    }
}

void Worksheet::unmergeCells(const std::string& rangeAddress) {
    const auto [first, last] = internal::parseWorksheetRangeAddress(rangeAddress);
    const auto canonical = CellReference{first.row, first.column}.address() + ":" +
                           CellReference{last.row, last.column}.address();
    const auto it = std::find(mergedRanges_.begin(), mergedRanges_.end(), canonical);
    if (it == mergedRanges_.end()) throw std::invalid_argument("Merged range not found: " + canonical);
    const auto index = static_cast<std::size_t>(std::distance(mergedRanges_.begin(), it));
    mergedRanges_.erase(it);
    mergedRangesParsed_.erase(mergedRangesParsed_.begin() + static_cast<std::ptrdiff_t>(index));
    dirty_ = true;
}

bool Worksheet::isMerged(const std::string& cellAddress) const {
    const auto target = CellReference::parse(cellAddress);
    for (const auto& range : mergedRangesParsed_) {
        if (target.row >= range.minRow && target.row <= range.maxRow &&
            target.column >= range.minColumn && target.column <= range.maxColumn) return true;
    }
    return false;
}

void Worksheet::freezePanes(const std::string& topLeftCell) {
    const auto ref = CellReference::parse(topLeftCell);
    freezePane_ = ref.address();
    dirty_ = true;
}

RowDimension& Worksheet::rowDimension(std::size_t row) {
    if (row == 0) throw std::invalid_argument("Row index is 1-based");
    if (row > MaxExcelRows) throw std::out_of_range("Row index exceeds Excel's 1,048,576-row limit");
    dirty_ = true;
    return rowDimensions_[row];
}

const RowDimension* Worksheet::tryRowDimension(std::size_t row) const noexcept {
    const auto it = rowDimensions_.find(row);
    return it == rowDimensions_.end() ? nullptr : &it->second;
}

ColumnDimension& Worksheet::columnDimension(std::size_t column) {
    if (column == 0) throw std::invalid_argument("Column index is 1-based");
    if (column > MaxExcelColumns) throw std::out_of_range("Column index exceeds Excel's 16,384-column limit");
    dirty_ = true;
    return columnDimensions_[column];
}

ColumnDimension& Worksheet::columnDimension(const std::string& columnName) {
    return columnDimension(CellReference::columnIndex(columnName));
}

const ColumnDimension* Worksheet::tryColumnDimension(std::size_t column) const noexcept {
    const auto it = columnDimensions_.find(column);
    return it == columnDimensions_.end() ? nullptr : &it->second;
}

Table& Worksheet::addTable(std::string name, std::string reference) {
    if (table(name)) throw std::invalid_argument("Duplicate table name: " + name);
    tables_.emplace_back(std::move(name), std::move(reference));
    dirty_ = true;
    return tables_.back();
}
Table* Worksheet::table(const std::string& name) noexcept {
    auto it = std::find_if(tables_.begin(), tables_.end(), [&](const Table& value){ return value.name() == name; });
    if (it != tables_.end()) dirty_ = true;
    return it == tables_.end() ? nullptr : &*it;
}
const Table* Worksheet::table(const std::string& name) const noexcept {
    auto it = std::find_if(tables_.begin(), tables_.end(), [&](const Table& value){ return value.name() == name; });
    return it == tables_.end() ? nullptr : &*it;
}

std::vector<Row> Worksheet::rows() {
    std::vector<Row> result;
    result.reserve(rowCount());
    for (std::size_t r = 1; r <= rowCount(); ++r)
        result.emplace_back(*this, r);
    return result;
}

std::vector<std::vector<CellValue>> Worksheet::iterRows(std::size_t minRow, std::size_t maxRow,
                                                         std::size_t minCol, std::size_t maxCol) const {
    const auto e = extents();
    if (minRow == 0) minRow = e.minRow;
    if (maxRow == 0) maxRow = e.maxRow;
    if (minCol == 0) minCol = e.minColumn;
    if (maxCol == 0) maxCol = e.maxColumn;
    std::vector<std::vector<CellValue>> result;
    result.reserve(maxRow - minRow + 1);
    for (std::size_t r = minRow; r <= maxRow; ++r) {
        std::vector<CellValue> row;
        row.reserve(maxCol - minCol + 1);
        for (std::size_t c = minCol; c <= maxCol; ++c) {
            const auto* cell = tryCell(r, c);
            row.push_back(cell ? cell->value() : CellValue{});
        }
        result.push_back(std::move(row));
    }
    return result;
}

std::vector<std::vector<CellValue>> Worksheet::iterCols(std::size_t minRow, std::size_t maxRow,
                                                         std::size_t minCol, std::size_t maxCol) const {
    const auto e = extents();
    if (minRow == 0) minRow = e.minRow;
    if (maxRow == 0) maxRow = e.maxRow;
    if (minCol == 0) minCol = e.minColumn;
    if (maxCol == 0) maxCol = e.maxColumn;
    std::vector<std::vector<CellValue>> result;
    result.reserve(maxCol - minCol + 1);
    for (std::size_t c = minCol; c <= maxCol; ++c) {
        std::vector<CellValue> col;
        col.reserve(maxRow - minRow + 1);
        for (std::size_t r = minRow; r <= maxRow; ++r) {
            const auto* cell = tryCell(r, c);
            col.push_back(cell ? cell->value() : CellValue{});
        }
        result.push_back(std::move(col));
    }
    return result;
}

} // namespace xlpp
