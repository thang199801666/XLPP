#include <XLPP/Worksheet/Worksheet.h>
#include <algorithm>
#include <stdexcept>
#include <utility>

namespace {
std::pair<xlpp::CellReference, xlpp::CellReference> parseRangeAddress(const std::string& address) {
    const auto colon = address.find(':');
    if (colon == std::string::npos) {
        const auto ref = xlpp::CellReference::parse(address);
        return {ref, ref};
    }
    if (address.find(':', colon + 1) != std::string::npos)
        throw std::invalid_argument("Invalid range address: " + address);
    auto first = xlpp::CellReference::parse(address.substr(0, colon));
    auto last = xlpp::CellReference::parse(address.substr(colon + 1));
    if (first.row > last.row) std::swap(first.row, last.row);
    if (first.column > last.column) std::swap(first.column, last.column);
    return {first, last};
}
}

namespace xlpp {
void Worksheet::rename(std::string name) {
    if (name.empty()) throw std::invalid_argument("Worksheet name cannot be empty");
    name_ = std::move(name);
    dirty_ = true;
}

Cell& Worksheet::cell(const std::string& address) {
    const auto ref = CellReference::parse(address);
    return cell(ref.row, ref.column);
}

Cell& Worksheet::cell(std::size_t row, std::size_t column) {
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
    const auto [first, last] = parseRangeAddress(address);
    return CellRange(*this, first.row, first.column, last.row, last.column);
}

CellRange Worksheet::range(std::size_t minRow, std::size_t minColumn,
                           std::size_t maxRow, std::size_t maxColumn) {
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

void Worksheet::insertRows(std::size_t index, std::size_t amount) {
    if (index == 0 || amount == 0) throw std::invalid_argument("Row index and amount must be positive");
    shiftRows(index, amount, true);
}

void Worksheet::deleteRows(std::size_t index, std::size_t amount) {
    if (index == 0 || amount == 0) throw std::invalid_argument("Row index and amount must be positive");
    shiftRows(index, amount, false);
}

void Worksheet::insertColumns(std::size_t index, std::size_t amount) {
    if (index == 0 || amount == 0) throw std::invalid_argument("Column index and amount must be positive");
    shiftColumns(index, amount, true);
}

void Worksheet::deleteColumns(std::size_t index, std::size_t amount) {
    if (index == 0 || amount == 0) throw std::invalid_argument("Column index and amount must be positive");
    shiftColumns(index, amount, false);
}

void Worksheet::shiftRows(std::size_t index, std::size_t amount, bool insert) {
    std::map<std::uint64_t, Cell> shifted;
    const auto deleteEnd = index + amount;
    for (auto& [key, source] : cells_) {
        auto row = source.row();
        const auto column = source.column();
        if (insert) {
            if (row >= index) row += amount;
        } else {
            if (row >= index && row < deleteEnd) continue;
            if (row >= deleteEnd) row -= amount;
        }
        source.setPosition(row, column);
        shifted.emplace(makeCellKey(row, column), std::move(source));
    }
    cells_ = std::move(shifted);
    dirty_ = true;
}

void Worksheet::shiftColumns(std::size_t index, std::size_t amount, bool insert) {
    std::map<std::uint64_t, Cell> shifted;
    const auto deleteEnd = index + amount;
    for (auto& [key, source] : cells_) {
        const auto row = source.row();
        auto column = source.column();
        if (insert) {
            if (column >= index) column += amount;
        } else {
            if (column >= index && column < deleteEnd) continue;
            if (column >= deleteEnd) column -= amount;
        }
        source.setPosition(row, column);
        shifted.emplace(makeCellKey(row, column), std::move(source));
    }
    cells_ = std::move(shifted);
    dirty_ = true;
}
}

namespace xlpp {
void Worksheet::mergeCells(const std::string& rangeAddress) {
    const auto [first, last] = parseRangeAddress(rangeAddress);
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
    const auto [first, last] = parseRangeAddress(rangeAddress);
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
    return rowDimensions_[row];
}

const RowDimension* Worksheet::tryRowDimension(std::size_t row) const noexcept {
    const auto it = rowDimensions_.find(row);
    return it == rowDimensions_.end() ? nullptr : &it->second;
}

ColumnDimension& Worksheet::columnDimension(std::size_t column) {
    if (column == 0 || column > 16384) throw std::invalid_argument("Column index must be between 1 and 16384");
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
    return it == tables_.end() ? nullptr : &*it;
}
const Table* Worksheet::table(const std::string& name) const noexcept {
    auto it = std::find_if(tables_.begin(), tables_.end(), [&](const Table& value){ return value.name() == name; });
    return it == tables_.end() ? nullptr : &*it;
}

Row::Row(Worksheet& sheet, std::size_t rowNumber) : sheet_(&sheet), rowNumber_(rowNumber) {
    if (rowNumber == 0) throw std::invalid_argument("Row number is 1-based");
}

Cell& Row::cell(std::size_t column) { return sheet_->cell(rowNumber_, column); }

const Cell* Row::tryCell(std::size_t column) const noexcept { return sheet_->tryCell(rowNumber_, column); }

std::vector<Cell*> Row::cells() {
    std::vector<Cell*> result;
    const auto e = sheet_->extents();
    for (std::size_t col = e.minColumn; col <= e.maxColumn; ++col) {
        auto* c = &sheet_->cell(rowNumber_, col);
        if (!c->empty()) result.push_back(c);
    }
    return result;
}

std::vector<CellValue> Row::values() const {
    std::vector<CellValue> result;
    const auto e = sheet_->extents();
    for (std::size_t col = e.minColumn; col <= e.maxColumn; ++col) {
        auto* c = sheet_->tryCell(rowNumber_, col);
        result.push_back(c ? c->value() : CellValue{});
    }
    return result;
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
