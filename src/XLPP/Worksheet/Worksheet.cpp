#include <XLPP/Worksheet/Worksheet.h>
#include "../Internal/WorksheetName.h"
#include <algorithm>
#include <array>
#include <cmath>
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

bool rangesOverlap(const std::pair<xlpp::CellReference, xlpp::CellReference>& a,
                   const std::pair<xlpp::CellReference, xlpp::CellReference>& b) noexcept {
    return a.first.row <= b.second.row && b.first.row <= a.second.row &&
           a.first.column <= b.second.column && b.first.column <= a.second.column;
}

bool asciiCaseEqual(std::string_view a, std::string_view b) noexcept {
    if (a.size() != b.size()) return false;
    for (std::size_t i = 0; i < a.size(); ++i) {
        unsigned char ac = static_cast<unsigned char>(a[i]);
        unsigned char bc = static_cast<unsigned char>(b[i]);
        if (ac >= 'A' && ac <= 'Z') ac = static_cast<unsigned char>(ac - 'A' + 'a');
        if (bc >= 'A' && bc <= 'Z') bc = static_cast<unsigned char>(bc - 'A' + 'a');
        if (ac != bc) return false;
    }
    return true;
}


bool validDataLabelPosition(const std::string& position) {
    if (position.empty()) return true;
    static const std::array<const char*, 9> validPositions{"bestFit", "b", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t"};
    return std::find_if(validPositions.begin(), validPositions.end(), [&](const char* candidate) {
        return position == candidate;
    }) != validPositions.end();
}

bool validChartLineFormat(const xlpp::ChartLineFormat& format) {
    return std::isfinite(format.widthPoints) && format.widthPoints >= 0.0 &&
           (!format.color.present() || !format.color.value.empty());
}

bool validChartFillFormat(const xlpp::ChartFillFormat& format) {
    return !format.color.present() || !format.color.value.empty();
}

bool validChartSeriesCache(const xlpp::ChartSeriesCache& cache) {
    return cache.valid(true);
}

bool validChartMarkerFormat(const xlpp::ChartMarkerFormat& format) {
    if (!format.present) return true;
    if (format.size != 0 && (format.size < 2 || format.size > 72)) return false;
    if (!format.symbol.empty()) {
        static const std::array<const char*, 11> symbols{"circle", "dash", "diamond", "dot", "none", "picture", "plus", "square", "star", "triangle", "x"};
        if (std::find_if(symbols.begin(), symbols.end(), [&](const char* candidate) { return format.symbol == candidate; }) == symbols.end())
            return false;
    }
    return validChartLineFormat(format.line) && validChartFillFormat(format.fill);
}
}

namespace xlpp {
void Worksheet::rename(std::string name) {
    internal::validateWorksheetName(name);
    name_ = std::move(name);
    dirty_ = true;
}

Cell& Worksheet::cell(const std::string& address) {
    const auto ref = CellReference::parse(address);
    return cell(ref.row, ref.column);
}

Cell& Worksheet::cell(std::size_t row, std::size_t column) {
    if (!CellReference::validGridPosition(row, column))
        throw std::out_of_range("Cell coordinate is outside the Excel grid");
    const auto key = makeCellKey(row, column);
    auto [it, inserted] = cells_.try_emplace(key, row, column);
    if (inserted && extentsCacheValid_) {
        if (cells_.size() == 1) extentsCache_ = {row, column, row, column};
        else {
            extentsCache_.minRow = std::min(extentsCache_.minRow, row);
            extentsCache_.minColumn = std::min(extentsCache_.minColumn, column);
            extentsCache_.maxRow = std::max(extentsCache_.maxRow, row);
            extentsCache_.maxColumn = std::max(extentsCache_.maxColumn, column);
        }
    }
    dirty_ = true;
    trackedCellKeys_.insert(key);
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
        if (!CellReference::validGridPosition(row, column)) return nullptr;
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
    if (minRow > maxRow) std::swap(minRow, maxRow);
    if (minColumn > maxColumn) std::swap(minColumn, maxColumn);
    if (!CellReference::validGridPosition(minRow, minColumn) ||
        !CellReference::validGridPosition(maxRow, maxColumn))
        throw std::out_of_range("Range is outside the Excel grid");
    return CellRange(*this, minRow, minColumn, maxRow, maxColumn);
}

std::size_t Worksheet::maxRow() const noexcept {
    return extents().maxRow;
}

std::size_t Worksheet::maxColumn() const noexcept {
    return extents().maxColumn;
}

WorksheetExtents Worksheet::extents() const noexcept {
    if (extentsCacheValid_) return extentsCache_;
    if (cells_.empty()) {
        extentsCache_ = {1, 1, 1, 1};
        extentsCacheValid_ = true;
        return extentsCache_;
    }
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
    extentsCache_ = {minRow, minColumn, maxRowValue, maxColumnValue};
    extentsCacheValid_ = true;
    return extentsCache_;
}

std::string Worksheet::dimensions() const {
    const auto e = extents();
    return CellReference{e.minRow, e.minColumn}.address() + ":" +
           CellReference{e.maxRow, e.maxColumn}.address();
}

void Worksheet::append(const std::vector<CellValue>& values) {
    if (values.size() > CellReference::MaxColumn)
        throw std::out_of_range("Append width exceeds the Excel column limit");
    const auto targetRow = cells_.empty() ? 1 : 1 + (cells_.rbegin()->first >> 20);
    if (targetRow > CellReference::MaxRow)
        throw std::out_of_range("Append would exceed the Excel row limit");

    // append() owns the complete mutation and does not expose a mutable Cell&,
    // so avoid the per-cell std::set insertion performed by cell().  The Cell
    // mutation revision still records every appended value for dependency
    // tracking, while bulk construction stays O(columns) with only the cell map.
    // targetRow is strictly greater than the current maximum row and columns
    // are appended in ascending order.  end() is therefore the exact insertion
    // hint for every cell in the new row, avoiding an O(log N) tree search per
    // value while retaining std::map node stability and the public cells() API.
    for (std::size_t column = 1; column <= values.size(); ++column) {
        const auto key = makeCellKey(targetRow, column);
        auto it = cells_.try_emplace(cells_.end(), key, targetRow, column);
        if (extentsCacheValid_) {
            if (cells_.size() == 1) extentsCache_ = {targetRow, column, targetRow, column};
            else {
                extentsCache_.minRow = std::min(extentsCache_.minRow, targetRow);
                extentsCache_.minColumn = std::min(extentsCache_.minColumn, column);
                extentsCache_.maxRow = std::max(extentsCache_.maxRow, targetRow);
                extentsCache_.maxColumn = std::max(extentsCache_.maxColumn, column);
            }
        }
        it->second.setValue(values[column - 1]);
    }
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


void Worksheet::mergeCells(const std::string& rangeAddress) {
    const auto [first, last] = parseRangeAddress(rangeAddress);
    const auto canonical = CellReference{first.row, first.column}.address() + ":" +
                           CellReference{last.row, last.column}.address();
    if (first.row == last.row && first.column == last.column)
        throw std::invalid_argument("Merged range must contain at least two cells");
    const auto candidate = std::pair{first, last};
    for (const auto& existing : mergedRanges_) {
        const auto parsedExisting = parseRangeAddress(existing);
        if (existing == canonical) return;
        if (rangesOverlap(candidate, parsedExisting))
            throw std::invalid_argument("Merged ranges cannot overlap: " + canonical + " intersects " + existing);
    }
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
    if (row == 0 || row > CellReference::MaxRow) throw std::invalid_argument("Row index must be between 1 and 1048576");
    dirty_ = true;
    return rowDimensions_[row];
}

const RowDimension* Worksheet::tryRowDimension(std::size_t row) const noexcept {
    if (row == 0 || row > CellReference::MaxRow) return nullptr;
    const auto it = rowDimensions_.find(row);
    return it == rowDimensions_.end() ? nullptr : &it->second;
}

ColumnDimension& Worksheet::columnDimension(std::size_t column) {
    if (column == 0 || column > 16384) throw std::invalid_argument("Column index must be between 1 and 16384");
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
    if (name.empty()) throw std::invalid_argument("Table name cannot be empty");
    if (std::any_of(tables_.begin(), tables_.end(), [&](const Table& value){ return asciiCaseEqual(value.name(), name); }))
        throw std::invalid_argument("Duplicate table name: " + name);
    const auto candidate = parseRangeAddress(reference);
    for (const auto& existing : tables_) {
        if (rangesOverlap(candidate, parseRangeAddress(existing.reference())))
            throw std::invalid_argument("Table ranges cannot overlap on a worksheet: " + reference + " intersects " + existing.reference());
    }
    tables_.emplace_back(std::move(name), std::move(reference));
    dirty_ = true;
    return tables_.back();
}
Table* Worksheet::table(const std::string& name) noexcept {
    auto it = std::find_if(tables_.begin(), tables_.end(), [&](const Table& value){ return asciiCaseEqual(value.name(), name); });
    if (it != tables_.end()) dirty_ = true;
    return it == tables_.end() ? nullptr : &*it;
}
const Table* Worksheet::table(const std::string& name) const noexcept {
    auto it = std::find_if(tables_.begin(), tables_.end(), [&](const Table& value){ return asciiCaseEqual(value.name(), name); });
    return it == tables_.end() ? nullptr : &*it;
}

Row::Row(Worksheet& sheet, std::size_t rowNumber) : sheet_(&sheet), rowNumber_(rowNumber) {
    if (rowNumber == 0 || rowNumber > CellReference::MaxRow)
        throw std::invalid_argument("Row number must be between 1 and 1048576");
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
const Image* Worksheet::imageByStableId(const std::string& stableId) const noexcept {
    const auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    return it == images_.end() ? nullptr : &*it;
}

Worksheet::ImportedImageEdit& Worksheet::ensureImportedImageEdit(const Image& image) {
    const auto existing = std::find_if(importedImageEdits_.begin(), importedImageEdits_.end(), [&](const ImportedImageEdit& edit) {
        return edit.stableId == image.stableId_;
    });
    if (existing != importedImageEdits_.end()) return *existing;
    ImportedImageEdit edit;
    edit.stableId = image.stableId_;
    edit.sourceDrawingPart = image.sourceDrawingPart_;
    edit.sourceMediaPart = image.sourceMediaPart_;
    edit.sourceRelationshipId = image.sourceRelationshipId_;
    edit.originalAnchor = image.anchorInfo_;
    edit.anchor = image.anchorInfo_;
    importedImageEdits_.push_back(std::move(edit));
    return importedImageEdits_.back();
}

bool Worksheet::moveImage(const std::string& stableId, const std::string& anchor) {
    auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end() || it->anchorInfo_.type == DrawingAnchorType::Absolute) return false;
    const auto ref = CellReference::parse(anchor);
    const auto oldRow = static_cast<long long>(it->anchorInfo_.from.row);
    const auto oldColumn = static_cast<long long>(it->anchorInfo_.from.column);
    const auto rowDelta = static_cast<long long>(ref.row) - oldRow;
    const auto columnDelta = static_cast<long long>(ref.column) - oldColumn;
    auto updated = it->anchorInfo_;
    updated.from.row = ref.row;
    updated.from.column = ref.column;
    if (updated.type == DrawingAnchorType::TwoCell) {
        const auto newToRow = static_cast<long long>(updated.to.row) + rowDelta;
        const auto newToColumn = static_cast<long long>(updated.to.column) + columnDelta;
        if (newToRow < 1 || newToColumn < 1) return false;
        updated.to.row = static_cast<std::size_t>(newToRow);
        updated.to.column = static_cast<std::size_t>(newToColumn);
    }
    auto& edit = ensureImportedImageEdit(*it);
    edit.anchor = updated;
    edit.moved = true;
    it->anchor_ = CellReference{ref.row, ref.column}.address();
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::moveImageAbsolute(const std::string& stableId, long long xEmu, long long yEmu) {
    auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end() || it->anchorInfo_.type != DrawingAnchorType::Absolute || xEmu < 0 || yEmu < 0) return false;
    auto updated = it->anchorInfo_;
    updated.xEmu = xEmu;
    updated.yEmu = yEmu;
    auto& edit = ensureImportedImageEdit(*it);
    edit.anchor = updated;
    edit.moved = true;
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::resizeImage(const std::string& stableId, double widthPixels, double heightPixels) {
    if (!(widthPixels > 0.0) || !(heightPixels > 0.0)) return false;
    auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end()) return false;
    auto updated = it->anchorInfo_;
    updated.widthEmu = static_cast<long long>(std::llround(widthPixels * 9525.0));
    updated.heightEmu = static_cast<long long>(std::llround(heightPixels * 9525.0));
    if (updated.type == DrawingAnchorType::TwoCell) {
        const auto resizeTerminalMarker = [](std::size_t fromIndex, long long fromOffset,
                                             std::size_t oldToIndex, long long oldToOffset,
                                             long long oldExtent, long long newExtent,
                                             std::size_t& newToIndex, long long& newToOffset) {
            if (oldExtent <= 0 || newExtent <= 0 || oldToIndex < fromIndex) return false;
            const auto span = oldToIndex - fromIndex;
            if (span == 0) {
                // The original image already ends in the starting cell. A
                // smaller/equal resize can safely retain that cell without
                // needing workbook font/column metrics.
                if (newExtent > oldExtent) return false;
                newToIndex = fromIndex;
                newToOffset = fromOffset + newExtent;
                return true;
            }
            const long double averageCellExtent =
                static_cast<long double>(oldExtent + fromOffset - oldToOffset) / static_cast<long double>(span);
            if (!(averageCellExtent > 0.0L)) return false;
            const long double terminal = static_cast<long double>(fromOffset + newExtent);
            auto cells = static_cast<long long>(std::floor(terminal / averageCellExtent));
            if (cells < 0) cells = 0;
            auto offset = static_cast<long long>(std::llround(terminal - static_cast<long double>(cells) * averageCellExtent));
            const auto roundedCell = static_cast<long long>(std::llround(averageCellExtent));
            if (roundedCell > 0 && offset >= roundedCell) { ++cells; offset = 0; }
            newToIndex = fromIndex + static_cast<std::size_t>(cells);
            newToOffset = std::max<long long>(0, offset);
            return true;
        };
        std::size_t toColumn = updated.to.column;
        std::size_t toRow = updated.to.row;
        long long toColumnOffset = updated.to.columnOffsetEmu;
        long long toRowOffset = updated.to.rowOffsetEmu;
        const bool columnOk = resizeTerminalMarker(updated.from.column, updated.from.columnOffsetEmu,
                                                   updated.to.column, updated.to.columnOffsetEmu,
                                                   it->anchorInfo_.widthEmu, updated.widthEmu,
                                                   toColumn, toColumnOffset);
        const bool rowOk = resizeTerminalMarker(updated.from.row, updated.from.rowOffsetEmu,
                                                updated.to.row, updated.to.rowOffsetEmu,
                                                it->anchorInfo_.heightEmu, updated.heightEmu,
                                                toRow, toRowOffset);
        if (!columnOk || !rowOk) return false;
        updated.to.column = toColumn;
        updated.to.columnOffsetEmu = toColumnOffset;
        updated.to.row = toRow;
        updated.to.rowOffsetEmu = toRowOffset;
    }
    auto& edit = ensureImportedImageEdit(*it);
    edit.anchor = updated;
    edit.resized = true;
    it->widthPixels_ = widthPixels;
    it->heightPixels_ = heightPixels;
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::replaceImage(const std::string& stableId, Image replacement) {
    if (replacement.extension_ != "png" && replacement.extension_ != "jpg" && replacement.extension_ != "jpeg") return false;
    auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end() || replacement.bytes_.empty()) return false;
    if (replacement.extension_ == "jpeg") replacement.extension_ = "jpg";
    auto& edit = ensureImportedImageEdit(*it);
    edit.replaced = true;
    edit.replacementBytes = replacement.bytes_;
    edit.replacementExtension = replacement.extension_;
    it->bytes_ = replacement.bytes_;
    it->extension_ = replacement.extension_;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::replaceImage(const std::string& stableId, const std::filesystem::path& path) {
    const auto* current = imageByStableId(stableId);
    if (!current) return false;
    return replaceImage(stableId, Image::fromFile(path, current->anchor()));
}

bool Worksheet::removeImage(const std::string& stableId) {
    const auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end()) return false;
    auto& edit = ensureImportedImageEdit(*it);
    edit.removed = true;
    const auto index = static_cast<std::size_t>(std::distance(images_.begin(), it));
    images_.erase(it);
    if (index < loadedImageCount_ && loadedImageCount_ > 0) --loadedImageCount_;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

const Chart* Worksheet::chartByStableId(const std::string& stableId) const noexcept {
    const auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    return it == charts_.end() ? nullptr : &*it;
}

Worksheet::ImportedChartEdit& Worksheet::ensureImportedChartEdit(const Chart& chart) {
    const auto existing = std::find_if(importedChartEdits_.begin(), importedChartEdits_.end(), [&](const ImportedChartEdit& edit) {
        return edit.stableId == chart.stableId_;
    });
    if (existing != importedChartEdits_.end()) return *existing;
    ImportedChartEdit edit;
    edit.stableId = chart.stableId_;
    edit.sourceDrawingPart = chart.sourceDrawingPart_;
    edit.sourceChartPart = chart.sourceChartPart_;
    edit.sourceRelationshipId = chart.sourceRelationshipId_;
    edit.chartType = chart.type_;
    edit.primaryXAxisId = chart.primaryXAxisId_;
    edit.primaryYAxisId = chart.primaryYAxisId_;
    edit.originalAnchor = chart.anchorInfo_;
    edit.anchor = chart.anchorInfo_;
    importedChartEdits_.push_back(std::move(edit));
    return importedChartEdits_.back();
}

bool Worksheet::moveChart(const std::string& stableId, const std::string& anchor) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || it->anchorInfo_.type == DrawingAnchorType::Absolute) return false;
    const auto ref = CellReference::parse(anchor);
    const auto rowDelta = static_cast<long long>(ref.row) - static_cast<long long>(it->anchorInfo_.from.row);
    const auto columnDelta = static_cast<long long>(ref.column) - static_cast<long long>(it->anchorInfo_.from.column);
    auto updated = it->anchorInfo_;
    updated.from.row = ref.row;
    updated.from.column = ref.column;
    if (updated.type == DrawingAnchorType::TwoCell) {
        const auto newToRow = static_cast<long long>(updated.to.row) + rowDelta;
        const auto newToColumn = static_cast<long long>(updated.to.column) + columnDelta;
        if (newToRow < 1 || newToColumn < 1) return false;
        updated.to.row = static_cast<std::size_t>(newToRow);
        updated.to.column = static_cast<std::size_t>(newToColumn);
    }
    auto& edit = ensureImportedChartEdit(*it);
    edit.anchor = updated;
    edit.moved = true;
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::moveChartAbsolute(const std::string& stableId, long long xEmu, long long yEmu) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || it->anchorInfo_.type != DrawingAnchorType::Absolute || xEmu < 0 || yEmu < 0) return false;
    auto updated = it->anchorInfo_;
    updated.xEmu = xEmu;
    updated.yEmu = yEmu;
    auto& edit = ensureImportedChartEdit(*it);
    edit.anchor = updated;
    edit.moved = true;
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::resizeChart(const std::string& stableId, double widthPixels, double heightPixels) {
    if (!(widthPixels > 0.0) || !(heightPixels > 0.0)) return false;
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    auto updated = it->anchorInfo_;
    const auto oldWidth = updated.widthEmu;
    const auto oldHeight = updated.heightEmu;
    updated.widthEmu = static_cast<long long>(std::llround(widthPixels * 9525.0));
    updated.heightEmu = static_cast<long long>(std::llround(heightPixels * 9525.0));
    if (updated.type == DrawingAnchorType::TwoCell) {
        const auto resizeTerminalMarker = [](std::size_t fromIndex, long long fromOffset,
                                             std::size_t oldToIndex, long long oldToOffset,
                                             long long oldExtent, long long newExtent,
                                             std::size_t& newToIndex, long long& newToOffset) {
            if (oldExtent <= 0 || newExtent <= 0 || oldToIndex < fromIndex) return false;
            const auto span = oldToIndex - fromIndex;
            if (span == 0) {
                if (newExtent > oldExtent) return false;
                newToIndex = fromIndex;
                newToOffset = fromOffset + newExtent;
                return true;
            }
            const long double averageCellExtent =
                static_cast<long double>(oldExtent + fromOffset - oldToOffset) / static_cast<long double>(span);
            if (!(averageCellExtent > 0.0L)) return false;
            const long double terminal = static_cast<long double>(fromOffset + newExtent);
            auto cells = static_cast<long long>(std::floor(terminal / averageCellExtent));
            if (cells < 0) cells = 0;
            auto offset = static_cast<long long>(std::llround(terminal - static_cast<long double>(cells) * averageCellExtent));
            const auto roundedCell = static_cast<long long>(std::llround(averageCellExtent));
            if (roundedCell > 0 && offset >= roundedCell) { ++cells; offset = 0; }
            newToIndex = fromIndex + static_cast<std::size_t>(cells);
            newToOffset = std::max<long long>(0, offset);
            return true;
        };
        std::size_t toColumn = updated.to.column;
        std::size_t toRow = updated.to.row;
        long long toColumnOffset = updated.to.columnOffsetEmu;
        long long toRowOffset = updated.to.rowOffsetEmu;
        if (!resizeTerminalMarker(updated.from.column, updated.from.columnOffsetEmu,
                                  updated.to.column, updated.to.columnOffsetEmu, oldWidth, updated.widthEmu,
                                  toColumn, toColumnOffset) ||
            !resizeTerminalMarker(updated.from.row, updated.from.rowOffsetEmu,
                                  updated.to.row, updated.to.rowOffsetEmu, oldHeight, updated.heightEmu,
                                  toRow, toRowOffset)) return false;
        updated.to.column = toColumn;
        updated.to.columnOffsetEmu = toColumnOffset;
        updated.to.row = toRow;
        updated.to.rowOffsetEmu = toRowOffset;
    }
    auto& edit = ensureImportedChartEdit(*it);
    edit.anchor = updated;
    edit.resized = true;
    it->anchorInfo_ = updated;
    it->width_ = std::max(1, static_cast<int>(std::llround(widthPixels)));
    it->height_ = std::max(1, static_cast<int>(std::llround(heightPixels)));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartTitle(const std::string& stableId, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    edit.titleChanged = true;
    edit.title = title;
    it->title_ = std::move(title);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartStyle(const std::string& stableId, std::string style) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || style.empty()) return false;
    try {
        const auto numeric = std::stoul(style);
        if (numeric == 0 || numeric > 48) return false;
    } catch (...) { return false; }
    auto& edit = ensureImportedChartEdit(*it);
    edit.styleChanged = true;
    edit.style = style;
    it->style_ = std::move(style);
    dirty_ = true; drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartTitleRichText(const std::string& stableId, ChartRichText richText) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || !richText.present || richText.runs.empty()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    edit.titleRichTextChanged = true;
    edit.titleRichText = richText;
    it->titleRichText_ = std::move(richText);
    it->title_ = it->titleRichText_.plainText();
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartXAxisTitle(const std::string& stableId, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    if (it->primaryXAxisId_ != 0) return setChartAxisTitle(stableId, it->primaryXAxisId_, std::move(title));
    auto& edit = ensureImportedChartEdit(*it);
    edit.xAxisTitleChanged = true;
    edit.xAxisTitle = title;
    it->xAxisTitle_ = std::move(title);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartYAxisTitle(const std::string& stableId, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    if (it->primaryYAxisId_ != 0) return setChartAxisTitle(stableId, it->primaryYAxisId_, std::move(title));
    auto& edit = ensureImportedChartEdit(*it);
    edit.yAxisTitleChanged = true;
    edit.yAxisTitle = title;
    it->yAxisTitle_ = std::move(title);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartAxisTitle(const std::string& stableId, std::uint64_t axisId, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || axisId == 0) return false;
    auto axis = std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate) {
        return candidate.id == axisId;
    });
    if (axis == it->axes_.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.axisTitleEdits.begin(), edit.axisTitleEdits.end(), [&](const auto& candidate) {
        return candidate.axisId == axisId;
    });
    if (existing == edit.axisTitleEdits.end()) {
        ImportedChartEdit::AxisTitleEdit axisEdit;
        axisEdit.axisId = axisId;
        axisEdit.title = title;
        edit.axisTitleEdits.push_back(std::move(axisEdit));
    } else {
        existing->title = title;
    }
    axis->title = title;
    if (axisId == it->primaryXAxisId_) it->xAxisTitle_ = title;
    if (axisId == it->primaryYAxisId_) it->yAxisTitle_ = title;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartAxisTitleRichText(const std::string& stableId, std::uint64_t axisId, ChartRichText richText) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || axisId == 0 || !richText.present || richText.runs.empty()) return false;
    auto axis = std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate) { return candidate.id == axisId; });
    if (axis == it->axes_.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    auto existing = std::find_if(edit.axisRichTitleEdits.begin(), edit.axisRichTitleEdits.end(), [&](const auto& candidate) { return candidate.axisId == axisId; });
    if (existing == edit.axisRichTitleEdits.end()) edit.axisRichTitleEdits.push_back({axisId, richText}); else existing->richText = richText;
    axis->titleRichText = richText;
    axis->title = richText.plainText();
    if (axisId == it->primaryXAxisId_) it->xAxisTitle_ = axis->title;
    if (axisId == it->primaryYAxisId_) it->yAxisTitle_ = axis->title;
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartAxisNumberFormat(const std::string& stableId, std::uint64_t axisId, std::string formatCode, bool sourceLinked) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || axisId == 0 || formatCode.empty()) return false;
    auto axis = std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate) { return candidate.id == axisId; });
    if (axis == it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind = ImportedChartEdit::AxisFormatEdit::Kind::NumberFormat; e.axisId = axisId; e.value1 = formatCode; e.flag = sourceLinked;
    auto& edit = ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e));
    axis->numberFormat = std::move(formatCode); axis->numberFormatSourceLinked = sourceLinked;
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartAxisTicks(const std::string& stableId, std::uint64_t axisId, std::string majorTickMark, std::string minorTickMark, std::string tickLabelPosition) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || axisId == 0) return false;
    auto axis = std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate) { return candidate.id == axisId; });
    if (axis == it->axes_.end()) return false;
    const auto validTick=[](const std::string& v){ return v.empty() || v=="none" || v=="in" || v=="out" || v=="cross"; };
    const auto validPos=[](const std::string& v){ return v.empty() || v=="high" || v=="low" || v=="nextTo" || v=="none"; };
    if (!validTick(majorTickMark) || !validTick(minorTickMark) || !validPos(tickLabelPosition)) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Ticks; e.axisId=axisId; e.value1=majorTickMark; e.value2=minorTickMark; e.value3=tickLabelPosition;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e));
    axis->majorTickMark=std::move(majorTickMark); axis->minorTickMark=std::move(minorTickMark); axis->tickLabelPosition=std::move(tickLabelPosition);
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisUnits(const std::string& stableId, std::uint64_t axisId, double majorUnit, double minorUnit) {
    if (!(majorUnit > 0.0) || minorUnit < 0.0) return false;
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if (it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Units; e.axisId=axisId; e.number1=majorUnit; e.number2=minorUnit;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e));
    axis->hasMajorUnit=true; axis->majorUnit=majorUnit; axis->hasMinorUnit=minorUnit>0.0; axis->minorUnit=minorUnit;
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisScaling(const std::string& stableId, std::uint64_t axisId, ChartAxisScaling scaling) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; });
    if(axis==it->axes_.end()) return false;
    if (scaling.hasMinimum && scaling.hasMaximum && !(scaling.minimum < scaling.maximum)) return false;
    if (scaling.hasLogBase && (!(scaling.logBase >= 2.0 && scaling.logBase <= 1000.0) ||
        (scaling.hasMinimum && scaling.minimum <= 0.0) || (scaling.hasMaximum && scaling.maximum <= 0.0))) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Scaling; e.axisId=axisId; e.scaling=scaling;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->scaling=std::move(scaling);
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisCrossing(const std::string& stableId, std::uint64_t axisId, std::string crosses, std::string crossBetween) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    const auto validCross=[](const std::string& v){ return v.empty() || v=="autoZero" || v=="max" || v=="min"; };
    const auto validBetween=[](const std::string& v){ return v.empty() || v=="between" || v=="midCat"; };
    if(!validCross(crosses) || !validBetween(crossBetween)) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Crossing; e.axisId=axisId; e.value1=crosses; e.value2=crossBetween;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->crosses=std::move(crosses); axis->crossBetween=std::move(crossBetween); axis->hasCrossesAt=false; axis->crossesAt=0.0;
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisCrossesAt(const std::string& stableId, std::uint64_t axisId, double crossesAt) {
    if (!std::isfinite(crossesAt)) return false;
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::CrossesAt; e.axisId=axisId; e.number1=crossesAt;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->hasCrossesAt=true; axis->crossesAt=crossesAt; axis->crosses.clear();
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::clearChartAxisCrossesAt(const std::string& stableId, std::uint64_t axisId) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::ClearCrossesAt; e.axisId=axisId;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->hasCrossesAt=false; axis->crossesAt=0.0;
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisDisplayUnits(const std::string& stableId, std::uint64_t axisId, ChartDisplayUnits units) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0 || !units.present) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; });
    if(axis==it->axes_.end() || axis->kind!=Chart::AxisKind::Value) return false;
    static const std::array<const char*, 9> validUnits{"hundreds","thousands","tenThousands","hundredThousands","millions","tenMillions","hundredMillions","billions","trillions"};
    if (units.hasCustomUnit) {
        if (!(units.customUnit > 0.0) || !std::isfinite(units.customUnit) || !units.builtInUnit.empty()) return false;
    } else if (std::find_if(validUnits.begin(), validUnits.end(), [&](const char* value){ return units.builtInUnit==value; })==validUnits.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::DisplayUnits; e.axisId=axisId; e.displayUnits=units;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->displayUnits=std::move(units);
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::clearChartAxisDisplayUnits(const std::string& stableId, std::uint64_t axisId) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::ClearDisplayUnits; e.axisId=axisId;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->displayUnits=ChartDisplayUnits{};
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisLineFormat(const std::string& stableId, std::uint64_t axisId, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||axisId==0||!format.present) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Line; e.axisId=axisId; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->lineFormat=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisGridlineFormat(const std::string& stableId, std::uint64_t axisId, bool major, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||axisId==0||!format.present) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=major?ImportedChartEdit::AxisFormatEdit::Kind::MajorGridline:ImportedChartEdit::AxisFormatEdit::Kind::MinorGridline; e.axisId=axisId; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); if(major){ axis->hasMajorGridlines=true; axis->majorGridlineFormat=std::move(format); } else { axis->hasMinorGridlines=true; axis->minorGridlineFormat=std::move(format); } dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartAxisGridlines(const std::string& stableId, std::uint64_t axisId, bool major) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=major?ImportedChartEdit::AxisFormatEdit::Kind::RemoveMajorGridline:ImportedChartEdit::AxisFormatEdit::Kind::RemoveMinorGridline; e.axisId=axisId;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e));
    if(major){ axis->hasMajorGridlines=false; axis->majorGridlineFormat=ChartLineFormat{}; } else { axis->hasMinorGridlines=false; axis->minorGridlineFormat=ChartLineFormat{}; }
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAreaLineFormat(const std::string& stableId, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false;
    ImportedChartEdit::AreaFormatEdit e; e.owner=ImportedChartEdit::AreaFormatEdit::Owner::ChartArea; e.kind=ImportedChartEdit::AreaFormatEdit::Kind::Line; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.areaFormatEdits.push_back(std::move(e)); it->chartAreaLineFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAreaFillFormat(const std::string& stableId, ChartFillFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false;
    ImportedChartEdit::AreaFormatEdit e; e.owner=ImportedChartEdit::AreaFormatEdit::Owner::ChartArea; e.kind=ImportedChartEdit::AreaFormatEdit::Kind::Fill; e.fill=format; auto& edit=ensureImportedChartEdit(*it); edit.areaFormatEdits.push_back(std::move(e)); it->chartAreaFillFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotAreaLineFormat(const std::string& stableId, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false;
    ImportedChartEdit::AreaFormatEdit e; e.owner=ImportedChartEdit::AreaFormatEdit::Owner::PlotArea; e.kind=ImportedChartEdit::AreaFormatEdit::Kind::Line; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.areaFormatEdits.push_back(std::move(e)); it->plotAreaLineFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotAreaFillFormat(const std::string& stableId, ChartFillFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false;
    ImportedChartEdit::AreaFormatEdit e; e.owner=ImportedChartEdit::AreaFormatEdit::Owner::PlotArea; e.kind=ImportedChartEdit::AreaFormatEdit::Kind::Fill; e.fill=format; auto& edit=ensureImportedChartEdit(*it); edit.areaFormatEdits.push_back(std::move(e)); it->plotAreaFillFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotAreaLayout(const std::string& stableId, ChartManualLayout layout) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!layout.present) return false;
    ImportedChartEdit::LayoutEdit e; e.owner=ImportedChartEdit::LayoutEdit::Owner::PlotArea; e.layout=layout; auto& edit=ensureImportedChartEdit(*it); edit.layoutEdits.push_back(std::move(e)); it->plotAreaLayout_=std::move(layout); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartView3D(const std::string& stableId, ChartView3D view) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||!view.present) return false;
    if(view.hasRotationX && (view.rotationX < -90 || view.rotationX > 90)) return false;
    if(view.hasRotationY && (view.rotationY < 0 || view.rotationY > 360)) return false;
    if(view.hasHeightPercent && (view.heightPercent < 5 || view.heightPercent > 500)) return false;
    if(view.hasDepthPercent && (view.depthPercent < 20 || view.depthPercent > 2000)) return false;
    if(view.hasPerspective && (view.perspective < 0 || view.perspective > 240)) return false;
    ImportedChartEdit::View3DEdit e; e.view=view; auto& edit=ensureImportedChartEdit(*it); edit.view3DEdits.push_back(std::move(e));
    it->view3D_=std::move(view); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartFloorFormat(const std::string& stableId, ChartWallFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||!format.present) return false;
    if(format.hasThickness && format.thickness < 0) return false;
    ImportedChartEdit::WallFormatEdit e; e.owner=ImportedChartEdit::WallFormatEdit::Owner::Floor; e.format=format; auto& edit=ensureImportedChartEdit(*it); edit.wallFormatEdits.push_back(std::move(e));
    it->floorFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartSideWallFormat(const std::string& stableId, ChartWallFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||!format.present) return false;
    if(format.hasThickness && format.thickness < 0) return false;
    ImportedChartEdit::WallFormatEdit e; e.owner=ImportedChartEdit::WallFormatEdit::Owner::SideWall; e.format=format; auto& edit=ensureImportedChartEdit(*it); edit.wallFormatEdits.push_back(std::move(e));
    it->sideWallFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartBackWallFormat(const std::string& stableId, ChartWallFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||!format.present) return false;
    if(format.hasThickness && format.thickness < 0) return false;
    ImportedChartEdit::WallFormatEdit e; e.owner=ImportedChartEdit::WallFormatEdit::Owner::BackWall; e.format=format; auto& edit=ensureImportedChartEdit(*it); edit.wallFormatEdits.push_back(std::move(e));
    it->backWallFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartDataTable(const std::string& stableId, ChartDataTable table) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()) return false;
    table.present=true; ImportedChartEdit::DataTableEdit e; e.table=table; auto& edit=ensureImportedChartEdit(*it); edit.dataTableEdits.push_back(std::move(e));
    it->dataTable_=std::move(table); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartDataTable(const std::string& stableId) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()) return false;
    ImportedChartEdit::DataTableEdit e; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.dataTableEdits.push_back(std::move(e));
    it->dataTable_=ChartDataTable{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotDropLines(const std::string& stableId, std::size_t plotIndex, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::DropLines; e.plotIndex=plotIndex; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasDropLines=true; it->plots_[plotIndex].dropLinesFormat=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartPlotDropLines(const std::string& stableId, std::size_t plotIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::DropLines; e.plotIndex=plotIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasDropLines=false; it->plots_[plotIndex].dropLinesFormat=ChartLineFormat{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotHighLowLines(const std::string& stableId, std::size_t plotIndex, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::HighLowLines; e.plotIndex=plotIndex; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasHighLowLines=true; it->plots_[plotIndex].highLowLinesFormat=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartPlotHighLowLines(const std::string& stableId, std::size_t plotIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::HighLowLines; e.plotIndex=plotIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasHighLowLines=false; it->plots_[plotIndex].highLowLinesFormat=ChartLineFormat{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotUpDownBars(const std::string& stableId, std::size_t plotIndex, ChartUpDownBars bars) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()||bars.gapWidth<0||bars.gapWidth>500) return false;
    bars.present=true; ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::UpDownBars; e.plotIndex=plotIndex; e.upDownBars=bars; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].upDownBars=std::move(bars); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartPlotUpDownBars(const std::string& stableId, std::size_t plotIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::UpDownBars; e.plotIndex=plotIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].upDownBars=ChartUpDownBars{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotFirstSliceAngle(const std::string& stableId, std::size_t plotIndex, int degrees) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||plotIndex>=it->plots_.size()||degrees<0||degrees>360) return false;
    const auto type=it->plots_[plotIndex].type; if(type!=Chart::Type::Pie && type!=Chart::Type::Doughnut) return false;
    ImportedChartEdit::PlotTypeSpecificEdit e; e.kind=ImportedChartEdit::PlotTypeSpecificEdit::Kind::FirstSliceAngle; e.plotIndex=plotIndex; e.integerValue=degrees;
    auto& edit=ensureImportedChartEdit(*it); edit.plotTypeSpecificEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasFirstSliceAngle=true; it->plots_[plotIndex].firstSliceAngle=degrees; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotDoughnutHoleSize(const std::string& stableId, std::size_t plotIndex, int percent) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||plotIndex>=it->plots_.size()||percent<10||percent>90||it->plots_[plotIndex].type!=Chart::Type::Doughnut) return false;
    ImportedChartEdit::PlotTypeSpecificEdit e; e.kind=ImportedChartEdit::PlotTypeSpecificEdit::Kind::DoughnutHoleSize; e.plotIndex=plotIndex; e.integerValue=percent;
    auto& edit=ensureImportedChartEdit(*it); edit.plotTypeSpecificEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasHoleSize=true; it->plots_[plotIndex].holeSize=percent; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotRadarStyle(const std::string& stableId, std::size_t plotIndex, std::string style) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||plotIndex>=it->plots_.size()||it->plots_[plotIndex].type!=Chart::Type::Radar) return false;
    if(style!="standard"&&style!="marker"&&style!="filled") return false;
    ImportedChartEdit::PlotTypeSpecificEdit e; e.kind=ImportedChartEdit::PlotTypeSpecificEdit::Kind::RadarStyle; e.plotIndex=plotIndex; e.textValue=style;
    auto& edit=ensureImportedChartEdit(*it); edit.plotTypeSpecificEdits.push_back(std::move(e));
    it->plots_[plotIndex].radarStyle=std::move(style); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotProjectedPieOptions(const std::string& stableId, std::size_t plotIndex, ChartProjectedPieOptions options) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    const auto type=it->plots_[plotIndex].type; if(type!=Chart::Type::PieOfPie && type!=Chart::Type::BarOfPie) return false;
    if(options.ofPieType!="pie"&&options.ofPieType!="bar") return false;
    if(options.gapWidth<0||options.gapWidth>500||options.secondPlotSize<5||options.secondPlotSize>200) return false;
    if(options.splitType!="auto"&&options.splitType!="cust"&&options.splitType!="percent"&&options.splitType!="pos"&&options.splitType!="val") return false;
    if(std::any_of(options.customSplitPoints.begin(), options.customSplitPoints.end(), [](int v){ return v<0; })) return false;
    options.present=true; options.ofPieType=(type==Chart::Type::BarOfPie?"bar":"pie");
    ImportedChartEdit::PlotTypeSpecificEdit e; e.kind=ImportedChartEdit::PlotTypeSpecificEdit::Kind::ProjectedPie; e.plotIndex=plotIndex; e.projectedPie=options;
    auto& edit=ensureImportedChartEdit(*it); edit.plotTypeSpecificEdits.push_back(std::move(e));
    it->plots_[plotIndex].projectedPie=std::move(options); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotLeaderLineFormat(const std::string& stableId, std::size_t plotIndex, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()||!format.present) return false;
    ImportedChartEdit::LeaderLineEdit e; e.plotLevel=true; e.ownerIndex=plotIndex; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.leaderLineEdits.push_back(std::move(e));
    auto& labels=it->plots_[plotIndex].dataLabels; labels.present=true; labels.showLeaderLines=true; labels.hasLeaderLines=true; labels.leaderLineFormat=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartPlotLeaderLines(const std::string& stableId, std::size_t plotIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::LeaderLineEdit e; e.plotLevel=true; e.ownerIndex=plotIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.leaderLineEdits.push_back(std::move(e));
    auto& labels=it->plots_[plotIndex].dataLabels; labels.showLeaderLines=false; labels.hasLeaderLines=false; labels.leaderLineFormat=ChartLineFormat{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartSeriesLeaderLineFormat(const std::string& stableId, std::size_t seriesIndex, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||seriesIndex>=it->series_.size()||!format.present) return false;
    ImportedChartEdit::LeaderLineEdit e; e.ownerIndex=seriesIndex; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.leaderLineEdits.push_back(std::move(e));
    auto labels=it->series_[seriesIndex].dataLabels(); labels.present=true; labels.showLeaderLines=true; labels.hasLeaderLines=true; labels.leaderLineFormat=std::move(format); it->series_[seriesIndex].setDataLabels(std::move(labels)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartSeriesLeaderLines(const std::string& stableId, std::size_t seriesIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||seriesIndex>=it->series_.size()) return false;
    ImportedChartEdit::LeaderLineEdit e; e.ownerIndex=seriesIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.leaderLineEdits.push_back(std::move(e));
    auto labels=it->series_[seriesIndex].dataLabels(); labels.showLeaderLines=false; labels.hasLeaderLines=false; labels.leaderLineFormat=ChartLineFormat{}; it->series_[seriesIndex].setDataLabels(std::move(labels)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartLegend(const std::string& stableId, bool show, std::string position) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    if (show) {
        static const std::array<const char*, 5> validPositions{"l", "r", "t", "b", "tr"};
        if (std::find_if(validPositions.begin(), validPositions.end(), [&](const char* candidate) {
                return position == candidate;
            }) == validPositions.end()) return false;
    }
    auto& edit = ensureImportedChartEdit(*it);
    edit.legendChanged = true;
    edit.showLegend = show;
    edit.legendPosition = position;
    it->showLegend_ = show;
    if (show) it->legendPosition_ = std::move(position);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartLegendLayout(const std::string& stableId, ChartManualLayout layout) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!layout.present) return false;
    ImportedChartEdit::LayoutEdit e; e.owner=ImportedChartEdit::LayoutEdit::Owner::Legend; e.layout=layout; auto& edit=ensureImportedChartEdit(*it); edit.layoutEdits.push_back(std::move(e)); it->legendFormat_.present=true; it->legendFormat_.layout=std::move(layout); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartLegendOverlay(const std::string& stableId, bool overlay) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()) return false; auto& edit=ensureImportedChartEdit(*it); edit.legendOverlayChanged=true; edit.legendOverlay=overlay; it->legendFormat_.present=true; it->legendFormat_.overlay=overlay; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartLegendLineFormat(const std::string& stableId, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false; auto& edit=ensureImportedChartEdit(*it); edit.legendLineFormatChanged=true; edit.legendLineFormat=format; it->legendFormat_.present=true; it->legendFormat_.line=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartLegendFillFormat(const std::string& stableId, ChartFillFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false; auto& edit=ensureImportedChartEdit(*it); edit.legendFillFormatChanged=true; edit.legendFillFormat=format; it->legendFormat_.present=true; it->legendFormat_.fill=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartSeriesTitle(const std::string& stableId, std::size_t seriesIndex, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.seriesTitleEdits.begin(), edit.seriesTitleEdits.end(), [&](const auto& candidate) {
        return candidate.seriesIndex == seriesIndex;
    });
    if (existing == edit.seriesTitleEdits.end()) {
        ImportedChartEdit::SeriesTitleEdit seriesEdit;
        seriesEdit.seriesIndex = seriesIndex;
        seriesEdit.title = title;
        edit.seriesTitleEdits.push_back(std::move(seriesEdit));
    } else {
        existing->title = title;
    }
    it->series_[seriesIndex].setTitle(std::move(title));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesReferences(const std::string& stableId, std::size_t seriesIndex,
                                          std::string categoriesReference, std::string valuesReference) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        categoriesReference.empty() || valuesReference.empty()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.seriesReferenceEdits.begin(), edit.seriesReferenceEdits.end(), [&](const auto& candidate) {
        return candidate.seriesIndex == seriesIndex;
    });
    if (existing == edit.seriesReferenceEdits.end()) {
        ImportedChartEdit::SeriesReferenceEdit seriesEdit;
        seriesEdit.seriesIndex = seriesIndex;
        seriesEdit.categoriesReference = categoriesReference;
        seriesEdit.valuesReference = valuesReference;
        edit.seriesReferenceEdits.push_back(std::move(seriesEdit));
    } else {
        existing->categoriesReference = categoriesReference;
        existing->valuesReference = valuesReference;
    }
    it->series_[seriesIndex].setCategoriesReference(std::move(categoriesReference));
    it->series_[seriesIndex].setValuesReference(std::move(valuesReference));
    it->series_[seriesIndex].setCategoriesCache(ChartSeriesCache{});
    it->series_[seriesIndex].setValuesCache(ChartSeriesCache{});
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesCategoryCache(const std::string& stableId, std::size_t seriesIndex, ChartSeriesCache cache) {
    auto it=std::find_if(charts_.begin(),charts_.end(),[&](const Chart& chart){return chart.imported_&&chart.stableId_==stableId;});
    if(it==charts_.end()||seriesIndex>=it->series_.size()||!cache.present||!validChartSeriesCache(cache)) return false;
    auto& edit=ensureImportedChartEdit(*it); ImportedChartEdit::SeriesCacheEdit e; e.kind=ImportedChartEdit::SeriesCacheEdit::Kind::Categories; e.seriesIndex=seriesIndex; e.cache=cache; edit.seriesCacheEdits.push_back(std::move(e));
    it->series_[seriesIndex].setCategoriesCache(std::move(cache)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartSeriesValueCache(const std::string& stableId, std::size_t seriesIndex, ChartSeriesCache cache) {
    auto it=std::find_if(charts_.begin(),charts_.end(),[&](const Chart& chart){return chart.imported_&&chart.stableId_==stableId;});
    if(it==charts_.end()||seriesIndex>=it->series_.size()||!cache.present||!cache.numeric||!validChartSeriesCache(cache)) return false;
    auto& edit=ensureImportedChartEdit(*it); ImportedChartEdit::SeriesCacheEdit e; e.kind=ImportedChartEdit::SeriesCacheEdit::Kind::Values; e.seriesIndex=seriesIndex; e.cache=cache; edit.seriesCacheEdits.push_back(std::move(e));
    it->series_[seriesIndex].setValuesCache(std::move(cache)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartSeriesTitleCache(const std::string& stableId, std::size_t seriesIndex, ChartSeriesCache cache) {
    auto it=std::find_if(charts_.begin(),charts_.end(),[&](const Chart& chart){return chart.imported_&&chart.stableId_==stableId;});
    if(it==charts_.end()||seriesIndex>=it->series_.size()||!cache.present||cache.numeric||!validChartSeriesCache(cache)) return false;
    if(it->series_[seriesIndex].titleReference().empty()) return false;
    auto& edit=ensureImportedChartEdit(*it); ImportedChartEdit::SeriesCacheEdit e; e.kind=ImportedChartEdit::SeriesCacheEdit::Kind::Title; e.seriesIndex=seriesIndex; e.cache=cache; edit.seriesCacheEdits.push_back(std::move(e));
    it->series_[seriesIndex].setTitleCache(std::move(cache)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::clearChartSeriesCaches(const std::string& stableId, std::size_t seriesIndex) {
    auto it=std::find_if(charts_.begin(),charts_.end(),[&](const Chart& chart){return chart.imported_&&chart.stableId_==stableId;});
    if(it==charts_.end()||seriesIndex>=it->series_.size()) return false;
    auto& edit=ensureImportedChartEdit(*it); ImportedChartEdit::SeriesCacheEdit e; e.kind=ImportedChartEdit::SeriesCacheEdit::Kind::ClearAll; e.seriesIndex=seriesIndex; edit.seriesCacheEdits.push_back(std::move(e));
    it->series_[seriesIndex].setTitleCache(ChartSeriesCache{}); it->series_[seriesIndex].setCategoriesCache(ChartSeriesCache{}); it->series_[seriesIndex].setValuesCache(ChartSeriesCache{});
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotDataLabels(const std::string& stableId, std::size_t plotIndex, Chart::DataLabels labels) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || plotIndex >= it->plots_.size()) return false;
    if (!validDataLabelPosition(labels.position)) return false;
    labels.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.plotDataLabelsEdits.begin(), edit.plotDataLabelsEdits.end(), [&](const auto& candidate) {
        return candidate.plotIndex == plotIndex;
    });
    if (existing == edit.plotDataLabelsEdits.end()) {
        ImportedChartEdit::PlotDataLabelsEdit labelsEdit;
        labelsEdit.plotIndex = plotIndex;
        labelsEdit.labels = labels;
        edit.plotDataLabelsEdits.push_back(std::move(labelsEdit));
    } else {
        existing->labels = labels;
    }
    it->plots_[plotIndex].dataLabels = std::move(labels);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesDataLabels(const std::string& stableId, std::size_t seriesIndex, Chart::DataLabels labels) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    if (!validDataLabelPosition(labels.position)) return false;
    labels.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.seriesDataLabelsEdits.begin(), edit.seriesDataLabelsEdits.end(), [&](const auto& candidate) {
        return candidate.seriesIndex == seriesIndex;
    });
    if (existing == edit.seriesDataLabelsEdits.end()) {
        ImportedChartEdit::SeriesDataLabelsEdit labelsEdit;
        labelsEdit.seriesIndex = seriesIndex;
        labelsEdit.labels = labels;
        edit.seriesDataLabelsEdits.push_back(std::move(labelsEdit));
    } else {
        existing->labels = labels;
    }
    it->series_[seriesIndex].dataLabels_ = std::move(labels);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartPlotDataLabelPoint(const std::string& stableId, std::size_t plotIndex, ChartDataLabelPoint label) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || plotIndex >= it->plots_.size() || !validDataLabelPosition(label.position)) return false;
    auto& labels = it->plots_[plotIndex].dataLabels;
    labels.present = true;
    const auto existingPoint = std::find_if(labels.points.begin(), labels.points.end(), [&](const auto& point) { return point.index == label.index; });
    if (existingPoint == labels.points.end()) labels.points.push_back(label); else *existingPoint = label;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelEdit pointEdit;
    pointEdit.plotLevel = true;
    pointEdit.ownerIndex = plotIndex;
    pointEdit.label = std::move(label);
    edit.pointDataLabelEdits.push_back(std::move(pointEdit));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartPlotDataLabelPoint(const std::string& stableId, std::size_t plotIndex, std::size_t pointIndex) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || plotIndex >= it->plots_.size()) return false;
    auto& points = it->plots_[plotIndex].dataLabels.points;
    const auto existing = std::find_if(points.begin(), points.end(), [&](const auto& point) { return point.index == pointIndex; });
    if (existing == points.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelEdit pointEdit;
    pointEdit.plotLevel = true;
    pointEdit.remove = true;
    pointEdit.ownerIndex = plotIndex;
    pointEdit.label.index = pointIndex;
    edit.pointDataLabelEdits.push_back(std::move(pointEdit));
    points.erase(existing);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesDataLabelPoint(const std::string& stableId, std::size_t seriesIndex, ChartDataLabelPoint label) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validDataLabelPosition(label.position)) return false;
    auto& labels = it->series_[seriesIndex].dataLabels_;
    labels.present = true;
    const auto existingPoint = std::find_if(labels.points.begin(), labels.points.end(), [&](const auto& point) { return point.index == label.index; });
    if (existingPoint == labels.points.end()) labels.points.push_back(label); else *existingPoint = label;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelEdit pointEdit;
    pointEdit.ownerIndex = seriesIndex;
    pointEdit.label = std::move(label);
    edit.pointDataLabelEdits.push_back(std::move(pointEdit));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartSeriesDataLabelPoint(const std::string& stableId, std::size_t seriesIndex, std::size_t pointIndex) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    auto& points = it->series_[seriesIndex].dataLabels_.points;
    const auto existing = std::find_if(points.begin(), points.end(), [&](const auto& point) { return point.index == pointIndex; });
    if (existing == points.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelEdit pointEdit;
    pointEdit.remove = true;
    pointEdit.ownerIndex = seriesIndex;
    pointEdit.label.index = pointIndex;
    edit.pointDataLabelEdits.push_back(std::move(pointEdit));
    points.erase(existing);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesDataLabelPointRichText(const std::string& stableId, std::size_t seriesIndex,
                                                          std::size_t pointIndex, ChartRichText richText) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !richText.present || richText.runs.empty()) return false;
    auto& labels = it->series_[seriesIndex].dataLabels_;
    auto point = std::find_if(labels.points.begin(), labels.points.end(), [&](const auto& candidate) { return candidate.index == pointIndex; });
    if (point == labels.points.end()) {
        ChartDataLabelPoint created;
        created.index = pointIndex;
        created.richText = richText;
        labels.present = true;
        labels.points.push_back(created);
    } else point->richText = richText;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelRichTextEdit richEdit;
    richEdit.seriesIndex = seriesIndex;
    richEdit.pointIndex = pointIndex;
    richEdit.richText = std::move(richText);
    edit.pointDataLabelRichTextEdits.push_back(std::move(richEdit));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesDataPointFormat(const std::string& stableId, std::size_t seriesIndex, ChartDataPointFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        !validChartLineFormat(format.line) || !validChartFillFormat(format.fill) || !validChartMarkerFormat(format.marker))
        return false;
    auto& points = it->series_[seriesIndex].dataPoints_;
    const auto existing = std::find_if(points.begin(), points.end(), [&](const auto& candidate) { return candidate.index == format.index; });
    if (existing == points.end()) points.push_back(format); else *existing = format;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::DataPointFormatEdit pointEdit;
    pointEdit.seriesIndex = seriesIndex;
    pointEdit.format = std::move(format);
    edit.dataPointFormatEdits.push_back(std::move(pointEdit));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartSeriesDataPointFormat(const std::string& stableId, std::size_t seriesIndex, std::size_t pointIndex) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    auto& points = it->series_[seriesIndex].dataPoints_;
    const auto existing = std::find_if(points.begin(), points.end(), [&](const auto& candidate) { return candidate.index == pointIndex; });
    if (existing == points.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::DataPointFormatEdit pointEdit;
    pointEdit.remove = true;
    pointEdit.seriesIndex = seriesIndex;
    pointEdit.format.index = pointIndex;
    edit.dataPointFormatEdits.push_back(std::move(pointEdit));
    points.erase(existing);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesLineFormat(const std::string& stableId, std::size_t seriesIndex, ChartLineFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validChartLineFormat(format)) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::SeriesFormatEdit formatEdit;
    formatEdit.kind = ImportedChartEdit::SeriesFormatEdit::Kind::Line;
    formatEdit.seriesIndex = seriesIndex;
    formatEdit.line = format;
    edit.seriesFormatEdits.push_back(std::move(formatEdit));
    it->series_[seriesIndex].lineFormat_ = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesFillFormat(const std::string& stableId, std::size_t seriesIndex, ChartFillFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validChartFillFormat(format)) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::SeriesFormatEdit formatEdit;
    formatEdit.kind = ImportedChartEdit::SeriesFormatEdit::Kind::Fill;
    formatEdit.seriesIndex = seriesIndex;
    formatEdit.fill = format;
    edit.seriesFormatEdits.push_back(std::move(formatEdit));
    it->series_[seriesIndex].fillFormat_ = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesMarkerFormat(const std::string& stableId, std::size_t seriesIndex, ChartMarkerFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validChartMarkerFormat(format)) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::SeriesFormatEdit formatEdit;
    formatEdit.kind = ImportedChartEdit::SeriesFormatEdit::Kind::Marker;
    formatEdit.seriesIndex = seriesIndex;
    formatEdit.marker = format;
    edit.seriesFormatEdits.push_back(std::move(formatEdit));
    it->series_[seriesIndex].markerFormat_ = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesTrendlineLineFormat(const std::string& stableId, std::size_t seriesIndex,
                                                   std::size_t trendlineIndex, ChartLineFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        trendlineIndex >= it->series_[seriesIndex].trendlines_.size() || !validChartLineFormat(format)) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::TrendlineFormatEdit formatEdit;
    formatEdit.seriesIndex = seriesIndex; formatEdit.trendlineIndex = trendlineIndex; formatEdit.line = format;
    edit.trendlineFormatEdits.push_back(std::move(formatEdit));
    it->series_[seriesIndex].trendlines_[trendlineIndex].lineFormat = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesErrorBarsLineFormat(const std::string& stableId, std::size_t seriesIndex,
                                                   ChartSeries::ErrorBarDirection direction, ChartLineFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validChartLineFormat(format)) return false;
    auto& bars = it->series_[seriesIndex].errorBars_;
    const auto existing = std::find_if(bars.begin(), bars.end(), [&](const auto& candidate) { return candidate.direction == direction; });
    if (existing == bars.end()) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::ErrorBarsFormatEdit formatEdit;
    formatEdit.seriesIndex = seriesIndex; formatEdit.direction = direction; formatEdit.line = format;
    edit.errorBarsFormatEdits.push_back(std::move(formatEdit));
    existing->lineFormat = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesTrendline(const std::string& stableId, std::size_t seriesIndex, std::size_t trendlineIndex,
                                         ChartSeries::Trendline trendline) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        trendlineIndex >= it->series_[seriesIndex].trendlines_.size()) return false;
    if (trendline.type == ChartSeries::TrendlineType::Polynomial && (trendline.order < 2 || trendline.order > 6)) return false;
    if (trendline.type == ChartSeries::TrendlineType::MovingAverage && (trendline.period < 2 || trendline.period > 255)) return false;
    if (!std::isfinite(trendline.forward) || !std::isfinite(trendline.backward) || trendline.forward < 0.0 || trendline.backward < 0.0) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::TrendlineEdit trendlineEdit;
    trendlineEdit.action = ImportedChartEdit::TrendlineEdit::Action::Set;
    trendlineEdit.seriesIndex = seriesIndex;
    trendlineEdit.trendlineIndex = trendlineIndex;
    trendlineEdit.trendline = trendline;
    edit.trendlineEdits.push_back(std::move(trendlineEdit));
    it->series_[seriesIndex].trendlines_[trendlineIndex] = trendline;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::addChartSeriesTrendline(const std::string& stableId, std::size_t seriesIndex, ChartSeries::Trendline trendline) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    if (trendline.type == ChartSeries::TrendlineType::Polynomial && (trendline.order < 2 || trendline.order > 6)) return false;
    if (trendline.type == ChartSeries::TrendlineType::MovingAverage && (trendline.period < 2 || trendline.period > 255)) return false;
    if (!std::isfinite(trendline.forward) || !std::isfinite(trendline.backward) || trendline.forward < 0.0 || trendline.backward < 0.0) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::TrendlineEdit trendlineEdit;
    trendlineEdit.action = ImportedChartEdit::TrendlineEdit::Action::Add;
    trendlineEdit.seriesIndex = seriesIndex;
    trendlineEdit.trendlineIndex = it->series_[seriesIndex].trendlines_.size();
    trendlineEdit.trendline = trendline;
    edit.trendlineEdits.push_back(std::move(trendlineEdit));
    it->series_[seriesIndex].trendlines_.push_back(trendline);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartSeriesTrendline(const std::string& stableId, std::size_t seriesIndex, std::size_t trendlineIndex) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        trendlineIndex >= it->series_[seriesIndex].trendlines_.size()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::TrendlineEdit trendlineEdit;
    trendlineEdit.action = ImportedChartEdit::TrendlineEdit::Action::Remove;
    trendlineEdit.seriesIndex = seriesIndex;
    trendlineEdit.trendlineIndex = trendlineIndex;
    edit.trendlineEdits.push_back(std::move(trendlineEdit));
    it->series_[seriesIndex].trendlines_.erase(it->series_[seriesIndex].trendlines_.begin() + static_cast<std::ptrdiff_t>(trendlineIndex));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesErrorBars(const std::string& stableId, std::size_t seriesIndex, ChartSeries::ErrorBars errorBars) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    if (errorBars.valueType == ChartSeries::ErrorValueType::Custom) {
        if (errorBars.plusReference.empty() || errorBars.minusReference.empty()) return false;
    } else if (!std::isfinite(errorBars.value) || errorBars.value < 0.0) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::ErrorBarsEdit barsEdit;
    barsEdit.seriesIndex = seriesIndex;
    barsEdit.errorBars = errorBars;
    barsEdit.direction = errorBars.direction;
    edit.errorBarsEdits.push_back(std::move(barsEdit));
    auto& bars = it->series_[seriesIndex].errorBars_;
    const auto existing = std::find_if(bars.begin(), bars.end(), [&](const auto& candidate) { return candidate.direction == errorBars.direction; });
    if (existing == bars.end()) bars.push_back(errorBars); else *existing = errorBars;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartSeriesErrorBars(const std::string& stableId, std::size_t seriesIndex, ChartSeries::ErrorBarDirection direction) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    auto& bars = it->series_[seriesIndex].errorBars_;
    const auto existing = std::find_if(bars.begin(), bars.end(), [&](const auto& candidate) { return candidate.direction == direction; });
    if (existing == bars.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::ErrorBarsEdit barsEdit;
    barsEdit.remove = true;
    barsEdit.seriesIndex = seriesIndex;
    barsEdit.direction = direction;
    edit.errorBarsEdits.push_back(std::move(barsEdit));
    bars.erase(existing);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChart(const std::string& stableId) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    edit.removed = true;
    const auto index = static_cast<std::size_t>(std::distance(charts_.begin(), it));
    charts_.erase(it);
    if (index < loadedChartCount_ && loadedChartCount_ > 0) --loadedChartCount_;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

} // namespace xlpp
