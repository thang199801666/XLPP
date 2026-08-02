#pragma once
#include <XLPP/Cell/Cell.h>
#include "Range.h"
#include "Dimensions.h"
#include "Filters.h"
#include "ConditionalFormatting.h"
#include "DataValidation.h"
#include "Tables.h"
#include "PageSetup.h"
#include "Protection.h"
#include "Drawings.h"
#include "SheetView.h"
#include <XLPP/Chart/Chart.h>
#include <XLPP/Pivot/PivotTable.h>
#include <map>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>
#include <filesystem>

namespace xlpp {

struct WorksheetExtents {
    std::size_t minRow;
    std::size_t minColumn;
    std::size_t maxRow;
    std::size_t maxColumn;
};

class Row {
public:
    Row(class Worksheet& sheet, std::size_t rowNumber);
    Cell& cell(std::size_t column);
    const Cell* tryCell(std::size_t column) const noexcept;
    std::size_t number() const noexcept { return rowNumber_; }
    std::vector<Cell*> cells();
    std::vector<CellValue> values() const;

private:
    class Worksheet* sheet_;
    std::size_t rowNumber_;
};

class Worksheet {
public:
    Worksheet() = default;
    explicit Worksheet(std::string name) : name_(std::move(name)) {}

    const std::string& name() const noexcept { return name_; }
    void rename(std::string name);

    Cell& cell(const std::string& address);
    Cell& cell(std::size_t row, std::size_t column);
    const Cell* tryCell(const std::string& address) const noexcept;
    const Cell* tryCell(std::size_t row, std::size_t column) const noexcept;

    CellRange range(const std::string& address);
    CellRange range(std::size_t minRow, std::size_t minColumn,
                    std::size_t maxRow, std::size_t maxColumn);

    void append(const std::vector<CellValue>& values);

    void mergeCells(const std::string& rangeAddress);
    void unmergeCells(const std::string& rangeAddress);
    bool isMerged(const std::string& cellAddress) const;
    const std::vector<std::string>& mergedRanges() const noexcept { return mergedRanges_; }

    void freezePanes(const std::string& topLeftCell);
    void clearFreezePanes() noexcept { freezePane_.reset(); }
    const std::optional<std::string>& frozenPane() const noexcept { return freezePane_; }

    RowDimension& rowDimension(std::size_t row);
    const RowDimension* tryRowDimension(std::size_t row) const noexcept;
    ColumnDimension& columnDimension(std::size_t column);
    ColumnDimension& columnDimension(const std::string& columnName);
    const ColumnDimension* tryColumnDimension(std::size_t column) const noexcept;
    const std::map<std::size_t, RowDimension>& rowDimensions() const noexcept { return rowDimensions_; }
    const std::map<std::size_t, ColumnDimension>& columnDimensions() const noexcept { return columnDimensions_; }

    AutoFilter& autoFilter() noexcept { return autoFilter_; }
    const AutoFilter& autoFilter() const noexcept { return autoFilter_; }

    ConditionalFormattingCollection& conditionalFormatting() noexcept { return conditionalFormatting_; }
    const ConditionalFormattingCollection& conditionalFormatting() const noexcept { return conditionalFormatting_; }
    DataValidationCollection& dataValidations() noexcept { return dataValidations_; }
    const DataValidationCollection& dataValidations() const noexcept { return dataValidations_; }

    Table& addTable(std::string name, std::string reference);
    Table* table(const std::string& name) noexcept;
    const Table* table(const std::string& name) const noexcept;
    std::vector<Table>& tables() noexcept { return tables_; }
    const std::vector<Table>& tables() const noexcept { return tables_; }

    PageSetup& pageSetup() noexcept { return pageSetup_; }
    const PageSetup& pageSetup() const noexcept { return pageSetup_; }
    PageMargins& pageMargins() noexcept { return pageMargins_; }
    const PageMargins& pageMargins() const noexcept { return pageMargins_; }
    PrintOptions& printOptions() noexcept { return printOptions_; }
    const PrintOptions& printOptions() const noexcept { return printOptions_; }
    HeaderFooter& headerFooter() noexcept { return headerFooter_; }
    const HeaderFooter& headerFooter() const noexcept { return headerFooter_; }
    WorksheetProtection& protection() noexcept { return protection_; }
    const WorksheetProtection& protection() const noexcept { return protection_; }
    Image& addImage(Image image) { images_.push_back(std::move(image)); return images_.back(); }
    Image& addImage(const std::filesystem::path& path, const std::string& anchor) { return addImage(Image::fromFile(path, anchor)); }
    std::vector<Image>& images() noexcept { return images_; }
    const std::vector<Image>& images() const noexcept { return images_; }

    SheetView& sheetView() noexcept { return sheetView_; }
    const SheetView& sheetView() const noexcept { return sheetView_; }
    void setSheetView(SheetView v) noexcept { sheetView_ = std::move(v); }

    void addChart(Chart chart) { charts_.push_back(std::move(chart)); }
    Chart& chart(std::size_t index) { return charts_.at(index); }
    const Chart& chart(std::size_t index) const { return charts_.at(index); }
    std::size_t chartCount() const noexcept { return charts_.size(); }
    std::vector<Chart>& charts() noexcept { return charts_; }
    const std::vector<Chart>& charts() const noexcept { return charts_; }

    void addPivotTable(PivotTable pt) { pivotTables_.push_back(std::move(pt)); }
    std::vector<PivotTable>& pivotTables() noexcept { return pivotTables_; }
    const std::vector<PivotTable>& pivotTables() const noexcept { return pivotTables_; }

    const std::string& printArea() const noexcept { return printArea_; }
    void setPrintArea(std::string v) { printArea_ = std::move(v); }
    const std::string& printTitlesRows() const noexcept { return printTitlesRows_; }
    void setPrintTitlesRows(std::string v) { printTitlesRows_ = std::move(v); }
    const std::string& printTitlesCols() const noexcept { return printTitlesCols_; }
    void setPrintTitlesCols(std::string v) { printTitlesCols_ = std::move(v); }

    void insertRows(std::size_t index, std::size_t amount = 1);
    void deleteRows(std::size_t index, std::size_t amount = 1);
    void insertColumns(std::size_t index, std::size_t amount = 1);
    void deleteColumns(std::size_t index, std::size_t amount = 1);

    std::size_t maxRow() const noexcept;
    std::size_t maxColumn() const noexcept;
    WorksheetExtents extents() const noexcept;
    std::string dimensions() const;
    bool empty() const noexcept { return cells_.empty(); }
    std::size_t rowCount() const noexcept { return extents().maxRow; }
    std::size_t columnCount() const noexcept { return extents().maxColumn; }
    Row row(std::size_t rowNumber) { return Row(*this, rowNumber); }
    std::vector<Row> rows();
    std::vector<std::vector<CellValue>> iterRows(std::size_t minRow = 0, std::size_t maxRow = 0,
                                                  std::size_t minCol = 0, std::size_t maxCol = 0) const;
    std::vector<std::vector<CellValue>> iterCols(std::size_t minRow = 0, std::size_t maxRow = 0,
                                                  std::size_t minCol = 0, std::size_t maxCol = 0) const;

    const std::map<std::uint64_t, Cell>& cells() const noexcept { return cells_; }

    // Dirty tracking for differential save. A worksheet becomes dirty when
    // any of its content is mutated; `save()` only re-serializes dirty sheets.
    bool dirty() const noexcept { return dirty_; }
    void markDirty() noexcept { dirty_ = true; }
    void clearDirty() const noexcept { dirty_ = false; }

private:
    void shiftRows(std::size_t index, std::size_t amount, bool insert);
    void shiftColumns(std::size_t index, std::size_t amount, bool insert);

    std::string name_;
    std::map<std::uint64_t, Cell> cells_;
    std::vector<std::string> mergedRanges_;
    struct MergedRangeCache { std::size_t minRow, minColumn, maxRow, maxColumn; };
    std::vector<MergedRangeCache> mergedRangesParsed_;
    std::optional<std::string> freezePane_;
    std::map<std::size_t, RowDimension> rowDimensions_;
    std::map<std::size_t, ColumnDimension> columnDimensions_;
    AutoFilter autoFilter_;
    ConditionalFormattingCollection conditionalFormatting_;
    DataValidationCollection dataValidations_;
    std::vector<Table> tables_;
    PageSetup pageSetup_;
    PageMargins pageMargins_;
    PrintOptions printOptions_;
    HeaderFooter headerFooter_;
    WorksheetProtection protection_;
    std::vector<Image> images_;
    SheetView sheetView_;
    std::string printArea_, printTitlesRows_, printTitlesCols_;
    std::vector<Chart> charts_;
    std::vector<PivotTable> pivotTables_;
    mutable bool dirty_{true};
};
}
