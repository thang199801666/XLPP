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
#include "Sparkline.h"
#include <XLPP/Chart/Chart.h>
#include <XLPP/Pivot/PivotTable.h>
#include <XLPP/Workbook/StructuralEdit.h>
#include <map>
#include <algorithm>
#include <cstdint>
#include <optional>
#include <set>
#include <string>
#include <vector>
#include <filesystem>

namespace xlpp {
namespace internal { struct WorkbookDrawingAccess; struct WorkbookSheetOperationsAccess; }

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
    friend class Workbook;
    friend struct internal::WorkbookSheetOperationsAccess;
public:
    Worksheet() = default;
    explicit Worksheet(std::string name) : name_(std::move(name)) {}

    const std::string& name() const noexcept { return name_; }
    void rename(std::string name);
    const std::string& vbaCodeName() const noexcept { return vbaCodeName_; }
    void setVbaCodeName(std::string value) { vbaCodeName_ = std::move(value); dirty_ = true; }

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
    void clearFreezePanes() noexcept { freezePane_.reset(); dirty_ = true; }
    const std::optional<std::string>& frozenPane() const noexcept { return freezePane_; }

    RowDimension& rowDimension(std::size_t row);
    const RowDimension* tryRowDimension(std::size_t row) const noexcept;
    ColumnDimension& columnDimension(std::size_t column);
    ColumnDimension& columnDimension(const std::string& columnName);
    const ColumnDimension* tryColumnDimension(std::size_t column) const noexcept;
    const std::map<std::size_t, RowDimension>& rowDimensions() const noexcept { return rowDimensions_; }
    const std::map<std::size_t, ColumnDimension>& columnDimensions() const noexcept { return columnDimensions_; }

    // Raw source sheetFormatPr is retained for loaded worksheets because its
    // default column/row metrics participate in two-cell DrawingML geometry.
    const std::string& preservedSheetFormatPrXml() const noexcept { return preservedSheetFormatPrXml_; }
    void setLoadedSheetFormatPrXml(std::string value) { preservedSheetFormatPrXml_ = std::move(value); }

    AutoFilter& autoFilter() noexcept { dirty_ = true; return autoFilter_; }
    const AutoFilter& autoFilter() const noexcept { return autoFilter_; }

    ConditionalFormattingCollection& conditionalFormatting() noexcept { dirty_ = true; return conditionalFormatting_; }
    const ConditionalFormattingCollection& conditionalFormatting() const noexcept { return conditionalFormatting_; }
    DataValidationCollection& dataValidations() noexcept { dirty_ = true; return dataValidations_; }
    const DataValidationCollection& dataValidations() const noexcept { return dataValidations_; }

    Table& addTable(std::string name, std::string reference);
    Table* table(const std::string& name) noexcept;
    const Table* table(const std::string& name) const noexcept;
    std::vector<Table>& tables() noexcept { dirty_ = true; return tables_; }
    const std::vector<Table>& tables() const noexcept { return tables_; }

    PageSetup& pageSetup() noexcept { dirty_ = true; return pageSetup_; }
    const PageSetup& pageSetup() const noexcept { return pageSetup_; }
    PageMargins& pageMargins() noexcept { dirty_ = true; return pageMargins_; }
    const PageMargins& pageMargins() const noexcept { return pageMargins_; }
    PrintOptions& printOptions() noexcept { dirty_ = true; return printOptions_; }
    const PrintOptions& printOptions() const noexcept { return printOptions_; }
    HeaderFooter& headerFooter() noexcept { dirty_ = true; return headerFooter_; }
    const HeaderFooter& headerFooter() const noexcept { return headerFooter_; }
    WorksheetProtection& protection() noexcept { dirty_ = true; return protection_; }
    const WorksheetProtection& protection() const noexcept { return protection_; }
    Image& addImage(Image image) { dirty_ = true; drawingAppendDirty_ = true; images_.push_back(std::move(image)); return images_.back(); }
    Image& addImage(const std::filesystem::path& path, const std::string& anchor) { return addImage(Image::fromFile(path, anchor)); }
    // Loader-only path. Existing images are exposed through the normal const
    // images() API but do not mark a preserved drawing for regeneration.
    Image& addLoadedImage(Image image) { images_.push_back(std::move(image)); loadedImageCount_ = images_.size(); return images_.back(); }
    std::vector<Image>& images() noexcept { dirty_ = true; drawingsDirty_ = true; drawingAppendDirty_ = false; return images_; }
    const std::vector<Image>& images() const noexcept { return images_; }
    std::size_t loadedImageCount() const noexcept { return loadedImageCount_; }
    std::size_t appendedImageCount() const noexcept { return images_.size() > loadedImageCount_ ? images_.size() - loadedImageCount_ : 0; }

    // Selective edits for package-imported images. These APIs intentionally do
    // not mark the whole drawing dirty: Workbook::save patches only the target
    // anchor/relationship/media while preserving sibling DrawingML verbatim.
    const Image* imageByStableId(const std::string& stableId) const noexcept;
    bool moveImage(const std::string& stableId, const std::string& anchor);
    bool moveImageAbsolute(const std::string& stableId, long long xEmu, long long yEmu);
    bool resizeImage(const std::string& stableId, double widthPixels, double heightPixels);
    bool replaceImage(const std::string& stableId, Image replacement);
    bool replaceImage(const std::string& stableId, const std::filesystem::path& path);
    bool removeImage(const std::string& stableId);

    // Stable-ID selective chart edits preserve unsupported chart XML and
    // sibling drawing objects verbatim. They are available only for charts
    // imported from an existing package.
    const Chart* chartByStableId(const std::string& stableId) const noexcept;
    bool moveChart(const std::string& stableId, const std::string& anchor);
    bool moveChartAbsolute(const std::string& stableId, long long xEmu, long long yEmu);
    bool resizeChart(const std::string& stableId, double widthPixels, double heightPixels);
    bool setChartTitle(const std::string& stableId, std::string title);
    bool setChartStyle(const std::string& stableId, std::string style);
    bool setChartTitleRichText(const std::string& stableId, ChartRichText richText);
    bool setChartXAxisTitle(const std::string& stableId, std::string title);
    bool setChartYAxisTitle(const std::string& stableId, std::string title);
    bool setChartAxisTitle(const std::string& stableId, std::uint64_t axisId, std::string title);
    bool setChartAxisTitleRichText(const std::string& stableId, std::uint64_t axisId, ChartRichText richText);
    bool setChartAxisNumberFormat(const std::string& stableId, std::uint64_t axisId, std::string formatCode, bool sourceLinked = false);
    bool setChartAxisTicks(const std::string& stableId, std::uint64_t axisId, std::string majorTickMark,
                           std::string minorTickMark, std::string tickLabelPosition);
    bool setChartAxisUnits(const std::string& stableId, std::uint64_t axisId, double majorUnit, double minorUnit = 0.0);
    bool setChartAxisScaling(const std::string& stableId, std::uint64_t axisId, ChartAxisScaling scaling);
    bool setChartAxisCrossing(const std::string& stableId, std::uint64_t axisId, std::string crosses, std::string crossBetween = {});
    bool setChartAxisCrossesAt(const std::string& stableId, std::uint64_t axisId, double crossesAt);
    bool clearChartAxisCrossesAt(const std::string& stableId, std::uint64_t axisId);
    bool setChartAxisDisplayUnits(const std::string& stableId, std::uint64_t axisId, ChartDisplayUnits units);
    bool clearChartAxisDisplayUnits(const std::string& stableId, std::uint64_t axisId);
    bool setChartAxisLineFormat(const std::string& stableId, std::uint64_t axisId, ChartLineFormat format);
    bool setChartAxisGridlineFormat(const std::string& stableId, std::uint64_t axisId, bool major, ChartLineFormat format);
    bool removeChartAxisGridlines(const std::string& stableId, std::uint64_t axisId, bool major);
    bool setChartAreaLineFormat(const std::string& stableId, ChartLineFormat format);
    bool setChartAreaFillFormat(const std::string& stableId, ChartFillFormat format);
    bool setChartPlotAreaLineFormat(const std::string& stableId, ChartLineFormat format);
    bool setChartPlotAreaFillFormat(const std::string& stableId, ChartFillFormat format);
    bool setChartPlotAreaLayout(const std::string& stableId, ChartManualLayout layout);
    bool setChartView3D(const std::string& stableId, ChartView3D view);
    bool setChartFloorFormat(const std::string& stableId, ChartWallFormat format);
    bool setChartSideWallFormat(const std::string& stableId, ChartWallFormat format);
    bool setChartBackWallFormat(const std::string& stableId, ChartWallFormat format);
    bool setChartDataTable(const std::string& stableId, ChartDataTable table);
    bool removeChartDataTable(const std::string& stableId);
    bool setChartPlotDropLines(const std::string& stableId, std::size_t plotIndex, ChartLineFormat format);
    bool removeChartPlotDropLines(const std::string& stableId, std::size_t plotIndex);
    bool setChartPlotHighLowLines(const std::string& stableId, std::size_t plotIndex, ChartLineFormat format);
    bool removeChartPlotHighLowLines(const std::string& stableId, std::size_t plotIndex);
    bool setChartPlotUpDownBars(const std::string& stableId, std::size_t plotIndex, ChartUpDownBars bars);
    bool removeChartPlotUpDownBars(const std::string& stableId, std::size_t plotIndex);
    bool setChartPlotFirstSliceAngle(const std::string& stableId, std::size_t plotIndex, int degrees);
    bool setChartPlotDoughnutHoleSize(const std::string& stableId, std::size_t plotIndex, int percent);
    bool setChartPlotRadarStyle(const std::string& stableId, std::size_t plotIndex, std::string style);
    bool setChartPlotProjectedPieOptions(const std::string& stableId, std::size_t plotIndex, ChartProjectedPieOptions options);
    bool setChartPlotLeaderLineFormat(const std::string& stableId, std::size_t plotIndex, ChartLineFormat format);
    bool removeChartPlotLeaderLines(const std::string& stableId, std::size_t plotIndex);
    bool setChartSeriesLeaderLineFormat(const std::string& stableId, std::size_t seriesIndex, ChartLineFormat format);
    bool removeChartSeriesLeaderLines(const std::string& stableId, std::size_t seriesIndex);
    bool setChartLegend(const std::string& stableId, bool show, std::string position = "r");
    bool setChartLegendLayout(const std::string& stableId, ChartManualLayout layout);
    bool setChartLegendOverlay(const std::string& stableId, bool overlay);
    bool setChartLegendLineFormat(const std::string& stableId, ChartLineFormat format);
    bool setChartLegendFillFormat(const std::string& stableId, ChartFillFormat format);
    bool setChartSeriesTitle(const std::string& stableId, std::size_t seriesIndex, std::string title);
    bool setChartSeriesReferences(const std::string& stableId, std::size_t seriesIndex,
                                  std::string categoriesReference, std::string valuesReference);
    bool setChartSeriesCategoryCache(const std::string& stableId, std::size_t seriesIndex, ChartSeriesCache cache);
    bool setChartSeriesValueCache(const std::string& stableId, std::size_t seriesIndex, ChartSeriesCache cache);
    bool setChartSeriesTitleCache(const std::string& stableId, std::size_t seriesIndex, ChartSeriesCache cache);
    bool clearChartSeriesCaches(const std::string& stableId, std::size_t seriesIndex);
    bool setChartPlotDataLabels(const std::string& stableId, std::size_t plotIndex, Chart::DataLabels labels);
    bool setChartSeriesDataLabels(const std::string& stableId, std::size_t seriesIndex, Chart::DataLabels labels);
    bool setChartPlotDataLabelPoint(const std::string& stableId, std::size_t plotIndex, ChartDataLabelPoint label);
    bool removeChartPlotDataLabelPoint(const std::string& stableId, std::size_t plotIndex, std::size_t pointIndex);
    bool setChartSeriesDataLabelPoint(const std::string& stableId, std::size_t seriesIndex, ChartDataLabelPoint label);
    bool removeChartSeriesDataLabelPoint(const std::string& stableId, std::size_t seriesIndex, std::size_t pointIndex);
    bool setChartSeriesDataLabelPointRichText(const std::string& stableId, std::size_t seriesIndex, std::size_t pointIndex, ChartRichText richText);
    bool setChartSeriesDataPointFormat(const std::string& stableId, std::size_t seriesIndex, ChartDataPointFormat format);
    bool removeChartSeriesDataPointFormat(const std::string& stableId, std::size_t seriesIndex, std::size_t pointIndex);
    bool setChartSeriesLineFormat(const std::string& stableId, std::size_t seriesIndex, ChartLineFormat format);
    bool setChartSeriesFillFormat(const std::string& stableId, std::size_t seriesIndex, ChartFillFormat format);
    bool setChartSeriesMarkerFormat(const std::string& stableId, std::size_t seriesIndex, ChartMarkerFormat format);
    bool setChartSeriesTrendlineLineFormat(const std::string& stableId, std::size_t seriesIndex, std::size_t trendlineIndex, ChartLineFormat format);
    bool setChartSeriesErrorBarsLineFormat(const std::string& stableId, std::size_t seriesIndex, ChartSeries::ErrorBarDirection direction, ChartLineFormat format);
    bool setChartSeriesTrendline(const std::string& stableId, std::size_t seriesIndex, std::size_t trendlineIndex,
                                 ChartSeries::Trendline trendline);
    bool addChartSeriesTrendline(const std::string& stableId, std::size_t seriesIndex, ChartSeries::Trendline trendline);
    bool removeChartSeriesTrendline(const std::string& stableId, std::size_t seriesIndex, std::size_t trendlineIndex);
    bool setChartSeriesErrorBars(const std::string& stableId, std::size_t seriesIndex, ChartSeries::ErrorBars errorBars);
    bool removeChartSeriesErrorBars(const std::string& stableId, std::size_t seriesIndex, ChartSeries::ErrorBarDirection direction);
    bool removeChart(const std::string& stableId);

    SheetView& sheetView() noexcept { dirty_ = true; return sheetView_; }
    const SheetView& sheetView() const noexcept { return sheetView_; }
    void setSheetView(SheetView v) noexcept { sheetView_ = std::move(v); dirty_ = true; }

    Chart& addChart(Chart chart) {
        dirty_ = true;
        // A loaded worksheet with a preserved drawing can append a new chart
        // without regenerating sibling DrawingML. New workbooks still take the
        // normal generated-drawing path because drawingsDirty_ starts true.
        if (!drawingsDirty_) drawingAppendDirty_ = true;
        else drawingAppendDirty_ = false;
        charts_.push_back(std::move(chart));
        return charts_.back();
    }
    // Loader-only path: records a chart discovered in an existing drawing
    // without marking that drawing for destructive regeneration.
    Chart& addLoadedChart(Chart chart) { charts_.push_back(std::move(chart)); loadedChartCount_ = charts_.size(); return charts_.back(); }
    Chart& chart(std::size_t index) { dirty_ = true; drawingsDirty_ = true; drawingAppendDirty_ = false; return charts_.at(index); }
    const Chart& chart(std::size_t index) const { return charts_.at(index); }
    std::size_t chartCount() const noexcept { return charts_.size(); }
    std::vector<Chart>& charts() noexcept { dirty_ = true; drawingsDirty_ = true; drawingAppendDirty_ = false; return charts_; }
    const std::vector<Chart>& charts() const noexcept { return charts_; }

    PivotTable& addPivotTable(PivotTable pt) {
        dirty_ = true;
        if (loadedPivotCount_ > 0 && !pivotsDirty_) pivotAppendDirty_ = true;
        else pivotsDirty_ = true;
        pivotTables_.push_back(std::move(pt));
        return pivotTables_.back();
    }
    // Loader-only path: expose an existing PivotTable through the object model
    // without forcing regeneration of its OOXML parts on an unrelated save.
    PivotTable& addLoadedPivotTable(PivotTable pt) { pivotTables_.push_back(std::move(pt)); loadedPivotCount_ = pivotTables_.size(); return pivotTables_.back(); }
    std::size_t loadedPivotCount() const noexcept { return loadedPivotCount_; }
    std::vector<PivotTable>& pivotTables() noexcept { dirty_ = true; pivotsDirty_ = true; pivotAppendDirty_ = false; return pivotTables_; }
    const std::vector<PivotTable>& pivotTables() const noexcept { return pivotTables_; }

    const std::string& printArea() const noexcept { return printArea_; }
    void setPrintArea(std::string v) { printArea_ = std::move(v); dirty_ = true; }
    const std::string& printTitlesRows() const noexcept { return printTitlesRows_; }
    void setPrintTitlesRows(std::string v) { printTitlesRows_ = std::move(v); dirty_ = true; }
    const std::string& printTitlesCols() const noexcept { return printTitlesCols_; }
    void setPrintTitlesCols(std::string v) { printTitlesCols_ = std::move(v); dirty_ = true; }

    // Manual page breaks (SpreadsheetML <rowBreaks>/<colBreaks>). Row breaks
    // hold the 0-based row index above which a page break occurs; column breaks
    // hold the 0-based column index to the left of which a break occurs.
    void addRowBreak(std::size_t row) {
        rowBreaks_.insert(row);
        syncBreakList(rowBreaks_, rowBreakList_);
        dirty_ = true;
    }
    void removeRowBreak(std::size_t row) noexcept {
        rowBreaks_.erase(row);
        syncBreakList(rowBreaks_, rowBreakList_);
        dirty_ = true;
    }
    void clearRowBreaks() noexcept {
        rowBreaks_.clear();
        rowBreakList_.clear();
        dirty_ = true;
    }
    const std::vector<std::size_t>& rowBreaks() const noexcept { return rowBreakList_; }
    void setRowBreaks(std::vector<std::size_t> rows) {
        rowBreakList_ = std::move(rows);
        rowBreaks_.clear();
        rowBreaks_.insert(rowBreakList_.begin(), rowBreakList_.end());
        dirty_ = true;
    }
    void addColumnBreak(std::size_t column) {
        columnBreaks_.insert(column);
        syncBreakList(columnBreaks_, columnBreakList_);
        dirty_ = true;
    }
    void removeColumnBreak(std::size_t column) noexcept {
        columnBreaks_.erase(column);
        syncBreakList(columnBreaks_, columnBreakList_);
        dirty_ = true;
    }
    void clearColumnBreaks() noexcept {
        columnBreaks_.clear();
        columnBreakList_.clear();
        dirty_ = true;
    }
    const std::vector<std::size_t>& columnBreaks() const noexcept { return columnBreakList_; }
    void setColumnBreaks(std::vector<std::size_t> columns) {
        columnBreakList_ = std::move(columns);
        columnBreaks_.clear();
        columnBreaks_.insert(columnBreakList_.begin(), columnBreakList_.end());
        dirty_ = true;
    }

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
    bool drawingsDirty() const noexcept { return drawingsDirty_; }
    bool drawingAppendDirty() const noexcept { return drawingAppendDirty_; }
    bool pivotsDirty() const noexcept { return pivotsDirty_; }
    bool pivotAppendDirty() const noexcept { return pivotAppendDirty_; }
    std::size_t appendedPivotCount() const noexcept { return pivotTables_.size() > loadedPivotCount_ ? pivotTables_.size() - loadedPivotCount_ : 0; }
    std::size_t loadedChartCount() const noexcept { return loadedChartCount_; }
    std::size_t appendedChartCount() const noexcept { return charts_.size() > loadedChartCount_ ? charts_.size() - loadedChartCount_ : 0; }
    void markDirty() noexcept { dirty_ = true; }
    bool hasTrackedCellChanges() const noexcept {
        if (!trackedCellKeys_.empty()) return true;
        return std::any_of(cells_.begin(), cells_.end(), [](const auto& entry) { return entry.second.mutationRevision() != 0; });
    }
    std::size_t trackedCellChangeCount() const noexcept {
        // Both containers are ordered by the packed cell key.  Merge-walk the
        // two sets rather than cloning trackedCellKeys_ or performing one
        // logarithmic lookup per mutated cell.  This keeps the query O(N+K)
        // with O(1) auxiliary memory even on million-cell worksheets.
        std::size_t count = trackedCellKeys_.size();
        auto tracked = trackedCellKeys_.begin();
        const auto trackedEnd = trackedCellKeys_.end();
        for (const auto& entry : cells_) {
            if (entry.second.mutationRevision() == 0) continue;
            while (tracked != trackedEnd && *tracked < entry.first) ++tracked;
            if (tracked == trackedEnd || *tracked != entry.first) ++count;
        }
        return count;
    }
    const std::set<std::uint64_t>& trackedCellKeys() const noexcept { return trackedCellKeys_; }
    void clearTrackedCellChanges() const noexcept {
        trackedCellKeys_.clear();
        for (const auto& entry : cells_) entry.second.clearMutationRevision();
    }
    void clearDirty() const noexcept { dirty_ = false; drawingsDirty_ = false; drawingAppendDirty_ = false; pivotsDirty_ = false; pivotAppendDirty_ = false; clearTrackedCellChanges(); }

private:
    void shiftRows(std::size_t index, std::size_t amount, bool insert,
                   const StructuralEditOptions* options = nullptr, StructuralEditReport* report = nullptr);
    void shiftColumns(std::size_t index, std::size_t amount, bool insert,
                      const StructuralEditOptions* options = nullptr, StructuralEditReport* report = nullptr);
    void shiftStructure(std::size_t index, std::size_t amount, bool insert, bool rows,
                        const StructuralEditOptions* options, StructuralEditReport* report);

    std::string name_;
    std::string vbaCodeName_;
    std::map<std::uint64_t, Cell> cells_;
    // An empty worksheet has the canonical A1:A1 extent, so the cache is
    // valid from construction. This lets monotonic append()/cell() insertion
    // maintain extents incrementally instead of forcing a first-query O(N)
    // scan after bulk ingestion.
    mutable WorksheetExtents extentsCache_{1, 1, 1, 1};
    mutable bool extentsCacheValid_{true};
    std::vector<std::string> mergedRanges_;
    struct MergedRangeCache { std::size_t minRow, minColumn, maxRow, maxColumn; };
    std::vector<MergedRangeCache> mergedRangesParsed_;
    std::optional<std::string> freezePane_;
    std::map<std::size_t, RowDimension> rowDimensions_;
    std::map<std::size_t, ColumnDimension> columnDimensions_;
    std::string preservedSheetFormatPrXml_;
    AutoFilter autoFilter_;
    ConditionalFormattingCollection conditionalFormatting_;
    DataValidationCollection dataValidations_;
    std::vector<Table> tables_;
    PageSetup pageSetup_;
    PageMargins pageMargins_;
    struct ImportedImageEdit {
        std::string stableId;
        std::string sourceDrawingPart;
        std::string sourceMediaPart;
        std::string sourceRelationshipId;
        DrawingAnchorInfo originalAnchor;
        DrawingAnchorInfo anchor;
        bool moved{false};
        bool resized{false};
        bool removed{false};
        bool replaced{false};
        std::vector<unsigned char> replacementBytes;
        std::string replacementExtension;
    };

    struct ImportedChartEdit {
        struct SeriesReferenceEdit {
            std::size_t seriesIndex{0};
            std::string categoriesReference;
            std::string valuesReference;
        };
        struct SeriesCacheEdit {
            enum class Kind { Categories, Values, Title, ClearAll };
            Kind kind{Kind::Categories};
            std::size_t seriesIndex{0};
            ChartSeriesCache cache;
        };
        std::string stableId;
        std::string sourceDrawingPart;
        std::string sourceChartPart;
        std::string sourceRelationshipId;
        Chart::Type chartType{Chart::Type::Bar};
        DrawingAnchorInfo originalAnchor;
        DrawingAnchorInfo anchor;
        bool moved{false};
        bool resized{false};
        struct SeriesTitleEdit {
            std::size_t seriesIndex{0};
            std::string title;
        };
        bool titleChanged{false};
        std::string title;
        bool styleChanged{false};
        std::string style;
        bool titleRichTextChanged{false};
        ChartRichText titleRichText;
        bool xAxisTitleChanged{false};
        std::string xAxisTitle;
        bool yAxisTitleChanged{false};
        std::string yAxisTitle;
        std::uint64_t primaryXAxisId{0};
        std::uint64_t primaryYAxisId{0};
        struct AxisTitleEdit {
            std::uint64_t axisId{0};
            std::string title;
        };
        struct AxisRichTitleEdit {
            std::uint64_t axisId{0};
            ChartRichText richText;
        };
        struct AxisFormatEdit {
            enum class Kind { NumberFormat, Ticks, Units, Scaling, Crossing, CrossesAt, ClearCrossesAt, DisplayUnits, ClearDisplayUnits, Line, MajorGridline, MinorGridline, RemoveMajorGridline, RemoveMinorGridline };
            Kind kind{Kind::NumberFormat};
            std::uint64_t axisId{0};
            std::string value1;
            std::string value2;
            std::string value3;
            bool flag{false};
            double number1{0.0};
            double number2{0.0};
            ChartLineFormat line;
            ChartAxisScaling scaling;
            ChartDisplayUnits displayUnits;
        };
        struct AreaFormatEdit {
            enum class Owner { ChartArea, PlotArea };
            enum class Kind { Line, Fill };
            Owner owner{Owner::ChartArea};
            Kind kind{Kind::Line};
            ChartLineFormat line;
            ChartFillFormat fill;
        };
        struct LayoutEdit {
            enum class Owner { PlotArea, Legend };
            Owner owner{Owner::PlotArea};
            ChartManualLayout layout;
        };
        struct DataTableEdit {
            bool remove{false};
            ChartDataTable table;
        };
        struct View3DEdit {
            ChartView3D view;
        };
        struct WallFormatEdit {
            enum class Owner { Floor, SideWall, BackWall };
            Owner owner{Owner::Floor};
            ChartWallFormat format;
        };
        struct PlotAuxiliaryEdit {
            enum class Kind { DropLines, HighLowLines, UpDownBars };
            Kind kind{Kind::DropLines};
            bool remove{false};
            std::size_t plotIndex{0};
            ChartLineFormat line;
            ChartUpDownBars upDownBars;
        };
        struct PlotTypeSpecificEdit {
            enum class Kind { FirstSliceAngle, DoughnutHoleSize, RadarStyle, ProjectedPie };
            Kind kind{Kind::FirstSliceAngle};
            std::size_t plotIndex{0};
            int integerValue{0};
            std::string textValue;
            ChartProjectedPieOptions projectedPie;
        };
        struct LeaderLineEdit {
            bool plotLevel{false};
            bool remove{false};
            std::size_t ownerIndex{0};
            ChartLineFormat line;
        };
        struct PlotDataLabelsEdit {
            std::size_t plotIndex{0};
            Chart::DataLabels labels;
        };
        struct SeriesDataLabelsEdit {
            std::size_t seriesIndex{0};
            Chart::DataLabels labels;
        };
        struct PointDataLabelEdit {
            bool plotLevel{false};
            bool remove{false};
            std::size_t ownerIndex{0};
            ChartDataLabelPoint label;
        };
        struct PointDataLabelRichTextEdit {
            std::size_t seriesIndex{0};
            std::size_t pointIndex{0};
            ChartRichText richText;
        };
        struct DataPointFormatEdit {
            bool remove{false};
            std::size_t seriesIndex{0};
            ChartDataPointFormat format;
        };
        struct SeriesFormatEdit {
            enum class Kind { Line, Fill, Marker };
            Kind kind{Kind::Line};
            std::size_t seriesIndex{0};
            ChartLineFormat line;
            ChartFillFormat fill;
            ChartMarkerFormat marker;
        };
        struct TrendlineFormatEdit {
            std::size_t seriesIndex{0};
            std::size_t trendlineIndex{0};
            ChartLineFormat line;
        };
        struct ErrorBarsFormatEdit {
            std::size_t seriesIndex{0};
            ChartSeries::ErrorBarDirection direction{ChartSeries::ErrorBarDirection::Y};
            ChartLineFormat line;
        };
        struct TrendlineEdit {
            enum class Action { Set, Add, Remove };
            Action action{Action::Set};
            std::size_t seriesIndex{0};
            std::size_t trendlineIndex{0};
            ChartSeries::Trendline trendline;
        };
        struct ErrorBarsEdit {
            bool remove{false};
            std::size_t seriesIndex{0};
            ChartSeries::ErrorBars errorBars;
            ChartSeries::ErrorBarDirection direction{ChartSeries::ErrorBarDirection::Y};
        };
        std::vector<AxisRichTitleEdit> axisRichTitleEdits;
        std::vector<AxisFormatEdit> axisFormatEdits;
        std::vector<AreaFormatEdit> areaFormatEdits;
        std::vector<LayoutEdit> layoutEdits;
        std::vector<DataTableEdit> dataTableEdits;
        std::vector<View3DEdit> view3DEdits;
        std::vector<WallFormatEdit> wallFormatEdits;
        std::vector<PlotAuxiliaryEdit> plotAuxiliaryEdits;
        std::vector<PlotTypeSpecificEdit> plotTypeSpecificEdits;
        std::vector<LeaderLineEdit> leaderLineEdits;
        bool legendOverlayChanged{false};
        bool legendOverlay{false};
        bool legendLineFormatChanged{false};
        ChartLineFormat legendLineFormat;
        bool legendFillFormatChanged{false};
        ChartFillFormat legendFillFormat;
        std::vector<AxisTitleEdit> axisTitleEdits;
        std::vector<PlotDataLabelsEdit> plotDataLabelsEdits;
        std::vector<SeriesDataLabelsEdit> seriesDataLabelsEdits;
        std::vector<PointDataLabelEdit> pointDataLabelEdits;
        std::vector<PointDataLabelRichTextEdit> pointDataLabelRichTextEdits;
        std::vector<DataPointFormatEdit> dataPointFormatEdits;
        std::vector<SeriesFormatEdit> seriesFormatEdits;
        std::vector<TrendlineFormatEdit> trendlineFormatEdits;
        std::vector<ErrorBarsFormatEdit> errorBarsFormatEdits;
        std::vector<TrendlineEdit> trendlineEdits;
        std::vector<ErrorBarsEdit> errorBarsEdits;
        bool legendChanged{false};
        bool showLegend{true};
        std::string legendPosition{"r"};
        bool removed{false};
        std::vector<SeriesTitleEdit> seriesTitleEdits;
        std::vector<SeriesReferenceEdit> seriesReferenceEdits;
        std::vector<SeriesCacheEdit> seriesCacheEdits;
    };

    ImportedImageEdit& ensureImportedImageEdit(const Image& image);
    ImportedChartEdit& ensureImportedChartEdit(const Chart& chart);
    friend class Workbook;
    friend struct internal::WorkbookDrawingAccess;

public:
    // --- Sparklines (x14 extension) ---
    // openpyxl does not model sparklines at all. XL++ models the common
    // sparklineGroup/sparkline structure so callers can author or edit
    // Excel-style inline sparklines while untouched extension XML stays
    // byte-preserved.
    const std::vector<SparklineGroup>& sparklineGroups() const noexcept { return sparklineGroups_; }
    std::vector<SparklineGroup>& sparklineGroups() noexcept { dirty_ = true; sparklinesDirty_ = true; return sparklineGroups_; }
    void setSparklineGroups(std::vector<SparklineGroup> groups) {
        sparklineGroups_ = std::move(groups);
        dirty_ = true;
        sparklinesDirty_ = true;
    }
    bool hasSparklines() const noexcept { return !sparklineGroups_.empty() || !sparklinesRawXml_.empty(); }
    const std::string& sparklineGroupsRawXml() const noexcept { return sparklinesRawXml_; }
    void setSparklineGroupsRawXml(std::string rawXml) {
        sparklinesRawXml_ = std::move(rawXml);
        dirty_ = true;
        sparklinesDirty_ = true;
    }

private:
    PrintOptions printOptions_;
    HeaderFooter headerFooter_;
    WorksheetProtection protection_;
    std::vector<Image> images_;
    SheetView sheetView_;
    static void syncBreakList(const std::set<std::size_t>& source, std::vector<std::size_t>& target) {
        target.assign(source.begin(), source.end());
    }
    std::string printArea_, printTitlesRows_, printTitlesCols_;
    std::set<std::size_t> rowBreaks_, columnBreaks_;
    std::vector<std::size_t> rowBreakList_, columnBreakList_;
    std::vector<Chart> charts_;
    std::vector<PivotTable> pivotTables_;
    std::size_t loadedPivotCount_{0};
    mutable bool dirty_{true};
    mutable std::set<std::uint64_t> trackedCellKeys_;
    // Separate mutation flags allow package-level copy-on-write preservation:
    // unrelated cell edits do not force imported drawings or pivots to be
    // reserialized through XLPP's intentionally smaller object model.
    mutable bool drawingsDirty_{true};
    mutable bool drawingAppendDirty_{false};
    std::size_t loadedImageCount_{0};
    std::vector<ImportedImageEdit> importedImageEdits_;
    std::size_t loadedChartCount_{0};
    std::vector<ImportedChartEdit> importedChartEdits_;
    mutable bool pivotsDirty_{true};
    mutable bool pivotAppendDirty_{false};
    std::vector<SparklineGroup> sparklineGroups_;
    std::string sparklinesRawXml_;
    mutable bool sparklinesDirty_{true};
};
}
