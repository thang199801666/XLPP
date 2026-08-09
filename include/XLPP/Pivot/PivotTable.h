#pragma once
#include <algorithm>
#include <cstdint>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {

struct PivotGrouping {
    enum class Kind { None, Numeric, Date };
    enum class DatePart { Seconds, Minutes, Hours, Days, Months, Quarters, Years };

    Kind kind{Kind::None};
    bool autoStart{true};
    bool autoEnd{true};
    double start{0.0};
    double end{0.0};
    double interval{1.0};
    DatePart datePart{DatePart::Months};
    std::string startDate;
    std::string endDate;

    bool active() const noexcept { return kind != Kind::None; }
};

struct PivotFilter {
    // Names intentionally follow ECMA-376 ST_PivotFilterType values where possible.
    std::string type{"unknown"};
    int fieldIndex{-1};
    int measureFieldIndex{-1};
    std::string value1;
    std::string value2;
    double top10Value{10.0};
    bool top10Percent{false};
    bool top10Top{true};
};

class PivotCache {
public:
    int cacheId() const noexcept { return cacheId_; }
    void setCacheId(int value) {
        if (value <= 0) throw std::invalid_argument("Pivot cache ID must be positive");
        cacheId_ = value;
    }

    const std::string& sourceData() const noexcept { return sourceData_; }
    void setSourceData(std::string value) { sourceData_ = std::move(value); }

    const std::vector<std::string>& fields() const noexcept { return fields_; }
    std::vector<std::string>& fields() noexcept { return fields_; }
    void setFields(std::vector<std::string> value) {
        if (!records_.empty()) {
            for (const auto& record : records_) {
                if (record.size() != value.size())
                    throw std::invalid_argument("Pivot cache record width must match field count");
            }
        }
        fields_ = std::move(value);
    }
    void addField(std::string value) {
        if (!records_.empty())
            throw std::logic_error("Cannot add a pivot cache field after records have been added");
        fields_.push_back(std::move(value));
    }

    int fieldIndex(const std::string& name) const noexcept {
        const auto it = std::find(fields_.begin(), fields_.end(), name);
        return it == fields_.end() ? -1 : static_cast<int>(std::distance(fields_.begin(), it));
    }

    const std::vector<std::vector<std::string>>& records() const noexcept { return records_; }
    std::vector<std::vector<std::string>>& records() noexcept { return records_; }
    void setRecords(std::vector<std::vector<std::string>> value) {
        if (!fields_.empty()) {
            for (const auto& record : value) {
                if (record.size() != fields_.size())
                    throw std::invalid_argument("Pivot cache record width must match field count");
            }
        }
        records_ = std::move(value);
    }
    void addRecord(std::vector<std::string> value) {
        if (!fields_.empty() && value.size() != fields_.size())
            throw std::invalid_argument("Pivot cache record width must match field count");
        records_.push_back(std::move(value));
    }
    void clearRecords() noexcept { records_.clear(); }

    bool refreshOnLoad() const noexcept { return refreshOnLoad_; }
    void setRefreshOnLoad(bool value) noexcept { refreshOnLoad_ = value; }
    bool saveData() const noexcept { return saveData_; }
    void setSaveData(bool value) noexcept { saveData_ = value; }
    bool enableRefresh() const noexcept { return enableRefresh_; }
    void setEnableRefresh(bool value) noexcept { enableRefresh_ = value; }
    int missingItemsLimit() const noexcept { return missingItemsLimit_; }
    void setMissingItemsLimit(int value) { if (value < 0) throw std::invalid_argument("Pivot missing-items limit cannot be negative"); missingItemsLimit_ = value; }
    bool backgroundQuery() const noexcept { return backgroundQuery_; }
    void setBackgroundQuery(bool value) noexcept { backgroundQuery_ = value; }
    bool optimizeMemory() const noexcept { return optimizeMemory_; }
    void setOptimizeMemory(bool value) noexcept { optimizeMemory_ = value; }
    bool upgradeOnRefresh() const noexcept { return upgradeOnRefresh_; }
    void setUpgradeOnRefresh(bool value) noexcept { upgradeOnRefresh_ = value; }
    bool supportSubquery() const noexcept { return supportSubquery_; }
    void setSupportSubquery(bool value) noexcept { supportSubquery_ = value; }
    bool supportAdvancedDrill() const noexcept { return supportAdvancedDrill_; }
    void setSupportAdvancedDrill(bool value) noexcept { supportAdvancedDrill_ = value; }
    const std::string& refreshedBy() const noexcept { return refreshedBy_; }
    void setRefreshedBy(std::string value) { refreshedBy_ = std::move(value); }

private:
    int cacheId_{1};
    std::string sourceData_;
    std::vector<std::string> fields_;
    std::vector<std::vector<std::string>> records_;
    bool refreshOnLoad_{true};
    bool saveData_{true};
    bool enableRefresh_{true};
    int missingItemsLimit_{0};
    bool backgroundQuery_{false};
    bool optimizeMemory_{false};
    bool upgradeOnRefresh_{false};
    bool supportSubquery_{true};
    bool supportAdvancedDrill_{true};
    std::string refreshedBy_;
};

class PivotField {
public:
    PivotField() = default;
    explicit PivotField(std::string name) : name_(std::move(name)) {}

    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
    const std::string& axis() const noexcept { return axis_; }
    void setAxis(std::string value) { axis_ = std::move(value); }
    bool showAll() const noexcept { return showAll_; }
    void setShowAll(bool value) noexcept { showAll_ = value; }
    int sortType() const noexcept { return sortType_; }
    void setSortType(int value) noexcept { sortType_ = value; }
    bool subtotalTop() const noexcept { return subtotalTop_; }
    void setSubtotalTop(bool value) noexcept { subtotalTop_ = value; }
    bool insertBlankRow() const noexcept { return insertBlankRow_; }
    void setInsertBlankRow(bool value) noexcept { insertBlankRow_ = value; }
    bool repeatItemLabels() const noexcept { return repeatItemLabels_; }
    void setRepeatItemLabels(bool value) noexcept { repeatItemLabels_ = value; }
    bool includeNewItemsInFilter() const noexcept { return includeNewItemsInFilter_; }
    void setIncludeNewItemsInFilter(bool value) noexcept { includeNewItemsInFilter_ = value; }
    bool multipleItemSelectionAllowed() const noexcept { return multipleItemSelectionAllowed_; }
    void setMultipleItemSelectionAllowed(bool value) noexcept { multipleItemSelectionAllowed_ = value; }
    int selectedItemIndex() const noexcept { return selectedItemIndex_; }
    void setSelectedItemIndex(int value) noexcept { selectedItemIndex_ = value; }
    int fieldIndex() const noexcept { return fieldIndex_; }
    void setFieldIndex(int value) noexcept { fieldIndex_ = value; }
    bool compact() const noexcept { return compact_; }
    void setCompact(bool value) noexcept { compact_ = value; }
    bool outline() const noexcept { return outline_; }
    void setOutline(bool value) noexcept { outline_ = value; }
    bool insertPageBreak() const noexcept { return insertPageBreak_; }
    void setInsertPageBreak(bool value) noexcept { insertPageBreak_ = value; }
    bool showDropDowns() const noexcept { return showDropDowns_; }
    void setShowDropDowns(bool value) noexcept { showDropDowns_ = value; }
    bool defaultSubtotal() const noexcept { return defaultSubtotal_; }
    void setDefaultSubtotal(bool value) noexcept { defaultSubtotal_ = value; }
    const std::vector<std::string>& subtotals() const noexcept { return subtotals_; }
    void setSubtotals(std::vector<std::string> value) { subtotals_ = std::move(value); }
    void addSubtotal(std::string value) {
        static const std::vector<std::string> supported{"sum","countA","avg","max","min","product","count","stdDev","stdDevP","var","varP"};
        if (std::find(supported.begin(), supported.end(), value) == supported.end()) throw std::invalid_argument("Unsupported pivot subtotal: " + value);
        if (std::find(subtotals_.begin(), subtotals_.end(), value) == subtotals_.end()) subtotals_.push_back(std::move(value));
    }
    const std::vector<int>& hiddenItemIndexes() const noexcept { return hiddenItemIndexes_; }
    void setHiddenItemIndexes(std::vector<int> value) {
        if (std::any_of(value.begin(), value.end(), [](int i){ return i < 0; })) throw std::invalid_argument("Pivot hidden item index cannot be negative");
        std::sort(value.begin(), value.end()); value.erase(std::unique(value.begin(), value.end()), value.end()); hiddenItemIndexes_ = std::move(value);
    }
    bool itemHidden(int index) const noexcept { return std::find(hiddenItemIndexes_.begin(), hiddenItemIndexes_.end(), index) != hiddenItemIndexes_.end(); }
    void setItemHidden(int index, bool hidden = true) {
        if (index < 0) throw std::invalid_argument("Pivot item index cannot be negative");
        auto it = std::find(hiddenItemIndexes_.begin(), hiddenItemIndexes_.end(), index);
        if (hidden && it == hiddenItemIndexes_.end()) { hiddenItemIndexes_.push_back(index); std::sort(hiddenItemIndexes_.begin(), hiddenItemIndexes_.end()); }
        else if (!hidden && it != hiddenItemIndexes_.end()) hiddenItemIndexes_.erase(it);
    }
    const PivotGrouping& grouping() const noexcept { return grouping_; }
    PivotGrouping& grouping() noexcept { return grouping_; }
    void setGrouping(PivotGrouping value) noexcept { grouping_ = value; }

private:
    std::string name_;
    std::string axis_{"axisRow"};
    bool showAll_{false};
    int sortType_{0}; // 0 manual, 1 ascending, 2 descending
    bool subtotalTop_{true};
    bool insertBlankRow_{false};
    bool repeatItemLabels_{false};
    bool includeNewItemsInFilter_{false};
    bool multipleItemSelectionAllowed_{false};
    int selectedItemIndex_{-1};
    int fieldIndex_{-1};
    bool compact_{true};
    bool outline_{true};
    bool insertPageBreak_{false};
    bool showDropDowns_{true};
    bool defaultSubtotal_{false};
    std::vector<std::string> subtotals_;
    std::vector<int> hiddenItemIndexes_;
    PivotGrouping grouping_{};
};

class PivotFieldReference {
public:
    int fieldIndex() const noexcept { return fieldIndex_; }
    void setFieldIndex(int value) noexcept { fieldIndex_ = value; }
    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
    const std::string& subtotal() const noexcept { return subtotal_; }
    void setSubtotal(std::string value) {
        static const std::vector<std::string> supported{
            "average", "count", "countNums", "max", "min", "product",
            "stdDev", "stdDevp", "sum", "var", "varp"
        };
        if (std::find(supported.begin(), supported.end(), value) == supported.end())
            throw std::invalid_argument("Unsupported pivot data aggregation: " + value);
        subtotal_ = std::move(value);
    }
    const std::string& caption() const noexcept { return caption_; }
    void setCaption(std::string value) { caption_ = std::move(value); }
    std::uint32_t numberFormatId() const noexcept { return numberFormatId_; }
    void setNumberFormatId(std::uint32_t value) noexcept { numberFormatId_ = value; }
    const std::string& showDataAs() const noexcept { return showDataAs_; }
    void setShowDataAs(std::string value) {
        static const std::vector<std::string> supported{"normal","difference","percent","percentDiff","runTotal","percentOfRow","percentOfCol","percentOfTotal","index",
            "percentOfParent","percentOfParentRow","percentOfParentCol","percentOfRunningTotal","rankAscending","rankDescending"};
        if (std::find(supported.begin(), supported.end(), value) == supported.end()) throw std::invalid_argument("Unsupported pivot show-data-as mode: " + value);
        showDataAs_ = std::move(value);
    }
    int baseField() const noexcept { return baseField_; }
    void setBaseField(int value) noexcept { baseField_ = value; }
    int baseItem() const noexcept { return baseItem_; }
    void setBaseItem(int value) noexcept { baseItem_ = value; }

private:
    int fieldIndex_{-1};
    std::string name_;
    std::string subtotal_{"sum"};
    std::string caption_;
    std::uint32_t numberFormatId_{0};
    std::string showDataAs_{"normal"};
    int baseField_{0};
    int baseItem_{0};
};

enum class PivotLayout { Compact, Outline, Tabular };

class PivotTable {
public:
    PivotTable() = default;
    explicit PivotTable(std::string name) : name_(std::move(name)) {}

    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
    const std::string& location() const noexcept { return location_; }
    void setLocation(std::string value) { location_ = std::move(value); }
    PivotLayout layout() const noexcept { return layout_; }
    void setLayout(PivotLayout value) noexcept { layout_ = value; }
    bool rowGrandTotals() const noexcept { return rowGrandTotals_; }
    void setRowGrandTotals(bool value) noexcept { rowGrandTotals_ = value; }
    bool columnGrandTotals() const noexcept { return columnGrandTotals_; }
    void setColumnGrandTotals(bool value) noexcept { columnGrandTotals_ = value; }
    bool preserveFormatting() const noexcept { return preserveFormatting_; }
    void setPreserveFormatting(bool value) noexcept { preserveFormatting_ = value; }
    bool useAutoFormatting() const noexcept { return useAutoFormatting_; }
    void setUseAutoFormatting(bool value) noexcept { useAutoFormatting_ = value; }
    const std::string& dataCaption() const noexcept { return dataCaption_; }
    void setDataCaption(std::string value) { dataCaption_ = std::move(value); }
    const std::string& styleName() const noexcept { return styleName_; }
    void setStyleName(std::string value) { styleName_ = std::move(value); }
    bool showRowHeaders() const noexcept { return showRowHeaders_; }
    void setShowRowHeaders(bool value) noexcept { showRowHeaders_ = value; }
    bool showColumnHeaders() const noexcept { return showColumnHeaders_; }
    void setShowColumnHeaders(bool value) noexcept { showColumnHeaders_ = value; }
    bool showRowStripes() const noexcept { return showRowStripes_; }
    void setShowRowStripes(bool value) noexcept { showRowStripes_ = value; }
    bool showColumnStripes() const noexcept { return showColumnStripes_; }
    void setShowColumnStripes(bool value) noexcept { showColumnStripes_ = value; }
    bool showLastColumn() const noexcept { return showLastColumn_; }
    void setShowLastColumn(bool value) noexcept { showLastColumn_ = value; }
    bool showEmptyRow() const noexcept { return showEmptyRow_; }
    void setShowEmptyRow(bool value) noexcept { showEmptyRow_ = value; }
    bool showEmptyColumn() const noexcept { return showEmptyColumn_; }
    void setShowEmptyColumn(bool value) noexcept { showEmptyColumn_ = value; }
    bool showDrill() const noexcept { return showDrill_; }
    void setShowDrill(bool value) noexcept { showDrill_ = value; }
    bool enableDrill() const noexcept { return enableDrill_; }
    void setEnableDrill(bool value) noexcept { enableDrill_ = value; }
    bool showDataTips() const noexcept { return showDataTips_; }
    void setShowDataTips(bool value) noexcept { showDataTips_ = value; }
    bool showMemberPropertyTips() const noexcept { return showMemberPropertyTips_; }
    void setShowMemberPropertyTips(bool value) noexcept { showMemberPropertyTips_ = value; }
    bool showHeaders() const noexcept { return showHeaders_; }
    void setShowHeaders(bool value) noexcept { showHeaders_ = value; }
    bool multipleFieldFilters() const noexcept { return multipleFieldFilters_; }
    void setMultipleFieldFilters(bool value) noexcept { multipleFieldFilters_ = value; }
    bool showValuesRow() const noexcept { return showValuesRow_; }
    void setShowValuesRow(bool value) noexcept { showValuesRow_ = value; }
    bool subtotalHiddenItems() const noexcept { return subtotalHiddenItems_; }
    void setSubtotalHiddenItems(bool value) noexcept { subtotalHiddenItems_ = value; }
    int pageWrap() const noexcept { return pageWrap_; }
    void setPageWrap(int value) { if (value < 0) throw std::invalid_argument("Pivot page-wrap cannot be negative"); pageWrap_ = value; }
    bool pageOverThenDown() const noexcept { return pageOverThenDown_; }
    void setPageOverThenDown(bool value) noexcept { pageOverThenDown_ = value; }
    const std::vector<PivotFilter>& filters() const noexcept { return filters_; }
    std::vector<PivotFilter>& filters() noexcept { return filters_; }
    PivotFilter& addFilter(PivotFilter filter) { filters_.push_back(std::move(filter)); return filters_.back(); }

    PivotCache& cache() noexcept { return cache_; }
    const PivotCache& cache() const noexcept { return cache_; }

    PivotField& addRowField(std::string name) {
        rowFields_.emplace_back(std::move(name));
        rowFields_.back().setAxis("axisRow");
        rowFields_.back().setFieldIndex(cache_.fieldIndex(rowFields_.back().name()));
        return rowFields_.back();
    }
    PivotField& addColumnField(std::string name) {
        columnFields_.emplace_back(std::move(name));
        columnFields_.back().setAxis("axisCol");
        columnFields_.back().setFieldIndex(cache_.fieldIndex(columnFields_.back().name()));
        return columnFields_.back();
    }
    PivotField& addPageField(std::string name) {
        pageFields_.emplace_back(std::move(name));
        pageFields_.back().setAxis("axisPage");
        pageFields_.back().setFieldIndex(cache_.fieldIndex(pageFields_.back().name()));
        return pageFields_.back();
    }
    PivotFieldReference& addDataField(int fieldIndex = 0) {
        PivotFieldReference reference;
        reference.setFieldIndex(fieldIndex);
        if (fieldIndex >= 0 && static_cast<std::size_t>(fieldIndex) < cache_.fields().size())
            reference.setName(cache_.fields()[static_cast<std::size_t>(fieldIndex)]);
        dataFields_.push_back(std::move(reference));
        return dataFields_.back();
    }
    PivotFieldReference& addDataField(std::string name, std::string subtotal = "sum") {
        PivotFieldReference reference;
        reference.setName(std::move(name));
        reference.setFieldIndex(cache_.fieldIndex(reference.name()));
        reference.setSubtotal(std::move(subtotal));
        dataFields_.push_back(std::move(reference));
        return dataFields_.back();
    }

    const std::vector<PivotField>& rowFields() const noexcept { return rowFields_; }
    const std::vector<PivotField>& columnFields() const noexcept { return columnFields_; }
    const std::vector<PivotField>& pageFields() const noexcept { return pageFields_; }
    const std::vector<PivotFieldReference>& dataFields() const noexcept { return dataFields_; }

    std::vector<PivotField>& rowFields() noexcept { return rowFields_; }
    std::vector<PivotField>& columnFields() noexcept { return columnFields_; }
    std::vector<PivotField>& pageFields() noexcept { return pageFields_; }
    std::vector<PivotFieldReference>& dataFields() noexcept { return dataFields_; }

private:
    std::string name_;
    std::string location_;
    PivotLayout layout_{PivotLayout::Compact};
    bool rowGrandTotals_{true};
    bool columnGrandTotals_{true};
    bool preserveFormatting_{true};
    bool useAutoFormatting_{true};
    std::string dataCaption_{"Values"};
    std::string styleName_{"PivotStyleLight16"};
    bool showRowHeaders_{true};
    bool showColumnHeaders_{true};
    bool showRowStripes_{false};
    bool showColumnStripes_{false};
    bool showLastColumn_{true};
    bool showEmptyRow_{false};
    bool showEmptyColumn_{false};
    bool showDrill_{true};
    bool enableDrill_{true};
    bool showDataTips_{true};
    bool showMemberPropertyTips_{true};
    bool showHeaders_{true};
    bool multipleFieldFilters_{false};
    bool showValuesRow_{false};
    bool subtotalHiddenItems_{false};
    int pageWrap_{0};
    bool pageOverThenDown_{false};
    std::vector<PivotFilter> filters_;
    PivotCache cache_;
    std::vector<PivotField> rowFields_;
    std::vector<PivotField> columnFields_;
    std::vector<PivotField> pageFields_;
    std::vector<PivotFieldReference> dataFields_;
};

} // namespace xlpp
