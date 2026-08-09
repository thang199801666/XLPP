#pragma once
#include <cstddef>
#include <map>
#include <optional>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {
enum class FilterOperator { Equal, NotEqual, LessThan, LessThanOrEqual, GreaterThan, GreaterThanOrEqual };

struct CustomFilter {
    FilterOperator op{FilterOperator::Equal};
    std::string value;
};

enum class DynamicFilterType {
    AboveAverage, BelowAverage, Tomorrow, Today, Yesterday,
    NextWeek, ThisWeek, LastWeek, NextMonth, ThisMonth, LastMonth,
    NextQuarter, ThisQuarter, LastQuarter, NextYear, ThisYear, LastYear,
    YearToDate, Quarter1, Quarter2, Quarter3, Quarter4,
    Month1, Month2, Month3, Month4, Month5, Month6,
    Month7, Month8, Month9, Month10, Month11, Month12
};

struct DynamicFilter {
    DynamicFilterType type{DynamicFilterType::Today};
    std::optional<double> value;
    std::optional<double> maxValue;
};

struct Top10Filter {
    bool top{true};
    bool percent{false};
    double value{10.0};
    std::optional<double> filterValue;
};

struct ColorFilter {
    std::size_t dxfId{0};
    bool cellColor{true};
};

struct IconFilter {
    std::string iconSet{"3Arrows"};
    std::size_t iconId{0};
};

enum class DateTimeGrouping { Year, Month, Day, Hour, Minute, Second };

struct DateGroupItem {
    int year{0};
    std::optional<int> month;
    std::optional<int> day;
    std::optional<int> hour;
    std::optional<int> minute;
    std::optional<int> second;
    DateTimeGrouping grouping{DateTimeGrouping::Year};
};

class FilterColumn {
public:
    explicit FilterColumn(std::size_t columnId = 0) : columnId_(columnId) {}
    std::size_t columnId() const noexcept { return columnId_; }
    void setColumnId(std::size_t value) noexcept { columnId_ = value; }
    void addValue(std::string value) { values_.push_back(std::move(value)); }
    void clearValues() noexcept { values_.clear(); }
    const std::vector<std::string>& values() const noexcept { return values_; }
    void addDateGroup(DateGroupItem value) { dateGroups_.push_back(std::move(value)); }
    void clearDateGroups() noexcept { dateGroups_.clear(); }
    const std::vector<DateGroupItem>& dateGroups() const noexcept { return dateGroups_; }
    void addCustomFilter(FilterOperator op, std::string value) { customFilters_.push_back({op, std::move(value)}); }
    void clearCustomFilters() noexcept { customFilters_.clear(); }
    const std::vector<CustomFilter>& customFilters() const noexcept { return customFilters_; }
    void setAndMode(bool value) noexcept { andMode_ = value; }
    bool andMode() const noexcept { return andMode_; }
    void setIncludeBlank(bool value) noexcept { includeBlank_ = value; }
    bool includeBlank() const noexcept { return includeBlank_; }

    void setDynamicFilter(DynamicFilter value) { dynamicFilter_ = std::move(value); clearExclusiveAdvancedFilters(1); }
    const std::optional<DynamicFilter>& dynamicFilter() const noexcept { return dynamicFilter_; }
    void clearDynamicFilter() noexcept { dynamicFilter_.reset(); }
    void setTop10Filter(Top10Filter value) { top10Filter_ = std::move(value); clearExclusiveAdvancedFilters(2); }
    const std::optional<Top10Filter>& top10Filter() const noexcept { return top10Filter_; }
    void clearTop10Filter() noexcept { top10Filter_.reset(); }
    void setColorFilter(ColorFilter value) { colorFilter_ = std::move(value); clearExclusiveAdvancedFilters(3); }
    const std::optional<ColorFilter>& colorFilter() const noexcept { return colorFilter_; }
    void clearColorFilter() noexcept { colorFilter_.reset(); }
    void setIconFilter(IconFilter value) { iconFilter_ = std::move(value); clearExclusiveAdvancedFilters(4); }
    const std::optional<IconFilter>& iconFilter() const noexcept { return iconFilter_; }
    void clearIconFilter() noexcept { iconFilter_.reset(); }

private:
    void clearExclusiveAdvancedFilters(int keep) noexcept {
        if (keep != 0) { values_.clear(); dateGroups_.clear(); customFilters_.clear(); includeBlank_ = false; }
        if (keep != 1) dynamicFilter_.reset();
        if (keep != 2) top10Filter_.reset();
        if (keep != 3) colorFilter_.reset();
        if (keep != 4) iconFilter_.reset();
    }

    std::size_t columnId_{0};
    std::vector<std::string> values_;
    std::vector<DateGroupItem> dateGroups_;
    std::vector<CustomFilter> customFilters_;
    bool andMode_{false};
    bool includeBlank_{false};
    std::optional<DynamicFilter> dynamicFilter_;
    std::optional<Top10Filter> top10Filter_;
    std::optional<ColorFilter> colorFilter_;
    std::optional<IconFilter> iconFilter_;
};

struct SortCondition {
    std::string reference;
    bool descending{false};
};

class SortState {
public:
    void setReference(std::string value) { reference_ = std::move(value); }
    const std::string& reference() const noexcept { return reference_; }
    void setCaseSensitive(bool value) noexcept { caseSensitive_ = value; }
    bool caseSensitive() const noexcept { return caseSensitive_; }
    void addCondition(std::string reference, bool descending = false) { conditions_.push_back({std::move(reference), descending}); }
    const std::vector<SortCondition>& conditions() const noexcept { return conditions_; }
    void clear() noexcept { reference_.clear(); conditions_.clear(); caseSensitive_ = false; }
private:
    std::string reference_;
    std::vector<SortCondition> conditions_;
    bool caseSensitive_{false};
};

class AutoFilter {
public:
    void setReference(std::string value) { reference_ = std::move(value); }
    const std::string& reference() const noexcept { return reference_; }
    bool enabled() const noexcept { return !reference_.empty(); }
    void clear() noexcept { reference_.clear(); columns_.clear(); sortState_.reset(); }
    FilterColumn& column(std::size_t columnId) { return columns_.try_emplace(columnId, columnId).first->second; }
    const FilterColumn* tryColumn(std::size_t columnId) const noexcept { auto it = columns_.find(columnId); return it == columns_.end() ? nullptr : &it->second; }
    const std::map<std::size_t, FilterColumn>& columns() const noexcept { return columns_; }
    SortState& sortState() { if (!sortState_) sortState_.emplace(); return *sortState_; }
    const std::optional<SortState>& sortStateValue() const noexcept { return sortState_; }
private:
    std::string reference_;
    std::map<std::size_t, FilterColumn> columns_;
    std::optional<SortState> sortState_;
};
}
