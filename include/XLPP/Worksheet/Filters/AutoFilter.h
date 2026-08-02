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

class FilterColumn {
public:
    explicit FilterColumn(std::size_t columnId = 0) : columnId_(columnId) {}
    std::size_t columnId() const noexcept { return columnId_; }
    void setColumnId(std::size_t value) noexcept { columnId_ = value; }
    void addValue(std::string value) { values_.push_back(std::move(value)); }
    void clearValues() noexcept { values_.clear(); }
    const std::vector<std::string>& values() const noexcept { return values_; }
    void addCustomFilter(FilterOperator op, std::string value) { customFilters_.push_back({op, std::move(value)}); }
    void clearCustomFilters() noexcept { customFilters_.clear(); }
    const std::vector<CustomFilter>& customFilters() const noexcept { return customFilters_; }
    void setAndMode(bool value) noexcept { andMode_ = value; }
    bool andMode() const noexcept { return andMode_; }
    void setIncludeBlank(bool value) noexcept { includeBlank_ = value; }
    bool includeBlank() const noexcept { return includeBlank_; }
private:
    std::size_t columnId_{0};
    std::vector<std::string> values_;
    std::vector<CustomFilter> customFilters_;
    bool andMode_{false};
    bool includeBlank_{false};
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
