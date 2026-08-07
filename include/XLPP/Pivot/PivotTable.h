#pragma once
#include <algorithm>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {

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

private:
    int cacheId_{1};
    std::string sourceData_;
    std::vector<std::string> fields_;
    std::vector<std::vector<std::string>> records_;
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
    int fieldIndex() const noexcept { return fieldIndex_; }
    void setFieldIndex(int value) noexcept { fieldIndex_ = value; }

private:
    std::string name_;
    std::string axis_{"axisRow"};
    bool showAll_{false};
    int sortType_{0};
    int fieldIndex_{-1};
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

private:
    int fieldIndex_{-1};
    std::string name_;
    std::string subtotal_{"sum"};
};

class PivotTable {
public:
    PivotTable() = default;
    explicit PivotTable(std::string name) : name_(std::move(name)) {}

    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
    const std::string& location() const noexcept { return location_; }
    void setLocation(std::string value) { location_ = std::move(value); }

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
    PivotCache cache_;
    std::vector<PivotField> rowFields_;
    std::vector<PivotField> columnFields_;
    std::vector<PivotField> pageFields_;
    std::vector<PivotFieldReference> dataFields_;
};

} // namespace xlpp
