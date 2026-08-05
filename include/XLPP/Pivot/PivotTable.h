#pragma once
#include <string>
#include <stdexcept>
#include <vector>

namespace xlpp {

class PivotCache {
public:
    int cacheId() const noexcept { return cacheId_; } void setCacheId(int v) noexcept { cacheId_ = v; }
    const std::string& sourceData() const noexcept { return sourceData_; }
    void setSourceData(std::string v) { sourceData_ = std::move(v); }

    const std::vector<std::string>& fields() const noexcept { return fields_; }
    std::vector<std::string>& fields() noexcept { return fields_; }
    void setFields(std::vector<std::string> v) { fields_ = std::move(v); }
    void addField(std::string v) { fields_.push_back(std::move(v)); }

    const std::vector<std::vector<std::string>>& records() const noexcept { return records_; }
    std::vector<std::vector<std::string>>& records() noexcept { return records_; }
    void setRecords(std::vector<std::vector<std::string>> v) {
        if (!fields_.empty()) {
            for (const auto& record : v)
                if (record.size() != fields_.size()) throw std::invalid_argument("Pivot cache record width must match field count");
        }
        records_ = std::move(v);
    }
    void addRecord(std::vector<std::string> v) {
        if (!fields_.empty() && v.size() != fields_.size())
            throw std::invalid_argument("Pivot cache record width must match field count");
        records_.push_back(std::move(v));
    }

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

    const std::string& name() const noexcept { return name_; } void setName(std::string v) { name_ = std::move(v); }
    const std::string& axis() const noexcept { return axis_; } void setAxis(std::string v) { axis_ = std::move(v); }
    bool showAll() const noexcept { return showAll_; } void setShowAll(bool v) noexcept { showAll_ = v; }
    int sortType() const noexcept { return sortType_; } void setSortType(int v) noexcept { sortType_ = v; }
    int fieldIndex() const noexcept { return fieldIndex_; } void setFieldIndex(int v) noexcept { fieldIndex_ = v; }

private:
    std::string name_, axis_{"axisRow"};
    bool showAll_{true};
    int sortType_{0};
    int fieldIndex_{-1};
};

class PivotFieldReference {
public:
    int fieldIndex() const noexcept { return fieldIndex_; } void setFieldIndex(int v) noexcept { fieldIndex_ = v; }

private:
    int fieldIndex_{0};
};

class PivotTable {
public:
    PivotTable() = default;
    explicit PivotTable(std::string name) : name_(std::move(name)) {}

    const std::string& name() const noexcept { return name_; } void setName(std::string v) { name_ = std::move(v); }
    const std::string& location() const noexcept { return location_; }
    void setLocation(std::string v) { location_ = std::move(v); }

    PivotCache& cache() noexcept { return cache_; }
    const PivotCache& cache() const noexcept { return cache_; }

    void addRowField(std::string name) { rowFields_.push_back(PivotField(std::move(name))); }
    void addColumnField(std::string name) { columnFields_.push_back(PivotField(std::move(name))); columnFields_.back().setAxis("axisCol"); }
    void addPageField(std::string name) { pageFields_.push_back(PivotField(std::move(name))); pageFields_.back().setAxis("axisPage"); }
    void addDataField(int fieldIndex = 0) { PivotFieldReference ref; ref.setFieldIndex(fieldIndex); dataFields_.push_back(ref); }

    const std::vector<PivotField>& rowFields() const noexcept { return rowFields_; }
    const std::vector<PivotField>& columnFields() const noexcept { return columnFields_; }
    const std::vector<PivotField>& pageFields() const noexcept { return pageFields_; }
    const std::vector<PivotFieldReference>& dataFields() const noexcept { return dataFields_; }

    std::vector<PivotField>& rowFields() noexcept { return rowFields_; }
    std::vector<PivotField>& columnFields() noexcept { return columnFields_; }
    std::vector<PivotField>& pageFields() noexcept { return pageFields_; }
    std::vector<PivotFieldReference>& dataFields() noexcept { return dataFields_; }

private:
    std::string name_, location_;
    PivotCache cache_;
    std::vector<PivotField> rowFields_, columnFields_, pageFields_;
    std::vector<PivotFieldReference> dataFields_;
};

}
