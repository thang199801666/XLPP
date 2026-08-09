#pragma once
#include <XLPP/Core/StableVector.h>
#include <string>
#include <vector>
#include <stdexcept>

namespace xlpp {
class TableStyleInfo {
public:
    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
    bool showFirstColumn() const noexcept { return showFirstColumn_; }
    void setShowFirstColumn(bool value) noexcept { showFirstColumn_ = value; }
    bool showLastColumn() const noexcept { return showLastColumn_; }
    void setShowLastColumn(bool value) noexcept { showLastColumn_ = value; }
    bool showRowStripes() const noexcept { return showRowStripes_; }
    void setShowRowStripes(bool value) noexcept { showRowStripes_ = value; }
    bool showColumnStripes() const noexcept { return showColumnStripes_; }
    void setShowColumnStripes(bool value) noexcept { showColumnStripes_ = value; }
private:
    std::string name_ = "TableStyleMedium2";
    bool showFirstColumn_ = false;
    bool showLastColumn_ = false;
    bool showRowStripes_ = true;
    bool showColumnStripes_ = false;
};

class TableColumn {
public:
    TableColumn() = default;
    TableColumn(std::size_t id, std::string name) : id_(id), name_(std::move(name)) {}
    std::size_t id() const noexcept { return id_; }
    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { if (value.empty()) throw std::invalid_argument("Table column name cannot be empty"); name_ = std::move(value); }
private:
    std::size_t id_ = 0;
    std::string name_;
};

class Table {
public:
    Table() = default;
    Table(std::string name, std::string reference) : name_(std::move(name)), displayName_(name_), reference_(std::move(reference)) {
        if (name_.empty() || reference_.empty()) throw std::invalid_argument("Table name and reference cannot be empty");
    }
    const std::string& name() const noexcept { return name_; }
    const std::string& displayName() const noexcept { return displayName_; }
    void setDisplayName(std::string value) { if (value.empty()) throw std::invalid_argument("Table display name cannot be empty"); displayName_ = std::move(value); }
    const std::string& reference() const noexcept { return reference_; }
    void setReference(std::string value) { if (value.empty()) throw std::invalid_argument("Table reference cannot be empty"); reference_ = std::move(value); }
    bool showHeaderRow() const noexcept { return showHeaderRow_; }
    void setShowHeaderRow(bool value) noexcept { showHeaderRow_ = value; }
    bool showTotalsRow() const noexcept { return showTotalsRow_; }
    void setShowTotalsRow(bool value) noexcept { showTotalsRow_ = value; }
    StableVector<TableColumn>& columns() noexcept { return columns_; }
    const StableVector<TableColumn>& columns() const noexcept { return columns_; }
    TableColumn& addColumn(std::string name) { columns_.emplace_back(columns_.size()+1, std::move(name)); return columns_.back(); }
    TableStyleInfo& styleInfo() noexcept { return styleInfo_; }
    const TableStyleInfo& styleInfo() const noexcept { return styleInfo_; }
private:
    std::string name_;
    std::string displayName_;
    std::string reference_;
    bool showHeaderRow_ = true;
    bool showTotalsRow_ = false;
    StableVector<TableColumn> columns_;
    TableStyleInfo styleInfo_;
};
}
