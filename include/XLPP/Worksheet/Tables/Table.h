#pragma once
#include <algorithm>
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
    // Totals-row aggregation function, e.g. "sum", "average", "countNums",
    // "min", "max", "stdDev", "var". Empty means no aggregation.
    const std::string& totalsRowFunction() const noexcept { return totalsRowFunction_; }
    void setTotalsRowFunction(std::string value) {
        static const std::vector<std::string> supported{
            "sum", "average", "count", "countNums", "min", "max", "stdDev",
            "stdDevp", "var", "varp"
        };
        if (!value.empty() && std::find(supported.begin(), supported.end(), value) == supported.end())
            throw std::invalid_argument("Unsupported table totals-row function: " + value);
        totalsRowFunction_ = std::move(value);
    }
    void clearTotalsRowFunction() noexcept { totalsRowFunction_.clear(); }
    // Optional label shown in the totals cell instead of the computed value.
    const std::string& totalsRowLabel() const noexcept { return totalsRowLabel_; }
    void setTotalsRowLabel(std::string value) { totalsRowLabel_ = std::move(value); }
    void clearTotalsRowLabel() noexcept { totalsRowLabel_.clear(); }
    // Calculated totals-row formula (x14:calculatedColumnFormula); kept raw.
    const std::string& totalsRowFormula() const noexcept { return totalsRowFormula_; }
    void setTotalsRowFormula(std::string value) { totalsRowFormula_ = std::move(value); }
    void clearTotalsRowFormula() noexcept { totalsRowFormula_.clear(); }
private:
    std::size_t id_ = 0;
    std::string name_;
    std::string totalsRowFunction_;
    std::string totalsRowLabel_;
    std::string totalsRowFormula_;
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
    std::vector<TableColumn>& columns() noexcept { return columns_; }
    const std::vector<TableColumn>& columns() const noexcept { return columns_; }
    TableColumn& addColumn(std::string name) { columns_.emplace_back(columns_.size()+1, std::move(name)); return columns_.back(); }
    TableStyleInfo& styleInfo() noexcept { return styleInfo_; }
    const TableStyleInfo& styleInfo() const noexcept { return styleInfo_; }
private:
    std::string name_;
    std::string displayName_;
    std::string reference_;
    bool showHeaderRow_ = true;
    bool showTotalsRow_ = false;
    std::vector<TableColumn> columns_;
    TableStyleInfo styleInfo_;
};
}
