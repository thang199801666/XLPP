#pragma once
#include <optional>
#include <string>
#include <utility>

namespace xlpp {

enum class FormulaType {
    Normal,
    Shared,
    Array,
    DataTable,
    // Excel 365 dynamic array formula (spill range). Serialized as t="array"
    // with ref + aca="1"; the formula uses the _xlfn. prefix for functions
    // like SORT, FILTER, UNIQUE, SEQUENCE, XLOOKUP.
    DynamicArray
};

class FormulaMetadata {
public:
    FormulaType type() const noexcept { return type_; }
    void setType(FormulaType value) noexcept { type_ = value; }

    const std::string& reference() const noexcept { return reference_; }
    void setReference(std::string value) { reference_ = std::move(value); }
    void clearReference() noexcept { reference_.clear(); }

    const std::optional<unsigned>& sharedIndex() const noexcept { return sharedIndex_; }
    void setSharedIndex(unsigned value) noexcept { sharedIndex_ = value; }
    void clearSharedIndex() noexcept { sharedIndex_.reset(); }

    bool alwaysCalculateArray() const noexcept { return alwaysCalculateArray_; }
    void setAlwaysCalculateArray(bool value) noexcept { alwaysCalculateArray_ = value; }

    bool calculateOnLoad() const noexcept { return calculateOnLoad_; }
    void setCalculateOnLoad(bool value) noexcept { calculateOnLoad_ = value; }

    bool empty() const noexcept {
        return type_ == FormulaType::Normal && reference_.empty() &&
               !sharedIndex_.has_value() && !alwaysCalculateArray_ &&
               !calculateOnLoad_;
    }

private:
    FormulaType type_{FormulaType::Normal};
    std::string reference_;
    std::optional<unsigned> sharedIndex_;
    bool alwaysCalculateArray_{false};
    bool calculateOnLoad_{false};
};

enum class CellError {
    Null,
    DivisionByZero,
    Value,
    Reference,
    Name,
    Number,
    NotAvailable,
    GettingData
};

inline std::string toString(CellError error) {
    switch (error) {
    case CellError::Null: return "#NULL!";
    case CellError::DivisionByZero: return "#DIV/0!";
    case CellError::Value: return "#VALUE!";
    case CellError::Reference: return "#REF!";
    case CellError::Name: return "#NAME?";
    case CellError::Number: return "#NUM!";
    case CellError::NotAvailable: return "#N/A";
    case CellError::GettingData: return "#GETTING_DATA";
    }
    return "#VALUE!";
}

inline CellError cellErrorFromString(const std::string& value) {
    if (value == "#NULL!") return CellError::Null;
    if (value == "#DIV/0!") return CellError::DivisionByZero;
    if (value == "#REF!") return CellError::Reference;
    if (value == "#NAME?") return CellError::Name;
    if (value == "#NUM!") return CellError::Number;
    if (value == "#N/A") return CellError::NotAvailable;
    if (value == "#GETTING_DATA") return CellError::GettingData;
    return CellError::Value;
}

}
