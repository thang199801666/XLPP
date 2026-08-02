#pragma once
#include "CellReference.h"
#include "DateTime.h"
#include <XLPP/Styles/Style.h>
#include "Hyperlink.h"
#include "Comment.h"
#include "Formula.h"
#include <optional>
#include <cstdint>
#include <string>
#include <string_view>
#include <variant>

namespace xlpp {
using CellValue = std::variant<std::monostate, double, bool, std::string, CellError, DateTime>;

class Cell {
public:
    Cell() = default;
    explicit Cell(std::string address) { setAddress(std::move(address)); }
    Cell(std::size_t row, std::size_t column) { setPosition(row, column); }

    const std::string& address() const noexcept { return address_; }
    std::size_t row() const noexcept { return row_; }
    std::size_t column() const noexcept { return column_; }

    const CellValue& value() const noexcept { return value_; }
    void setValue(CellValue value) { value_ = std::move(value); }
    void setValue(std::monostate value) { value_ = value; }
    void setValue(double value) { value_ = value; }
    void setValue(bool value) { value_ = value; }
    void setValue(const std::string& value) { value_ = value; }
    void setValue(std::string&& value) { value_ = std::move(value); }
    void setValue(const char* value) { value_ = std::string(value ? value : ""); }
    void setValue(std::string_view value) { value_ = std::string(value); }
    void setValue(std::int64_t value) { value_ = static_cast<double>(value); }
    void setValue(int value) { value_ = static_cast<double>(value); }
    void setValue(DateTime value) { value_ = value; }
    void setError(CellError value) { value_ = value; }
    void setStringValue(const std::string& value) { value_ = value; }
    void setStringValue(std::string&& value) { value_ = std::move(value); }
    void setStringValue(const char* value) { value_ = std::string(value ? value : ""); }
    void setNumericValue(double value) { value_ = value; }
    void setBoolValue(bool value) { value_ = value; }
    double numericValueOr(double fallback) const noexcept {
        if (const auto* v = std::get_if<double>(&value_)) return *v;
        if (const auto* v = std::get_if<DateTime>(&value_)) return xlpp::toExcelSerial(*v, false);
        return fallback;
    }
    std::string stringValueOr(std::string fallback) const noexcept {
        if (const auto* v = std::get_if<std::string>(&value_)) return *v;
        return std::move(fallback);
    }
    bool isError() const noexcept { return std::holds_alternative<CellError>(value_); }
    std::optional<CellError> error() const noexcept { if (const auto* value = std::get_if<CellError>(&value_)) return *value; return std::nullopt; }

    std::optional<DateTime> date() const noexcept {
        if (const auto* value = std::get_if<DateTime>(&value_)) return *value;
        return std::nullopt;
    }
    // Sets a date/time value and applies a matching default number format when
    // the cell currently uses the "General" format.
    void setDate(const DateTime& value) {
        value_ = value;
        if (style_.numberFormat() == "General") style_.setNumberFormat("yyyy-mm-dd");
    }
    void setDate(int year, int month, int day) { setDate(DateTime{year, month, day}); }
    void setDateTime(const DateTime& value) {
        value_ = value;
        if (style_.numberFormat() == "General") style_.setNumberFormat("yyyy-mm-dd h:mm:ss");
    }

    bool empty() const noexcept { return std::holds_alternative<std::monostate>(value_) && formula_.empty(); }
    void clear() noexcept { value_ = std::monostate{}; formula_.clear(); formulaMetadata_ = FormulaMetadata{}; namedStyle_.reset(); }

    bool hasValue() const noexcept { return !std::holds_alternative<std::monostate>(value_); }
    bool isNumeric() const noexcept { return std::holds_alternative<double>(value_); }
    bool isString() const noexcept { return std::holds_alternative<std::string>(value_); }
    bool isBoolean() const noexcept { return std::holds_alternative<bool>(value_); }
    bool isDate() const noexcept { return std::holds_alternative<DateTime>(value_); }
    const char* valueType() const noexcept {
        if (std::holds_alternative<std::monostate>(value_)) return "empty";
        if (std::holds_alternative<double>(value_)) return "numeric";
        if (std::holds_alternative<std::string>(value_)) return "string";
        if (std::holds_alternative<bool>(value_)) return "bool";
        if (std::holds_alternative<CellError>(value_)) return "error";
        if (std::holds_alternative<DateTime>(value_)) return "date";
        return "unknown";
    }

    Style& style() noexcept { return style_; }
    const Style& style() const noexcept { return style_; }
    Font& font() noexcept { return style_.font(); }
    const Font& font() const noexcept { return style_.font(); }
    Fill& fill() noexcept { return style_.fill(); }
    const Fill& fill() const noexcept { return style_.fill(); }
    Border& border() noexcept { return style_.border(); }
    const Border& border() const noexcept { return style_.border(); }
    Alignment& alignment() noexcept { return style_.alignment(); }
    const Alignment& alignment() const noexcept { return style_.alignment(); }
    const std::string& numberFormat() const noexcept { return style_.numberFormat(); }
    void setNumberFormat(std::string value) { style_.setNumberFormat(std::move(value)); }
    std::optional<std::size_t> styleIndex() const noexcept { return styleIndex_; }
    void setRawStyleIndex(std::size_t value) noexcept { styleIndex_ = value; }
    void clearRawStyleIndex() noexcept { styleIndex_.reset(); }

    bool hasFormula() const noexcept { return !formula_.empty(); }
    const std::string& formula() const noexcept { return formula_; }
    void setFormula(std::string formula) { formula_ = std::move(formula); }
    void setSharedFormula(std::string formula, unsigned sharedIndex, std::string reference = {}) { formula_ = std::move(formula); formulaMetadata_.setType(FormulaType::Shared); formulaMetadata_.setSharedIndex(sharedIndex); formulaMetadata_.setReference(std::move(reference)); }
    void setArrayFormula(std::string formula, std::string reference) { formula_ = std::move(formula); formulaMetadata_.setType(FormulaType::Array); formulaMetadata_.setReference(std::move(reference)); }
    FormulaMetadata& formulaMetadata() noexcept { return formulaMetadata_; }
    const FormulaMetadata& formulaMetadata() const noexcept { return formulaMetadata_; }
    void clearFormula() noexcept { formula_.clear(); formulaMetadata_ = FormulaMetadata{}; }

    bool hasHyperlink() const noexcept { return hyperlink_.has_value(); }
    Hyperlink& hyperlink() { if(!hyperlink_) hyperlink_.emplace(); return *hyperlink_; }
    const std::optional<Hyperlink>& hyperlinkValue() const noexcept { return hyperlink_; }
    void setHyperlink(Hyperlink value) { hyperlink_=std::move(value); }
    void clearHyperlink() noexcept { hyperlink_.reset(); }
    bool hasComment() const noexcept { return comment_.has_value(); }
    Comment& comment() { if(!comment_) comment_.emplace(); return *comment_; }
    const std::optional<Comment>& commentValue() const noexcept { return comment_; }
    void setComment(Comment value) { comment_=std::move(value); }
    void clearComment() noexcept { comment_.reset(); }

    const std::optional<std::string>& namedStyle() const noexcept { return namedStyle_; }
    void setNamedStyle(std::optional<std::string> name) noexcept { namedStyle_ = std::move(name); }

    CellReference offset(int rows, int cols) const {
        if (rows < -static_cast<int>(row_) + 1 || cols < -static_cast<int>(column_) + 1)
            throw std::invalid_argument("Offset would produce negative coordinate");
        return CellReference{
            static_cast<std::size_t>(static_cast<int>(row_) + rows),
            static_cast<std::size_t>(static_cast<int>(column_) + cols)
        };
    }

private:
    friend class Worksheet;
    void setAddress(std::string address) {
        const auto ref = CellReference::parse(address);
        setPosition(ref.row, ref.column);
    }
    void setPosition(std::size_t row, std::size_t column) {
        if (row == 0 || column == 0) throw std::invalid_argument("Cell coordinates are 1-based");
        row_ = row;
        column_ = column;
        address_ = CellReference{row, column}.address();
    }

    std::string address_;
    std::size_t row_{1};
    std::size_t column_{1};
    CellValue value_;
    std::string formula_;
    FormulaMetadata formulaMetadata_;
    Style style_;
    std::optional<Hyperlink> hyperlink_;
    std::optional<Comment> comment_;
    std::optional<std::string> namedStyle_;
    std::optional<std::size_t> styleIndex_;
};
}
