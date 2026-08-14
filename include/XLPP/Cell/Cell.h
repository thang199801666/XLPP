#pragma once
#include "CellReference.h"
#include "DateTime.h"
#include <XLPP/Styles/Style.h>
#include <XLPP/Detail/CompactString.h>
#include <XLPP/Detail/CompactOptional.h>
#include <XLPP/Detail/CompactValue.h>
#include "Hyperlink.h"
#include "Comment.h"
#include "Formula.h"
#include "RichText.h"
#include <optional>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <variant>

namespace xlpp {
using CellValue = std::variant<std::monostate, double, bool, std::string, CellError, DateTime>;

// Excel 365 dynamic array / new functions that require the _xlfn. prefix in
// the stored formula so Excel 2016-2019 & 365 parse them correctly. Returns the
// prefixed form, or the input unchanged if it already has a prefix.
inline std::string xlfn(std::string_view function) {
    if (function.empty()) return {};
    if (function.size() > 6 && function.substr(0, 6) == "_xlfn.") return std::string(function);
    return "_xlfn." + std::string(function);
}

class Cell {
public:
    Cell() = default;
    explicit Cell(std::string address) { setAddress(std::move(address)); }
    Cell(std::size_t row, std::size_t column) { setPosition(row, column); }

    const std::string& address() const noexcept { return address_; }
    std::size_t row() const noexcept { return static_cast<std::size_t>(coordinate_ >> 20); }
    std::size_t column() const noexcept { return static_cast<std::size_t>(coordinate_ & CellKeyColumnMask); }
    std::uint64_t mutationRevision() const noexcept { return mutationRevision_; }
    void clearMutationRevision() const noexcept { mutationRevision_ = 0; }

    const CellValue& value() const noexcept { return value_; }
    void setValue(CellValue value) { noteMutation(); richText_.reset(); value_ = std::move(value); }
    void setValue(std::monostate value) { noteMutation(); richText_.reset(); value_ = value; }
    void setValue(double value) { noteMutation(); richText_.reset(); value_ = value; }
    void setValue(bool value) { noteMutation(); richText_.reset(); value_ = value; }
    void setValue(const std::string& value) { noteMutation(); richText_.reset(); value_ = value; }
    void setValue(std::string&& value) { noteMutation(); richText_.reset(); value_ = std::move(value); }
    void setValue(const char* value) { noteMutation(); richText_.reset(); value_ = std::string(value ? value : ""); }
    void setValue(std::string_view value) { noteMutation(); richText_.reset(); value_ = std::string(value); }
    void setValue(std::int64_t value) { noteMutation(); richText_.reset(); value_ = static_cast<double>(value); }
    void setValue(int value) { noteMutation(); richText_.reset(); value_ = static_cast<double>(value); }
    void setValue(DateTime value) { noteMutation(); richText_.reset(); value_ = value; }
    void setError(CellError value) { noteMutation(); richText_.reset(); value_ = value; }
    void setStringValue(const std::string& value) { noteMutation(); richText_.reset(); value_ = value; }
    void setStringValue(std::string&& value) { noteMutation(); richText_.reset(); value_ = std::move(value); }
    void setStringValue(const char* value) { noteMutation(); richText_.reset(); value_ = std::string(value ? value : ""); }
    void setNumericValue(double value) { noteMutation(); richText_.reset(); value_ = value; }
    void setBoolValue(bool value) { noteMutation(); richText_.reset(); value_ = value; }

    bool hasRichText() const noexcept { return richText_.has_value(); }
    RichText& richText() {
        noteMutation();
        if (!richText_) richText_.emplace(RichText::fromPlain(stringValueOr({})));
        value_ = richText_->plainText();
        return *richText_;
    }
    const std::optional<RichText>& richTextValue() const noexcept { return richText_.optionalRef(); }
    void setRichText(RichText value) {
        noteMutation();
        value_ = value.plainText();
        richText_.set(std::move(value));
    }
    void clearRichText() noexcept { noteMutation(); richText_.reset(); }
    double numericValueOr(double fallback) const noexcept {
        if (const auto* v = std::get_if<double>(&value_)) return *v;
        if (const auto* v = std::get_if<DateTime>(&value_)) return xlpp::toExcelSerial(*v, false);
        return fallback;
    }
    std::string stringValueOr(std::string fallback) const noexcept {
        if (const auto* v = std::get_if<std::string>(&value_)) return *v;
        return fallback;
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
        noteMutation();
        richText_.reset();
        value_ = value;
        if (styleOrDefault().numberFormat() == "General") mutableStyle().setNumberFormat("yyyy-mm-dd");
    }
    void setDate(int year, int month, int day) { setDate(DateTime{year, month, day}); }
    void setDateTime(const DateTime& value) {
        noteMutation();
        richText_.reset();
        value_ = value;
        if (styleOrDefault().numberFormat() == "General") mutableStyle().setNumberFormat("yyyy-mm-dd h:mm:ss");
    }

    bool empty() const noexcept { return std::holds_alternative<std::monostate>(value_) && formula_.isDefault(); }
    void clear() noexcept { noteMutation(); value_ = std::monostate{}; richText_.reset(); formula_ = internal::CompactString{}; formulaMetadata_.reset(); namedStyle_.reset(); }

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

    // Mutable style access lazily materializes the Style object.  Most cells are
    // unstyled, so their const/read/save path shares a static default Style and
    // pays no per-cell Style storage. Mutable access is also a mutation-tracking
    // boundary, matching other reference-returning model APIs.
    Style& style() { noteMutation(); return mutableStyle(); }
    const Style& style() const noexcept { return styleOrDefault(); }
    bool hasNonDefaultStyle() const noexcept { return style_ && !style_->isDefault(); }
    Font& font() { return style().font(); }
    const Font& font() const noexcept { return styleOrDefault().font(); }
    Fill& fill() { return style().fill(); }
    const Fill& fill() const noexcept { return styleOrDefault().fill(); }
    Border& border() { return style().border(); }
    const Border& border() const noexcept { return styleOrDefault().border(); }
    Alignment& alignment() { return style().alignment(); }
    const Alignment& alignment() const noexcept { return styleOrDefault().alignment(); }
    const std::string& numberFormat() const noexcept { return styleOrDefault().numberFormat(); }
    void setNumberFormat(std::string value) { noteMutation(); mutableStyle().setNumberFormat(std::move(value)); }
    std::optional<std::size_t> styleIndex() const noexcept {
        return styleIndex_ == NoStyleIndex ? std::nullopt : std::optional<std::size_t>{styleIndex_};
    }
    void setRawStyleIndex(std::size_t value) noexcept { noteMutation(); styleIndex_ = value; }
    void clearRawStyleIndex() noexcept { noteMutation(); styleIndex_ = NoStyleIndex; }

    bool hasFormula() const noexcept { return !formula_.isDefault(); }
    const std::string& formula() const noexcept { return formula_.get(defaultFormula()); }
    void setFormula(std::string formula) { noteMutation(); formula_.set(std::move(formula), defaultFormula()); formulaMetadata_.reset(); }
    // Replaces only the formula expression while preserving shared/array/data-
    // table metadata. Structural-reference rewriting uses this path so moving
    // cells does not accidentally demote formula kind.
    void setFormulaTextPreservingMetadata(std::string formula) { noteMutation(); formula_.set(std::move(formula), defaultFormula()); }
    void setSharedFormula(std::string formula, unsigned sharedIndex, std::string reference = {}) {
        noteMutation();
        formula_.set(std::move(formula), defaultFormula());
        FormulaMetadata metadata;
        metadata.setType(FormulaType::Shared);
        metadata.setSharedIndex(sharedIndex);
        metadata.setReference(std::move(reference));
        formulaMetadata_.set(std::move(metadata));
    }
    void setArrayFormula(std::string formula, std::string reference) {
        noteMutation();
        formula_.set(std::move(formula), defaultFormula());
        FormulaMetadata metadata;
        metadata.setType(FormulaType::Array);
        metadata.setReference(std::move(reference));
        formulaMetadata_.set(std::move(metadata));
    }
    // Excel 365 dynamic array formula: t="array" + ref + aca="1" on save.
    // The formula string should use the _xlfn. prefix for new functions
    // (e.g. "_xlfn.SORT(A1:A10)") — see xlfn() helper.
    void setDynamicArrayFormula(std::string formula, std::string reference) {
        noteMutation();
        formula_.set(std::move(formula), defaultFormula());
        FormulaMetadata metadata;
        metadata.setType(FormulaType::DynamicArray);
        metadata.setReference(std::move(reference));
        metadata.setAlwaysCalculateArray(true);
        formulaMetadata_.set(std::move(metadata));
    }
    FormulaMetadata& formulaMetadata() { noteMutation(); return mutableFormulaMetadata(); }
    const FormulaMetadata& formulaMetadata() const noexcept { return formulaMetadataOrDefault(); }
    bool hasFormulaMetadata() const noexcept { return static_cast<bool>(formulaMetadata_); }
    void clearFormula() noexcept { noteMutation(); formula_ = internal::CompactString{}; formulaMetadata_.reset(); }

    bool hasHyperlink() const noexcept { return hyperlink_.has_value(); }
    Hyperlink& hyperlink() { noteMutation(); if(!hyperlink_) hyperlink_.emplace(); return *hyperlink_; }
    const std::optional<Hyperlink>& hyperlinkValue() const noexcept { return hyperlink_.optionalRef(); }
    void setHyperlink(Hyperlink value) { noteMutation(); hyperlink_.set(std::move(value)); }
    void clearHyperlink() noexcept { noteMutation(); hyperlink_.reset(); }
    bool hasComment() const noexcept { return comment_.has_value(); }
    Comment& comment() { noteMutation(); if(!comment_) comment_.emplace(); return *comment_; }
    const std::optional<Comment>& commentValue() const noexcept { return comment_.optionalRef(); }
    void setComment(Comment value) { noteMutation(); comment_.set(std::move(value)); }
    void clearComment() noexcept { noteMutation(); comment_.reset(); }

    const std::optional<std::string>& namedStyle() const noexcept { return namedStyle_.optionalRef(); }
    void setNamedStyle(std::optional<std::string> name) { noteMutation(); namedStyle_.set(std::move(name)); }

    CellReference offset(int rows, int cols) const {
        const auto currentRow = row();
        const auto currentColumn = column();
        if (rows < -static_cast<int>(currentRow) + 1 || cols < -static_cast<int>(currentColumn) + 1)
            throw std::invalid_argument("Offset would produce negative coordinate");
        return CellReference{
            static_cast<std::size_t>(static_cast<int>(currentRow) + rows),
            static_cast<std::size_t>(static_cast<int>(currentColumn) + cols)
        };
    }

private:
    friend class Worksheet;
    static constexpr std::uint64_t CellKeyColumnMask = (std::uint64_t{1} << 20) - 1;
    static constexpr std::size_t NoStyleIndex = std::numeric_limits<std::size_t>::max();
    static const std::string& defaultFormula() noexcept { static const std::string value; return value; }
    static const Style& defaultStyle() noexcept { static const Style value; return value; }
    static const FormulaMetadata& defaultFormulaMetadata() noexcept { static const FormulaMetadata value; return value; }
    Style& mutableStyle() { return style_.getOrCreate(); }
    const Style& styleOrDefault() const noexcept { return style_.getOr(defaultStyle()); }
    FormulaMetadata& mutableFormulaMetadata() { return formulaMetadata_.getOrCreate(); }
    const FormulaMetadata& formulaMetadataOrDefault() const noexcept { return formulaMetadata_.getOr(defaultFormulaMetadata()); }
    void noteMutation() const noexcept { ++mutationRevision_; }
    void setAddress(std::string address) {
        const auto ref = CellReference::parse(address);
        setPosition(ref.row, ref.column);
    }
    void setPosition(std::size_t row, std::size_t column) {
        if (!CellReference::validGridPosition(row, column))
            throw std::invalid_argument("Cell coordinates are outside the Excel grid");
        coordinate_ = makeCellKey(row, column);
        address_ = CellReference{row, column}.address();
    }

    std::string address_;
    std::uint64_t coordinate_{makeCellKey(1, 1)};
    mutable std::uint64_t mutationRevision_{0};
    CellValue value_;
    internal::CompactString formula_;
    internal::CompactValue<FormulaMetadata> formulaMetadata_;
    internal::CompactValue<Style> style_;
    internal::CompactOptional<Hyperlink> hyperlink_;
    internal::CompactOptional<Comment> comment_;
    internal::CompactOptional<RichText> richText_;
    internal::CompactOptional<std::string> namedStyle_;
    std::size_t styleIndex_{NoStyleIndex};
};
}
