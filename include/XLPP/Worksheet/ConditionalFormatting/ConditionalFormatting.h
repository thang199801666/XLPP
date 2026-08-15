#pragma once
#include <XLPP/Styles/Style.h>
#include <cstddef>
#include <optional>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {

enum class ConditionalRuleType {
    Formula,
    CellIs,
    DataBar,
    ColorScale,
    IconSet,
    ContainsText,
    NotContainsText,
    BeginsWith,
    EndsWith,
    AboveAverage,
    BelowAverage,
    AboveOrEqualAverage,
    BelowOrEqualAverage,
    Top10,
    Bottom10,
    DuplicateValues,
    UniqueValues,
    TimePeriod
};

enum class ConditionalOperator {
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
    Between,
    NotBetween
};

// Conditional format value object (cfvo): the min/max/midpoint of a data bar,
// color scale, or icon set threshold.
struct Cfvo {
    // "num", "percent", "percentile", "min", "max", "formula"
    std::string type{"num"};
    double value{0};
    bool hasValue{false};
    std::string formula;
    std::optional<std::string> color; // ARGB for color scale stops

    Cfvo() = default;
    explicit Cfvo(std::string type_, double val) : type(std::move(type_)), value(val), hasValue(true) {}
    Cfvo(std::string type_, std::string formula_) : type(std::move(type_)), hasValue(true), formula(std::move(formula_)) {}
};

// Data bar conditional formatting (Excel 2010+).
class DataBar {
public:
    std::string color{"FF638EC6"};
    Cfvo min;
    Cfvo max;
    bool showValue{true};
    // direction: "leftToRight", "rightToLeft", "context"
    std::string direction{"leftToRight"};
    std::optional<double> axisPosition;
};

// Color scale conditional formatting (2-3 stops).
class ColorScale {
public:
    std::vector<Cfvo> stops; // each stop has type/value + color
    void addStop(Cfvo stop) { stops.push_back(std::move(stop)); }
};

// Icon set conditional formatting (Excel 2010+).
enum class IconSetStyle {
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights,
    ThreeSigns,
    ThreeSymbols,
    ThreeStars,
    ThreeTriangles,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
    FiveBoxes
};

class IconSet {
public:
    std::string icons{"3Arrows"};
    std::vector<Cfvo> thresholds;
    bool reverse{false};
    bool showValue{true};
    std::optional<IconSetStyle> style;
    void addThreshold(Cfvo stop) { thresholds.push_back(std::move(stop)); }
};

class ConditionalRule {
public:
    static ConditionalRule formula(std::string expression) {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::Formula;
        rule.formulas_.push_back(std::move(expression));
        return rule;
    }

    static ConditionalRule cellIs(ConditionalOperator op, std::string value) {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::CellIs;
        rule.operator_ = op;
        rule.formulas_.push_back(std::move(value));
        return rule;
    }

    static ConditionalRule cellIsBetween(std::string lower, std::string upper, bool negate = false) {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::CellIs;
        rule.operator_ = negate ? ConditionalOperator::NotBetween : ConditionalOperator::Between;
        rule.formulas_.push_back(std::move(lower));
        rule.formulas_.push_back(std::move(upper));
        return rule;
    }

    static ConditionalRule dataBar(std::string color = "FF638EC6") {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::DataBar;
        rule.dataBar_.color = std::move(color);
        return rule;
    }

    static ConditionalRule colorScale() {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::ColorScale;
        return rule;
    }

    static ConditionalRule iconSet(std::string icons = "3Arrows") {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::IconSet;
        rule.iconSet_.icons = std::move(icons);
        return rule;
    }

    static ConditionalRule containsText(std::string text) {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::ContainsText;
        rule.text_ = std::move(text);
        return rule;
    }
    static ConditionalRule notContainsText(std::string text) {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::NotContainsText;
        rule.text_ = std::move(text);
        return rule;
    }
    static ConditionalRule beginsWith(std::string text) {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::BeginsWith;
        rule.text_ = std::move(text);
        return rule;
    }
    static ConditionalRule endsWith(std::string text) {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::EndsWith;
        rule.text_ = std::move(text);
        return rule;
    }
    static ConditionalRule aboveAverage() {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::AboveAverage;
        return rule;
    }
    static ConditionalRule belowAverage() {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::BelowAverage;
        return rule;
    }
    static ConditionalRule top10(bool top, int value, bool percent = false) {
        ConditionalRule rule;
        rule.type_ = top ? ConditionalRuleType::Top10 : ConditionalRuleType::Bottom10;
        rule.top10_.rank = value;
        rule.top10_.percent = percent;
        return rule;
    }
    static ConditionalRule duplicateValues() {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::DuplicateValues;
        return rule;
    }
    static ConditionalRule uniqueValues() {
        ConditionalRule rule;
        rule.type_ = ConditionalRuleType::UniqueValues;
        return rule;
    }

    ConditionalRuleType type() const noexcept { return type_; }
    ConditionalOperator op() const noexcept { return operator_; }
    void setOperator(ConditionalOperator value) noexcept { operator_ = value; }

    const std::vector<std::string>& formulas() const noexcept { return formulas_; }
    void setFormulas(std::vector<std::string> values) { formulas_ = std::move(values); }
    void addFormula(std::string value) { formulas_.push_back(std::move(value)); }

    std::size_t priority() const noexcept { return priority_; }
    void setPriority(std::size_t value) noexcept { priority_ = value; }

    bool stopIfTrue() const noexcept { return stopIfTrue_; }
    void setStopIfTrue(bool value) noexcept { stopIfTrue_ = value; }

    bool hasDifferentialStyle() const noexcept { return hasDifferentialStyle_; }
    Style& differentialStyle() noexcept { hasDifferentialStyle_ = true; return differentialStyle_; }
    const Style& differentialStyle() const noexcept { return differentialStyle_; }
    void setDifferentialStyle(Style value) { differentialStyle_ = std::move(value); hasDifferentialStyle_ = true; }
    void clearDifferentialStyle() noexcept { differentialStyle_ = Style{}; hasDifferentialStyle_ = false; }

    // DataBar access
    DataBar& getDataBar() noexcept { return dataBar_; }
    const DataBar& getDataBar() const noexcept { return dataBar_; }

    // ColorScale access
    ColorScale& getColorScale() noexcept { return colorScale_; }
    const ColorScale& getColorScale() const noexcept { return colorScale_; }

    // IconSet access
    IconSet& getIconSet() noexcept { return iconSet_; }
    const IconSet& getIconSet() const noexcept { return iconSet_; }

    // Text-rule access (containsText/notContainsText/beginsWith/endsWith).
    const std::string& text() const noexcept { return text_; }
    void setText(std::string value) { text_ = std::move(value); }

    // Above/below average flag (equal average when false).
    bool equalAverage() const noexcept { return equalAverage_; }
    void setEqualAverage(bool value) noexcept { equalAverage_ = value; }
    bool stdDev() const noexcept { return stdDev_; }
    void setStdDev(bool value) noexcept { stdDev_ = value; }
    int stdDevCount() const noexcept { return stdDevCount_; }
    void setStdDevCount(int value) noexcept { stdDevCount_ = value; }

    // Top10 access.
    struct Top10Config { int rank{10}; bool percent{false}; };
    const Top10Config& top10Config() const noexcept { return top10_; }
    Top10Config& top10Config() noexcept { return top10_; }
    void setTop10(int rank, bool percent = false) { top10_.rank = rank; top10_.percent = percent; }

private:
    ConditionalRuleType type_{ConditionalRuleType::Formula};
    ConditionalOperator operator_{ConditionalOperator::Equal};
    std::vector<std::string> formulas_;
    std::size_t priority_{0};
    bool stopIfTrue_{false};
    Style differentialStyle_;
    bool hasDifferentialStyle_{false};
    DataBar dataBar_;
    ColorScale colorScale_;
    IconSet iconSet_;
    std::string text_;
    bool equalAverage_{false};
    bool stdDev_{false};
    int stdDevCount_{2};
    Top10Config top10_{};
};

class ConditionalFormattingEntry {
public:
    explicit ConditionalFormattingEntry(std::string reference) : reference_(std::move(reference)) {
        if (reference_.empty()) throw std::invalid_argument("Conditional formatting reference cannot be empty");
    }

    const std::string& reference() const noexcept { return reference_; }
    void setReference(std::string value) {
        if (value.empty()) throw std::invalid_argument("Conditional formatting reference cannot be empty");
        reference_ = std::move(value);
    }

    ConditionalRule& addRule(ConditionalRule rule) {
        rules_.push_back(std::move(rule));
        return rules_.back();
    }

    const std::vector<ConditionalRule>& rules() const noexcept { return rules_; }
    std::vector<ConditionalRule>& rules() noexcept { return rules_; }
    bool empty() const noexcept { return rules_.empty(); }

private:
    std::string reference_;
    std::vector<ConditionalRule> rules_;
};

class ConditionalFormattingCollection {
public:
    ConditionalFormattingEntry& add(std::string reference) {
        entries_.emplace_back(std::move(reference));
        return entries_.back();
    }

    ConditionalRule& addRule(std::string reference, ConditionalRule rule) {
        for (auto& entry : entries_) {
            if (entry.reference() == reference) return entry.addRule(std::move(rule));
        }
        return add(std::move(reference)).addRule(std::move(rule));
    }

    const std::vector<ConditionalFormattingEntry>& entries() const noexcept { return entries_; }
    std::vector<ConditionalFormattingEntry>& entries() noexcept { return entries_; }
    bool empty() const noexcept { return entries_.empty(); }
    void clear() noexcept { entries_.clear(); }

private:
    std::vector<ConditionalFormattingEntry> entries_;
};

} // namespace xlpp
