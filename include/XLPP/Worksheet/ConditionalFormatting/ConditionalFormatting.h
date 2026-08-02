#pragma once
#include <XLPP/Styles/Style.h>
#include <cstddef>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {

enum class ConditionalRuleType {
    Formula,
    CellIs
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

private:
    ConditionalRuleType type_{ConditionalRuleType::Formula};
    ConditionalOperator operator_{ConditionalOperator::Equal};
    std::vector<std::string> formulas_;
    std::size_t priority_{0};
    bool stopIfTrue_{false};
    Style differentialStyle_;
    bool hasDifferentialStyle_{false};
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

}
