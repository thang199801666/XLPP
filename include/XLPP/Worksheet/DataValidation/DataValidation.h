#pragma once
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {

enum class DataValidationType { None, Whole, Decimal, List, Date, Time, TextLength, Custom };
enum class DataValidationOperator { Between, NotBetween, Equal, NotEqual, LessThan, LessThanOrEqual, GreaterThan, GreaterThanOrEqual };
enum class DataValidationErrorStyle { Stop, Warning, Information };

class DataValidation {
public:
    explicit DataValidation(DataValidationType type = DataValidationType::None) : type_(type) {}

    DataValidationType type() const noexcept { return type_; }
    void setType(DataValidationType value) noexcept { type_ = value; }
    DataValidationOperator op() const noexcept { return operator_; }
    void setOperator(DataValidationOperator value) noexcept { operator_ = value; }
    DataValidationErrorStyle errorStyle() const noexcept { return errorStyle_; }
    void setErrorStyle(DataValidationErrorStyle value) noexcept { errorStyle_ = value; }

    const std::string& formula1() const noexcept { return formula1_; }
    void setFormula1(std::string value) { formula1_ = std::move(value); }
    const std::string& formula2() const noexcept { return formula2_; }
    void setFormula2(std::string value) { formula2_ = std::move(value); }

    const std::string& reference() const noexcept { return reference_; }
    void setReference(std::string value) {
        if (value.empty()) throw std::invalid_argument("Data validation reference cannot be empty");
        reference_ = std::move(value);
    }

    bool allowBlank() const noexcept { return allowBlank_; }
    void setAllowBlank(bool value) noexcept { allowBlank_ = value; }
    bool showDropDown() const noexcept { return showDropDown_; }
    void setShowDropDown(bool value) noexcept { showDropDown_ = value; }
    bool showInputMessage() const noexcept { return showInputMessage_; }
    void setShowInputMessage(bool value) noexcept { showInputMessage_ = value; }
    bool showErrorMessage() const noexcept { return showErrorMessage_; }
    void setShowErrorMessage(bool value) noexcept { showErrorMessage_ = value; }

    const std::string& promptTitle() const noexcept { return promptTitle_; }
    void setPromptTitle(std::string value) { promptTitle_ = std::move(value); }
    const std::string& prompt() const noexcept { return prompt_; }
    void setPrompt(std::string value) { prompt_ = std::move(value); }
    const std::string& errorTitle() const noexcept { return errorTitle_; }
    void setErrorTitle(std::string value) { errorTitle_ = std::move(value); }
    const std::string& error() const noexcept { return error_; }
    void setError(std::string value) { error_ = std::move(value); }

    static DataValidation list(std::string reference, std::string formula) {
        DataValidation value(DataValidationType::List);
        value.setReference(std::move(reference));
        value.setFormula1(std::move(formula));
        return value;
    }

private:
    DataValidationType type_{DataValidationType::None};
    DataValidationOperator operator_{DataValidationOperator::Between};
    DataValidationErrorStyle errorStyle_{DataValidationErrorStyle::Stop};
    std::string formula1_, formula2_, reference_;
    bool allowBlank_{false}, showDropDown_{false}, showInputMessage_{false}, showErrorMessage_{false};
    std::string promptTitle_, prompt_, errorTitle_, error_;
};

class DataValidationCollection {
public:
    DataValidation& add(DataValidation validation) {
        if (validation.reference().empty()) throw std::invalid_argument("Data validation reference cannot be empty");
        items_.push_back(std::move(validation));
        return items_.back();
    }
    DataValidation& add(DataValidationType type, std::string reference) {
        DataValidation validation(type);
        validation.setReference(std::move(reference));
        return add(std::move(validation));
    }
    const std::vector<DataValidation>& items() const noexcept { return items_; }
    std::vector<DataValidation>& items() noexcept { return items_; }
    bool empty() const noexcept { return items_.empty(); }
    void clear() noexcept { items_.clear(); }
private:
    std::vector<DataValidation> items_;
};

}
