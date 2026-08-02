#pragma once
#include <optional>
#include <string>
#include <stdexcept>

namespace xlpp {
class DefinedName {
public:
    DefinedName() = default;
    DefinedName(std::string name, std::string value) : name_(std::move(name)), value_(std::move(value)) {
        if (name_.empty() || value_.empty()) throw std::invalid_argument("Defined name and value cannot be empty");
    }
    const std::string& name() const noexcept { return name_; }
    const std::string& value() const noexcept { return value_; }
    void setValue(std::string value) { if (value.empty()) throw std::invalid_argument("Defined name value cannot be empty"); value_ = std::move(value); }
    const std::optional<std::size_t>& localSheetId() const noexcept { return localSheetId_; }
    void setLocalSheetId(std::size_t value) noexcept { localSheetId_ = value; }
    void clearLocalSheetId() noexcept { localSheetId_.reset(); }
    bool hidden() const noexcept { return hidden_; }
    void setHidden(bool value) noexcept { hidden_ = value; }
    const std::string& comment() const noexcept { return comment_; }
    void setComment(std::string value) { comment_ = std::move(value); }
private:
    std::string name_;
    std::string value_;
    std::optional<std::size_t> localSheetId_;
    bool hidden_ = false;
    std::string comment_;
};
}
