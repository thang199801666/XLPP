#pragma once
#include <string>
#include <vector>

namespace xlpp {

class CustomProperty {
public:
    CustomProperty() = default;
    CustomProperty(std::string name, std::string value) : name_(std::move(name)), value_(std::move(value)), type_("lpwstr") {}
    CustomProperty(std::string name, int value) : name_(std::move(name)), value_(std::to_string(value)), type_("i4") {}
    CustomProperty(std::string name, double value) : name_(std::move(name)), value_(std::to_string(value)), type_("r8") {}
    CustomProperty(std::string name, bool value) : name_(std::move(name)), value_(value ? "true" : "false"), type_("bool") {}

    const std::string& name() const noexcept { return name_; } void setName(std::string v) { name_ = std::move(v); }
    const std::string& value() const noexcept { return value_; } void setValue(std::string v) { value_ = std::move(v); }
    const std::string& type() const noexcept { return type_; } void setType(std::string v) { type_ = std::move(v); }
private:
    std::string name_, value_, type_;
};

class CustomProperties {
public:
    void add(CustomProperty prop) { properties_.push_back(std::move(prop)); }
    const std::vector<CustomProperty>& items() const noexcept { return properties_; }
    std::vector<CustomProperty>& items() noexcept { return properties_; }
    bool empty() const noexcept { return properties_.empty(); }
private:
    std::vector<CustomProperty> properties_;
};

}
