#pragma once
#include "Style.h"
#include <string>
#include <utility>

namespace xlpp {
class NamedStyle {
public:
    NamedStyle() = default;
    explicit NamedStyle(std::string name) : name_(std::move(name)) {}
    NamedStyle(std::string name, Style style) : name_(std::move(name)), style_(std::move(style)) {}

    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
    Style& style() noexcept { return style_; }
    const Style& style() const noexcept { return style_; }
private:
    std::string name_;
    Style style_;
};
}
