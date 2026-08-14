#pragma once
#include "Color.h"
#include <XLPP/Detail/CompactString.h>
#include <string>
#include <cstddef>

namespace xlpp {
class BorderSide {
public:
    const std::string& style() const noexcept { return style_.get(defaultStyle()); }
    void setStyle(std::string value) { style_.set(std::move(value), defaultStyle()); }
    Color& color() noexcept { return color_; }
    const Color& color() const noexcept { return color_; }
    bool isDefault() const noexcept { return style_.isDefault() && color_.isDefault(); }
    std::size_t hash() const noexcept {
        std::size_t h = 0;
        auto combine = [&h](auto&& v) { h ^= std::hash<std::decay_t<decltype(v)>>{}(v) + 0x9e3779b9 + (h << 6) + (h >> 2); };
        combine(style());
        combine(color_.hash());
        return h;
    }
    friend bool operator==(const BorderSide&, const BorderSide&) = default;
private:
    static const std::string& defaultStyle() noexcept {
        static const std::string value;
        return value;
    }
    internal::CompactString style_;
    Color color_;
};

class Border {
public:
    BorderSide& left() noexcept { return left_; }
    const BorderSide& left() const noexcept { return left_; }
    BorderSide& right() noexcept { return right_; }
    const BorderSide& right() const noexcept { return right_; }
    BorderSide& top() noexcept { return top_; }
    const BorderSide& top() const noexcept { return top_; }
    BorderSide& bottom() noexcept { return bottom_; }
    const BorderSide& bottom() const noexcept { return bottom_; }
    BorderSide& diagonal() noexcept { return diagonal_; }
    const BorderSide& diagonal() const noexcept { return diagonal_; }
    bool isDefault() const noexcept {
        return left_.isDefault() && right_.isDefault() && top_.isDefault() && bottom_.isDefault() && diagonal_.isDefault();
    }
    std::size_t hash() const noexcept {
        std::size_t h = 0;
        auto combine = [&h](auto&& v) { h ^= v + 0x9e3779b9 + (h << 6) + (h >> 2); };
        combine(left_.hash());
        combine(right_.hash());
        combine(top_.hash());
        combine(bottom_.hash());
        combine(diagonal_.hash());
        return h;
    }
    friend bool operator==(const Border&, const Border&) = default;
private:
    BorderSide left_, right_, top_, bottom_, diagonal_;
};
}
