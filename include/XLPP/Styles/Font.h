#pragma once
#include "Color.h"
#include <XLPP/Detail/CompactString.h>
#include <string>
#include <cstddef>
#include <functional>

namespace xlpp {
class Font {
public:
    const std::string& name() const noexcept { return name_.get(defaultName()); }
    void setName(std::string value) { name_.set(std::move(value), defaultName()); }
    double size() const noexcept { return size_; }
    void setSize(double value) { size_ = value; }
    bool bold() const noexcept { return bold_; }
    void setBold(bool value) noexcept { bold_ = value; }
    bool italic() const noexcept { return italic_; }
    void setItalic(bool value) noexcept { italic_ = value; }
    bool underline() const noexcept { return underline_; }
    void setUnderline(bool value) noexcept { underline_ = value; }
    bool strike() const noexcept { return strike_; }
    void setStrike(bool value) noexcept { strike_ = value; }
    Color& color() noexcept { return color_; }
    const Color& color() const noexcept { return color_; }
    bool isDefault() const noexcept {
        return name_.isDefault() && size_ == 11.0 && !bold_ && !italic_ && !underline_ && !strike_ && color_.isDefault();
    }

    std::size_t hash() const noexcept {
        std::size_t h = 0;
        auto combine = [&h](auto&& v) { h ^= std::hash<std::decay_t<decltype(v)>>{}(v) + 0x9e3779b9 + (h << 6) + (h >> 2); };
        combine(name());
        combine(size_);
        combine(bold_);
        combine(italic_);
        combine(underline_);
        combine(strike_);
        combine(color_.hash());
        return h;
    }
    friend bool operator==(const Font&, const Font&) = default;
private:
    static const std::string& defaultName() noexcept {
        static const std::string value{"Calibri"};
        return value;
    }
    internal::CompactString name_;
    double size_{11.0};
    bool bold_{false};
    bool italic_{false};
    bool underline_{false};
    bool strike_{false};
    Color color_;
};
}
