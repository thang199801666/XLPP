#pragma once
#include <XLPP/Detail/CompactString.h>
#include <string>
#include <cstddef>
#include <functional>

namespace xlpp {
class Alignment {
public:
    const std::string& horizontal() const noexcept { return horizontal_.get(defaultText()); }
    void setHorizontal(std::string value) { horizontal_.set(std::move(value), defaultText()); }
    const std::string& vertical() const noexcept { return vertical_.get(defaultText()); }
    void setVertical(std::string value) { vertical_.set(std::move(value), defaultText()); }
    bool wrapText() const noexcept { return wrapText_; }
    void setWrapText(bool value) noexcept { wrapText_ = value; }
    bool shrinkToFit() const noexcept { return shrinkToFit_; }
    void setShrinkToFit(bool value) noexcept { shrinkToFit_ = value; }
    int textRotation() const noexcept { return textRotation_; }
    void setTextRotation(int value) { textRotation_ = value; }
    int indent() const noexcept { return indent_; }
    void setIndent(int value) { indent_ = value; }
    bool isDefault() const noexcept {
        return horizontal_.isDefault() && vertical_.isDefault() && !wrapText_ && !shrinkToFit_ && textRotation_ == 0 && indent_ == 0;
    }

    std::size_t hash() const noexcept {
        std::size_t h = 0;
        auto combine = [&h](auto&& v) { h ^= std::hash<std::decay_t<decltype(v)>>{}(v) + 0x9e3779b9 + (h << 6) + (h >> 2); };
        combine(horizontal());
        combine(vertical());
        combine(wrapText_);
        combine(shrinkToFit_);
        combine(textRotation_);
        combine(indent_);
        return h;
    }
    friend bool operator==(const Alignment&, const Alignment&) = default;
private:
    static const std::string& defaultText() noexcept {
        static const std::string value;
        return value;
    }
    internal::CompactString horizontal_;
    internal::CompactString vertical_;
    bool wrapText_{false};
    bool shrinkToFit_{false};
    int textRotation_{0};
    int indent_{0};
};
}
