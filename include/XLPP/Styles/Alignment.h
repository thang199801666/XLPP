#pragma once
#include <string>
#include <cstddef>
#include <functional>

namespace xlpp {
class Alignment {
public:
    const std::string& horizontal() const noexcept { return horizontal_; }
    void setHorizontal(std::string value) { horizontal_ = std::move(value); }
    const std::string& vertical() const noexcept { return vertical_; }
    void setVertical(std::string value) { vertical_ = std::move(value); }
    bool wrapText() const noexcept { return wrapText_; }
    void setWrapText(bool value) noexcept { wrapText_ = value; }
    bool shrinkToFit() const noexcept { return shrinkToFit_; }
    void setShrinkToFit(bool value) noexcept { shrinkToFit_ = value; }
    int textRotation() const noexcept { return textRotation_; }
    void setTextRotation(int value) { textRotation_ = value; }
    int indent() const noexcept { return indent_; }
    void setIndent(int value) { indent_ = value; }

    std::size_t hash() const noexcept {
        std::size_t h = 0;
        auto combine = [&h](auto&& v) { h ^= std::hash<std::decay_t<decltype(v)>>{}(v) + 0x9e3779b9 + (h << 6) + (h >> 2); };
        combine(horizontal_);
        combine(vertical_);
        combine(wrapText_);
        combine(shrinkToFit_);
        combine(textRotation_);
        combine(indent_);
        return h;
    }
    friend bool operator==(const Alignment&, const Alignment&) = default;
private:
    std::string horizontal_;
    std::string vertical_;
    bool wrapText_{false};
    bool shrinkToFit_{false};
    int textRotation_{0};
    int indent_{0};
};
}
