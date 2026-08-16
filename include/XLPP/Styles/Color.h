#pragma once
#include <XLPP/Detail/CompactString.h>
#include <string>
#include <utility>
#include <cstddef>
#include <functional>

namespace xlpp {
class Color {
public:
    Color() = default;
    explicit Color(std::string argb) { setArgb(std::move(argb)); }

    const std::string& argb() const noexcept { return argb_.get(defaultArgb()); }
    // Setting an ARGB color clears any theme reference.
    void setArgb(std::string value) {
        argb_.set(std::move(value), defaultArgb());
        theme_ = NoTheme;
        tint_ = 0.0f;
    }
    bool empty() const noexcept { return argb_.isDefault() && theme_ == NoTheme; }
    bool isDefault() const noexcept { return empty(); }

    // Theme color reference (ECMA-376 §18.8.27): the index into the workbook
    // theme palette (e.g. 0 = lt1, 1 = dk1, 2 = lt2, ...). A negative value
    // means the color is not theme-based. Setting a theme reference clears the
    // explicit ARGB value.
    bool hasTheme() const noexcept { return theme_ >= 0; }
    int theme() const noexcept { return theme_; }
    void setTheme(int index) {
        theme_ = index;
        tint_ = 0.0f;
        argb_.set("", defaultArgb());
    }
    // Tint/shade adjustment (-1.0 darkens, +1.0 lightens); only meaningful
    // together with a theme reference.
    float tint() const noexcept { return tint_; }
    void setTint(float value) noexcept { tint_ = value; }

    std::size_t hash() const noexcept {
        std::size_t h = std::hash<std::string>{}(argb());
        h ^= static_cast<std::size_t>(theme_) + 0x9e3779b9 + (h << 6) + (h >> 2);
        h ^= std::hash<float>{}(tint_) + 0x9e3779b9 + (h << 6) + (h >> 2);
        return h;
    }
    friend bool operator==(const Color&, const Color&) = default;

    static constexpr int NoTheme = -1;
private:
    static const std::string& defaultArgb() noexcept {
        static const std::string value;
        return value;
    }
    int theme_{NoTheme};
    float tint_{0.0f};
    internal::CompactString argb_;
};
}
