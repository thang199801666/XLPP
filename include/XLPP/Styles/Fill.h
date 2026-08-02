#pragma once
#include "Color.h"
#include <string>
#include <cstddef>
#include <functional>

namespace xlpp {
class Fill {
public:
    const std::string& patternType() const noexcept { return patternType_; }
    void setPatternType(std::string value) { patternType_ = std::move(value); }
    Color& foregroundColor() noexcept { return foregroundColor_; }
    const Color& foregroundColor() const noexcept { return foregroundColor_; }
    Color& backgroundColor() noexcept { return backgroundColor_; }
    const Color& backgroundColor() const noexcept { return backgroundColor_; }

    std::size_t hash() const noexcept {
        std::size_t h = 0;
        auto combine = [&h](auto&& v) { h ^= std::hash<std::decay_t<decltype(v)>>{}(v) + 0x9e3779b9 + (h << 6) + (h >> 2); };
        combine(patternType_);
        combine(foregroundColor_.hash());
        combine(backgroundColor_.hash());
        return h;
    }
    friend bool operator==(const Fill&, const Fill&) = default;
private:
    std::string patternType_{"none"};
    Color foregroundColor_;
    Color backgroundColor_;
};
}
