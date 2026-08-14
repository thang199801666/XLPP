#pragma once
#include "Color.h"
#include <XLPP/Detail/CompactString.h>
#include <string>
#include <cstddef>
#include <functional>

namespace xlpp {
class Fill {
public:
    const std::string& patternType() const noexcept { return patternType_.get(defaultPatternType()); }
    void setPatternType(std::string value) { patternType_.set(std::move(value), defaultPatternType()); }
    Color& foregroundColor() noexcept { return foregroundColor_; }
    const Color& foregroundColor() const noexcept { return foregroundColor_; }
    Color& backgroundColor() noexcept { return backgroundColor_; }
    const Color& backgroundColor() const noexcept { return backgroundColor_; }
    bool isDefault() const noexcept {
        return patternType_.isDefault() && foregroundColor_.isDefault() && backgroundColor_.isDefault();
    }

    std::size_t hash() const noexcept {
        std::size_t h = 0;
        auto combine = [&h](auto&& v) { h ^= std::hash<std::decay_t<decltype(v)>>{}(v) + 0x9e3779b9 + (h << 6) + (h >> 2); };
        combine(patternType());
        combine(foregroundColor_.hash());
        combine(backgroundColor_.hash());
        return h;
    }
    friend bool operator==(const Fill&, const Fill&) = default;
private:
    static const std::string& defaultPatternType() noexcept {
        static const std::string value{"none"};
        return value;
    }
    internal::CompactString patternType_;
    Color foregroundColor_;
    Color backgroundColor_;
};
}
