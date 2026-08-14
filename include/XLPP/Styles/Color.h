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
    explicit Color(std::string argb) { argb_.set(std::move(argb), defaultArgb()); }

    const std::string& argb() const noexcept { return argb_.get(defaultArgb()); }
    void setArgb(std::string value) { argb_.set(std::move(value), defaultArgb()); }
    bool empty() const noexcept { return argb_.isDefault(); }
    bool isDefault() const noexcept { return argb_.isDefault(); }

    std::size_t hash() const noexcept { return std::hash<std::string>{}(argb()); }
    friend bool operator==(const Color&, const Color&) = default;
private:
    static const std::string& defaultArgb() noexcept {
        static const std::string value;
        return value;
    }
    internal::CompactString argb_;
};
}
