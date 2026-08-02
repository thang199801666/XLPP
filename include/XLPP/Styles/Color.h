#pragma once
#include <string>
#include <utility>
#include <cstddef>
#include <functional>

namespace xlpp {
class Color {
public:
    Color() = default;
    explicit Color(std::string argb) : argb_(std::move(argb)) {}

    const std::string& argb() const noexcept { return argb_; }
    void setArgb(std::string value) { argb_ = std::move(value); }
    bool empty() const noexcept { return argb_.empty(); }

    std::size_t hash() const noexcept { return std::hash<std::string>{}(argb_); }
    friend bool operator==(const Color&, const Color&) = default;
private:
    std::string argb_;
};
}
