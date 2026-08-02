#pragma once
#include "Font.h"
#include "Fill.h"
#include "Border.h"
#include "Alignment.h"
#include <string>
#include <cstddef>
#include <functional>

namespace xlpp {
class Style {
public:
    Font& font() noexcept { return font_; }
    const Font& font() const noexcept { return font_; }
    Fill& fill() noexcept { return fill_; }
    const Fill& fill() const noexcept { return fill_; }
    Border& border() noexcept { return border_; }
    const Border& border() const noexcept { return border_; }
    Alignment& alignment() noexcept { return alignment_; }
    const Alignment& alignment() const noexcept { return alignment_; }
    const std::string& numberFormat() const noexcept { return numberFormat_; }
    void setNumberFormat(std::string value) { numberFormat_ = std::move(value); }
    int numFmtId() const noexcept { return numFmtId_; }
    void setNumFmtId(int value) noexcept { numFmtId_ = value; }
    bool locked() const noexcept { return locked_; }
    void setLocked(bool value) noexcept { locked_ = value; }
    bool hidden() const noexcept { return hidden_; }
    void setHidden(bool value) noexcept { hidden_ = value; }

    bool isDefault() const noexcept { return *this == Style{}; }
    friend bool operator==(const Style&, const Style&) = default;

    std::size_t hash() const noexcept {
        auto combine = [](std::size_t& seed, auto&& value) {
            seed ^= std::hash<std::decay_t<decltype(value)>>{}(value) + 0x9e3779b9 + (seed << 6) + (seed >> 2);
        };
        std::size_t h = 0;
        combine(h, font_.hash());
        combine(h, fill_.hash());
        combine(h, border_.hash());
        combine(h, alignment_.hash());
        combine(h, numberFormat_);
        combine(h, numFmtId_);
        combine(h, locked_);
        combine(h, hidden_);
        return h;
    }
private:
    Font font_;
    Fill fill_;
    Border border_;
    Alignment alignment_;
    std::string numberFormat_{"General"};
    int numFmtId_{0};
    bool locked_{true};
    bool hidden_{false};
};
}
