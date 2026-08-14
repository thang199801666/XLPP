#pragma once
#include "Font.h"
#include "Fill.h"
#include "Border.h"
#include "Alignment.h"
#include <XLPP/Detail/CompactString.h>
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
    const std::string& numberFormat() const noexcept { return numberFormat_.get(defaultNumberFormat()); }
    void setNumberFormat(std::string value) { numberFormat_.set(std::move(value), defaultNumberFormat()); }
    int numFmtId() const noexcept { return numFmtId_; }
    void setNumFmtId(int value) noexcept { numFmtId_ = value; }
    bool locked() const noexcept { return locked_; }
    void setLocked(bool value) noexcept { locked_ = value; }
    bool hidden() const noexcept { return hidden_; }
    void setHidden(bool value) noexcept { hidden_ = value; }

    bool isDefault() const noexcept {
        return font_.isDefault() && fill_.isDefault() && border_.isDefault() && alignment_.isDefault() &&
               numberFormat_.isDefault() && numFmtId_ == 0 && locked_ && !hidden_;
    }
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
        combine(h, numberFormat());
        combine(h, numFmtId_);
        combine(h, locked_);
        combine(h, hidden_);
        return h;
    }
private:
    static const std::string& defaultNumberFormat() noexcept {
        static const std::string value{"General"};
        return value;
    }
    Font font_;
    Fill fill_;
    Border border_;
    Alignment alignment_;
    internal::CompactString numberFormat_;
    int numFmtId_{0};
    bool locked_{true};
    bool hidden_{false};
};
}
