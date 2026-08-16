#pragma once
#include "Color.h"
#include <XLPP/Detail/CompactString.h>
#include <string>
#include <cstddef>
#include <functional>
#include <utility>

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
    // Outlined font (OpenType "outline" style, rare in practice).
    bool outline() const noexcept { return outline_; }
    void setOutline(bool value) noexcept { outline_ = value; }
    // Condensed font (OpenType "condense" style, rare in practice).
    bool condense() const noexcept { return condense_; }
    void setCondense(bool value) noexcept { condense_ = value; }
    // Vertical alignment: "" (baseline), "superscript" or "subscript".
    const std::string& vertAlign() const noexcept { return vertAlign_.get(emptyDefault()); }
    void setVertAlign(std::string value) noexcept { vertAlign_.set(std::move(value), emptyDefault()); }
    // Character set id (0 = ANSI). 1 = Default, 2 = Symbol, 77 = Mac, 128 = SHIFTJIS,
    // 130 = HANGEUL, 134 = GB2312, 162 = GREEK, 163 = TURKISH, 177 = HEBREW, etc.
    unsigned char charset() const noexcept { return charset_; }
    void setCharset(unsigned char value) noexcept { charset_ = value; }
    // Font family: 1 = Roman, 2 = Swiss, 3 = Modern, 4 = Script, 5 = Decorative.
    unsigned char family() const noexcept { return family_; }
    void setFamily(unsigned char value) noexcept { family_ = value; }
    // Font scheme: "" (none), "minor" or "major".
    const std::string& scheme() const noexcept { return scheme_.get(emptyDefault()); }
    void setScheme(std::string value) noexcept { scheme_.set(std::move(value), emptyDefault()); }
    Color& color() noexcept { return color_; }
    const Color& color() const noexcept { return color_; }
    bool isDefault() const noexcept {
        return name_.isDefault() && size_ == 11.0 && !bold_ && !italic_ && !underline_ && !strike_
            && !outline_ && !condense_ && vertAlign_.isDefault() && charset_ == 0 && family_ == 0
            && scheme_.isDefault() && color_.isDefault();
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
        combine(outline_);
        combine(condense_);
        combine(vertAlign());
        combine(charset_);
        combine(family_);
        combine(scheme());
        combine(color_.hash());
        return h;
    }
    friend bool operator==(const Font&, const Font&) = default;
private:
    static const std::string& defaultName() noexcept {
        static const std::string value{"Calibri"};
        return value;
    }
    static const std::string& emptyDefault() noexcept {
        static const std::string value;
        return value;
    }
    internal::CompactString name_;
    double size_{11.0};
    bool bold_{false};
    bool italic_{false};
    bool underline_{false};
    bool strike_{false};
    bool outline_{false};
    bool condense_{false};
    internal::CompactString vertAlign_;
    unsigned char charset_{0};
    unsigned char family_{0};
    internal::CompactString scheme_;
    Color color_;
};
}
