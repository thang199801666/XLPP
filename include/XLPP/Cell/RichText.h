#pragma once
#include <string>
#include <optional>
#include <vector>

namespace xlpp {

class RichTextRun {
public:
    RichTextRun() = default;
    explicit RichTextRun(std::string text) : text_(std::move(text)) {}

    const std::string& text() const noexcept { return text_; }
    void setText(std::string v) { text_ = std::move(v); }
    bool bold() const noexcept { return bold_; } void setBold(bool v) noexcept { bold_ = v; }
    bool italic() const noexcept { return italic_; } void setItalic(bool v) noexcept { italic_ = v; }
    bool underline() const noexcept { return underline_; } void setUnderline(bool v) noexcept { underline_ = v; }
    bool strike() const noexcept { return strike_; } void setStrike(bool v) noexcept { strike_ = v; }
    const std::string& color() const noexcept { return color_; } void setColor(std::string v) { color_ = std::move(v); }
    double size() const noexcept { return size_; } void setSize(double v) noexcept { size_ = v; }
    const std::string& fontName() const noexcept { return fontName_; } void setFontName(std::string v) { fontName_ = std::move(v); }

private:
    std::string text_;
    std::string color_, fontName_;
    double size_{0.0};
    bool bold_{false}, italic_{false}, underline_{false}, strike_{false};
};

class RichText {
public:
    void addRun(RichTextRun run) { runs_.push_back(std::move(run)); }
    const std::vector<RichTextRun>& runs() const noexcept { return runs_; }
    std::vector<RichTextRun>& runs() noexcept { return runs_; }
    bool empty() const noexcept { return runs_.empty(); }
    std::string plainText() const {
        std::string result;
        for (const auto& r : runs_) result += r.text();
        return result;
    }

    static RichText fromPlain(std::string text) {
        RichText rt;
        rt.addRun(RichTextRun(std::move(text)));
        return rt;
    }

private:
    std::vector<RichTextRun> runs_;
};

}
