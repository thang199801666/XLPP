#pragma once
#include <string>

namespace xlpp {
enum class PageOrientation { Default, Portrait, Landscape };
enum class PaperSize : unsigned { Default = 0, Letter = 1, Legal = 5, A4 = 9, A3 = 8 };

class PageMargins {
public:
    double left() const noexcept { return left_; } void setLeft(double v) noexcept { left_ = v; }
    double right() const noexcept { return right_; } void setRight(double v) noexcept { right_ = v; }
    double top() const noexcept { return top_; } void setTop(double v) noexcept { top_ = v; }
    double bottom() const noexcept { return bottom_; } void setBottom(double v) noexcept { bottom_ = v; }
    double header() const noexcept { return header_; } void setHeader(double v) noexcept { header_ = v; }
    double footer() const noexcept { return footer_; } void setFooter(double v) noexcept { footer_ = v; }
private:
    double left_{0.7}, right_{0.7}, top_{0.75}, bottom_{0.75}, header_{0.3}, footer_{0.3};
};

class PageSetup {
public:
    PageOrientation orientation() const noexcept { return orientation_; }
    void setOrientation(PageOrientation v) noexcept { orientation_ = v; }
    PaperSize paperSize() const noexcept { return paperSize_; }
    void setPaperSize(PaperSize v) noexcept { paperSize_ = v; }
    unsigned scale() const noexcept { return scale_; }
    void setScale(unsigned v) noexcept { scale_ = v; }
    unsigned fitToWidth() const noexcept { return fitToWidth_; }
    void setFitToWidth(unsigned v) noexcept { fitToWidth_ = v; }
    unsigned fitToHeight() const noexcept { return fitToHeight_; }
    void setFitToHeight(unsigned v) noexcept { fitToHeight_ = v; }
    bool fitToPage() const noexcept { return fitToPage_; }
    void setFitToPage(bool v) noexcept { fitToPage_ = v; }
    bool blackAndWhite() const noexcept { return blackAndWhite_; }
    void setBlackAndWhite(bool v) noexcept { blackAndWhite_ = v; }
    bool draft() const noexcept { return draft_; }
    void setDraft(bool v) noexcept { draft_ = v; }
    unsigned firstPageNumber() const noexcept { return firstPageNumber_; }
    void setFirstPageNumber(unsigned v) noexcept { firstPageNumber_ = v; }
    bool useFirstPageNumber() const noexcept { return useFirstPageNumber_; }
    void setUseFirstPageNumber(bool v) noexcept { useFirstPageNumber_ = v; }
private:
    PageOrientation orientation_{PageOrientation::Default};
    PaperSize paperSize_{PaperSize::Default};
    unsigned scale_{100}, fitToWidth_{0}, fitToHeight_{0}, firstPageNumber_{1};
    bool fitToPage_{false}, blackAndWhite_{false}, draft_{false}, useFirstPageNumber_{false};
};

class PrintOptions {
public:
    bool horizontalCentered() const noexcept { return horizontalCentered_; }
    void setHorizontalCentered(bool v) noexcept { horizontalCentered_ = v; }
    bool verticalCentered() const noexcept { return verticalCentered_; }
    void setVerticalCentered(bool v) noexcept { verticalCentered_ = v; }
    bool headings() const noexcept { return headings_; }
    void setHeadings(bool v) noexcept { headings_ = v; }
    bool gridLines() const noexcept { return gridLines_; }
    void setGridLines(bool v) noexcept { gridLines_ = v; }
private:
    bool horizontalCentered_{false}, verticalCentered_{false}, headings_{false}, gridLines_{false};
};

class HeaderFooter {
public:
    const std::string& oddHeader() const noexcept { return oddHeader_; }
    void setOddHeader(std::string v) { oddHeader_ = std::move(v); }
    const std::string& oddFooter() const noexcept { return oddFooter_; }
    void setOddFooter(std::string v) { oddFooter_ = std::move(v); }
    const std::string& evenHeader() const noexcept { return evenHeader_; }
    void setEvenHeader(std::string v) { evenHeader_ = std::move(v); }
    const std::string& evenFooter() const noexcept { return evenFooter_; }
    void setEvenFooter(std::string v) { evenFooter_ = std::move(v); }
    bool differentOddEven() const noexcept { return differentOddEven_; }
    void setDifferentOddEven(bool v) noexcept { differentOddEven_ = v; }
    bool differentFirst() const noexcept { return differentFirst_; }
    void setDifferentFirst(bool v) noexcept { differentFirst_ = v; }
private:
    std::string oddHeader_, oddFooter_, evenHeader_, evenFooter_;
    bool differentOddEven_{false}, differentFirst_{false};
};
}
