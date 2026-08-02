#pragma once
#include <string>
#include <optional>

namespace xlpp {

class SheetView {
public:
    int workbookViewId() const noexcept { return workbookViewId_; }
    void setWorkbookViewId(int v) noexcept { workbookViewId_ = v; }

    const std::optional<std::string>& tabColor() const noexcept { return tabColor_; }
    void setTabColor(std::string argb) { tabColor_ = std::move(argb); }
    void clearTabColor() noexcept { tabColor_.reset(); }

    int zoomScale() const noexcept { return zoomScale_; }
    void setZoomScale(int v) noexcept { zoomScale_ = v; }
    int zoomScaleNormal() const noexcept { return zoomScaleNormal_; }
    void setZoomScaleNormal(int v) noexcept { zoomScaleNormal_ = v; }

    bool showGridLines() const noexcept { return showGridLines_; }
    void setShowGridLines(bool v) noexcept { showGridLines_ = v; }
    bool tabSelected() const noexcept { return tabSelected_; }
    void setTabSelected(bool v) noexcept { tabSelected_ = v; }
    bool rightToLeft() const noexcept { return rightToLeft_; }
    void setRightToLeft(bool v) noexcept { rightToLeft_ = v; }
    bool showOutlineSymbols() const noexcept { return showOutlineSymbols_; }
    void setShowOutlineSymbols(bool v) noexcept { showOutlineSymbols_ = v; }

    const std::string& pane() const noexcept { return pane_; }
    void setPane(std::string v) { pane_ = std::move(v); }
    const std::string& topLeftCell() const noexcept { return topLeftCell_; }
    void setTopLeftCell(std::string v) { topLeftCell_ = std::move(v); }
    int xSplit() const noexcept { return xSplit_; } void setXSplit(int v) noexcept { xSplit_ = v; }
    int ySplit() const noexcept { return ySplit_; } void setYSplit(int v) noexcept { ySplit_ = v; }

private:
    int workbookViewId_{0};
    std::optional<std::string> tabColor_;
    int zoomScale_{100}, zoomScaleNormal_{100};
    bool showGridLines_{true}, tabSelected_{false}, rightToLeft_{false}, showOutlineSymbols_{true};
    std::string pane_, topLeftCell_;
    int xSplit_{0}, ySplit_{0};
};

}
