#pragma once
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

enum class ShapeType {
    TextBox,
    Rectangle,
    RoundedRectangle,
    Oval,
    Line,
    Arrow
};

enum class ShapeTextHorizontalAlignment {
    Left,
    Center,
    Right,
    Justify
};

enum class ShapeTextVerticalAlignment {
    Top,
    Middle,
    Bottom
};

// A text box or auto-shape anchored to a worksheet cell. Shapes are
// serialized as xdr:sp elements in the worksheet drawing part and require no
// external relationships, so they only affect the drawing XML.
class Shape {
public:
    Shape() = default;
    Shape(ShapeType type, std::string anchor, std::string text)
        : type_(type), anchor_(std::move(anchor)), text_(std::move(text)) {}

    ShapeType type() const noexcept { return type_; }
    void setType(ShapeType value) noexcept { type_ = value; }

    const std::string& anchor() const noexcept { return anchor_; }
    void setAnchor(std::string value) { anchor_ = std::move(value); }

    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }

    const std::string& text() const noexcept { return text_; }
    void setText(std::string value) { text_ = std::move(value); }

    double widthPixels() const noexcept { return widthPixels_; }
    void setWidthPixels(double value) noexcept { widthPixels_ = value; }
    double heightPixels() const noexcept { return heightPixels_; }
    void setHeightPixels(double value) noexcept { heightPixels_ = value; }

    double rotation() const noexcept { return rotation_; }
    void setRotation(double degrees) noexcept { rotation_ = degrees; }
    bool flipHorizontal() const noexcept { return flipH_; }
    void setFlipHorizontal(bool value) noexcept { flipH_ = value; }
    bool flipVertical() const noexcept { return flipV_; }
    void setFlipVertical(bool value) noexcept { flipV_ = value; }

    // Fill. fillColor is an RGB hex string ("RRGGBB", optionally prefixed
    // with an alpha "AARRGGBB"); opacity is 0..1.
    bool hasFill() const noexcept { return hasFill_; }
    void setHasFill(bool value) noexcept { hasFill_ = value; }
    const std::string& fillColor() const noexcept { return fillColor_; }
    void setFillColor(std::string value) { fillColor_ = std::move(value); }
    double fillOpacity() const noexcept { return fillOpacity_; }
    void setFillOpacity(double value) noexcept { fillOpacity_ = value; }

    // Outline. lineWidth is in points.
    bool hasLine() const noexcept { return hasLine_; }
    void setHasLine(bool value) noexcept { hasLine_ = value; }
    const std::string& lineColor() const noexcept { return lineColor_; }
    void setLineColor(std::string value) { lineColor_ = std::move(value); }
    double lineWidthPt() const noexcept { return lineWidthPt_; }
    void setLineWidthPt(double value) noexcept { lineWidthPt_ = value; }

    const std::string& textColor() const noexcept { return textColor_; }
    void setTextColor(std::string value) { textColor_ = std::move(value); }
    double fontSizePt() const noexcept { return fontSizePt_; }
    void setFontSizePt(double value) noexcept { fontSizePt_ = value; }
    bool bold() const noexcept { return bold_; }
    void setBold(bool value) noexcept { bold_ = value; }
    bool italic() const noexcept { return italic_; }
    void setItalic(bool value) noexcept { italic_ = value; }
    ShapeTextHorizontalAlignment textHorizontalAlignment() const noexcept { return textHAlign_; }
    void setTextHorizontalAlignment(ShapeTextHorizontalAlignment value) noexcept { textHAlign_ = value; }
    ShapeTextVerticalAlignment textVerticalAlignment() const noexcept { return textVAlign_; }
    void setTextVerticalAlignment(ShapeTextVerticalAlignment value) noexcept { textVAlign_ = value; }
    bool wordWrap() const noexcept { return wordWrap_; }
    void setWordWrap(bool value) noexcept { wordWrap_ = value; }
    bool autoSizeHeight() const noexcept { return autoSizeHeight_; }
    void setAutoSizeHeight(bool value) noexcept { autoSizeHeight_ = value; }

private:
    ShapeType type_{ShapeType::TextBox};
    std::string anchor_{"A1"}, name_{"TextBox 1"}, text_;
    double widthPixels_{150.0}, heightPixels_{75.0};
    double rotation_{0.0};
    bool flipH_{false}, flipV_{false};
    bool hasFill_{true};
    std::string fillColor_{"FFFFFF"};
    double fillOpacity_{1.0};
    bool hasLine_{true};
    std::string lineColor_{"000000"};
    double lineWidthPt_{0.75};
    std::string textColor_{"000000"};
    double fontSizePt_{11.0};
    bool bold_{false}, italic_{false};
    ShapeTextHorizontalAlignment textHAlign_{ShapeTextHorizontalAlignment::Left};
    ShapeTextVerticalAlignment textVAlign_{ShapeTextVerticalAlignment::Top};
    bool wordWrap_{true}, autoSizeHeight_{true};
};

} // namespace xlpp
