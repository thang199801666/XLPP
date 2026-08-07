#pragma once
#include <cstddef>
#include <filesystem>
#include <fstream>
#include <stdexcept>
#include <algorithm>
#include <string>
#include <vector>

namespace xlpp {

class Worksheet;

enum class DrawingAnchorType {
    OneCell,
    TwoCell,
    Absolute
};

struct DrawingMarker {
    std::size_t row{1};
    std::size_t column{1};
    long long rowOffsetEmu{0};
    long long columnOffsetEmu{0};
};

struct DrawingAnchorInfo {
    DrawingAnchorType type{DrawingAnchorType::OneCell};
    DrawingMarker from{};
    DrawingMarker to{};
    long long xEmu{0};
    long long yEmu{0};
    long long widthEmu{0};
    long long heightEmu{0};
    std::string editAs;
};

class Image {
    friend class Worksheet;
public:
    Image() = default;
    Image(std::string anchor, std::vector<unsigned char> bytes, std::string extension)
        : anchor_(std::move(anchor)), extension_(std::move(extension)), bytes_(std::move(bytes)) {}

    static Image fromFile(const std::filesystem::path& path, std::string anchor) {
        std::ifstream stream(path, std::ios::binary);
        if (!stream) throw std::runtime_error("Cannot open image file: " + path.string());
        std::vector<unsigned char> data((std::istreambuf_iterator<char>(stream)), std::istreambuf_iterator<char>());
        auto extension = path.extension().string();
        if (!extension.empty() && extension.front() == '.') extension.erase(extension.begin());
        std::transform(extension.begin(), extension.end(), extension.begin(), [](unsigned char c){ return static_cast<char>(std::tolower(c)); });
        if (extension == "jpeg") extension = "jpg";
        if (extension != "png" && extension != "jpg") throw std::invalid_argument("XL++ images currently support PNG and JPEG");
        Image result(std::move(anchor), std::move(data), std::move(extension));
        result.name_ = path.stem().string();
        return result;
    }

    const std::string& anchor() const noexcept { return anchor_; }
    void setAnchor(std::string v) { anchor_ = std::move(v); }
    const std::vector<unsigned char>& bytes() const noexcept { return bytes_; }
    const std::string& extension() const noexcept { return extension_; }
    double widthPixels() const noexcept { return widthPixels_; }
    void setWidthPixels(double v) noexcept { widthPixels_ = v; }
    double heightPixels() const noexcept { return heightPixels_; }
    void setHeightPixels(double v) noexcept { heightPixels_ = v; }
    const std::string& name() const noexcept { return name_; }
    void setName(std::string v) { name_ = std::move(v); }

    // Package-origin metadata is populated when an existing worksheet drawing
    // is read. It gives callers a stable read-only view of how Excel anchored
    // the image while allowing XL++ to preserve unsupported drawing XML raw.
    const DrawingAnchorInfo& anchorInfo() const noexcept { return anchorInfo_; }
    const std::string& stableId() const noexcept { return stableId_; }
    const std::string& sourceDrawingPart() const noexcept { return sourceDrawingPart_; }
    const std::string& sourceMediaPart() const noexcept { return sourceMediaPart_; }
    const std::string& sourceRelationshipId() const noexcept { return sourceRelationshipId_; }
    bool imported() const noexcept { return imported_; }

    // Loader-facing setters. Public so bindings/tests can construct package
    // fixtures without friending Workbook internals; normal callers generally
    // only need the const inspection accessors above.
    void setAnchorInfo(DrawingAnchorInfo value) noexcept { anchorInfo_ = std::move(value); }
    void setStableId(std::string value) { stableId_ = std::move(value); }
    void setSourceDrawingPart(std::string value) { sourceDrawingPart_ = std::move(value); }
    void setSourceMediaPart(std::string value) { sourceMediaPart_ = std::move(value); }
    void setSourceRelationshipId(std::string value) { sourceRelationshipId_ = std::move(value); }
    void setImported(bool value) noexcept { imported_ = value; }

private:
    std::string anchor_{"A1"}, extension_{"png"}, name_{"Image"};
    std::vector<unsigned char> bytes_;
    double widthPixels_{96.0}, heightPixels_{96.0};
    DrawingAnchorInfo anchorInfo_{};
    std::string stableId_;
    std::string sourceDrawingPart_;
    std::string sourceMediaPart_;
    std::string sourceRelationshipId_;
    bool imported_{false};
};
}
