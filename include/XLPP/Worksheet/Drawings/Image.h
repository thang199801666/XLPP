#pragma once
#include <cstddef>
#include <filesystem>
#include <fstream>
#include <stdexcept>
#include <algorithm>
#include <string>
#include <vector>
namespace xlpp {
class Image {
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
private:
    std::string anchor_{"A1"}, extension_{"png"}, name_{"Image"};
    std::vector<unsigned char> bytes_;
    double widthPixels_{96.0}, heightPixels_{96.0};
};
}
