#pragma once
#include <filesystem>
#include <string>

namespace xlpp_numbered_tests {
inline std::filesystem::path& outputDirectoryStorage() {
    static std::filesystem::path path = std::filesystem::current_path() / "XLPP.TestOutputs";
    return path;
}
inline void setOutputDirectory(std::filesystem::path path) {
    outputDirectoryStorage() = std::move(path);
    std::filesystem::create_directories(outputDirectoryStorage());
}
inline std::filesystem::path outputPath(const std::string& fileName) {
    std::filesystem::create_directories(outputDirectoryStorage());
    return outputDirectoryStorage() / fileName;
}
}
