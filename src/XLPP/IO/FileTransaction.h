#pragma once

#include <cstdint>
#include <filesystem>
#include <string_view>

namespace xlpp::internal {

std::filesystem::path uniqueTemporaryPath(const std::filesystem::path& directory,
                                          std::string_view prefix);
std::filesystem::path xlppTemporaryPath(std::string_view purpose);
std::filesystem::path atomicSaveTemporaryPath(const std::filesystem::path& target);
void syncFileToDisk(const std::filesystem::path& path);
void replaceFileAtomically(const std::filesystem::path& temporary,
                           const std::filesystem::path& destination,
                           bool durable = true);

class ScopedTemporaryFile {
public:
    explicit ScopedTemporaryFile(std::filesystem::path path) : path_(std::move(path)) {}
    ScopedTemporaryFile(const ScopedTemporaryFile&) = delete;
    ScopedTemporaryFile& operator=(const ScopedTemporaryFile&) = delete;
    ScopedTemporaryFile(ScopedTemporaryFile&&) = delete;
    ScopedTemporaryFile& operator=(ScopedTemporaryFile&&) = delete;
    ~ScopedTemporaryFile();

    const std::filesystem::path& path() const noexcept { return path_; }
    void release() noexcept { active_ = false; }

private:
    std::filesystem::path path_;
    bool active_{true};
};

} // namespace xlpp::internal
