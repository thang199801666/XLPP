#pragma once
#include <cstddef>
#include <string_view>
#include <filesystem>
#include <stdexcept>
#include <string>

#ifndef _WIN32
#error MappedFile currently requires Windows
#endif

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

namespace xlpp::internal {

// Memory-mapped file for zero-copy read access. The entire file is mapped as a
// single read-only view, providing string_view access without copying.
class MappedFile {
public:
    explicit MappedFile(const std::filesystem::path& path) {
        file_ = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
                            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file_ == INVALID_HANDLE_VALUE) {
            throw std::runtime_error("Cannot open file: " + path.string());
        }

        LARGE_INTEGER li;
        if (!GetFileSizeEx(file_, &li)) {
            CloseHandle(file_);
            file_ = INVALID_HANDLE_VALUE;
            throw std::runtime_error("Cannot get file size: " + path.string());
        }
        size_ = static_cast<std::size_t>(li.QuadPart);

        mapping_ = CreateFileMappingW(file_, nullptr, PAGE_READONLY,
                                      static_cast<DWORD>(li.HighPart),
                                      static_cast<DWORD>(li.LowPart), nullptr);
        if (!mapping_) {
            CloseHandle(file_);
            file_ = INVALID_HANDLE_VALUE;
            throw std::runtime_error("Cannot create file mapping: " + path.string());
        }

        view_ = MapViewOfFile(mapping_, FILE_MAP_READ, 0, 0, 0);
        if (!view_) {
            CloseHandle(mapping_);
            mapping_ = nullptr;
            CloseHandle(file_);
            file_ = INVALID_HANDLE_VALUE;
            throw std::runtime_error("Cannot map view of file: " + path.string());
        }
        // The file handle can be closed after a successful mapping; the mapping
        // object holds its own reference to the file data.
        CloseHandle(file_);
        file_ = INVALID_HANDLE_VALUE;
    }

    ~MappedFile() {
        if (view_) UnmapViewOfFile(view_);
        if (mapping_) CloseHandle(mapping_);
        if (file_ != INVALID_HANDLE_VALUE) CloseHandle(file_);
    }

    MappedFile(const MappedFile&) = delete;
    MappedFile& operator=(const MappedFile&) = delete;
    MappedFile(MappedFile&& other) noexcept
        : file_(other.file_), mapping_(other.mapping_), view_(other.view_), size_(other.size_) {
        other.file_ = INVALID_HANDLE_VALUE;
        other.mapping_ = nullptr;
        other.view_ = nullptr;
        other.size_ = 0;
    }
    MappedFile& operator=(MappedFile&& other) noexcept {
        if (this != &other) {
            if (view_) UnmapViewOfFile(view_);
            if (mapping_) CloseHandle(mapping_);
            if (file_ != INVALID_HANDLE_VALUE) CloseHandle(file_);
            file_ = other.file_;
            mapping_ = other.mapping_;
            view_ = other.view_;
            size_ = other.size_;
            other.file_ = INVALID_HANDLE_VALUE;
            other.mapping_ = nullptr;
            other.view_ = nullptr;
            other.size_ = 0;
        }
        return *this;
    }

    const char* data() const noexcept { return static_cast<const char*>(view_); }
    std::size_t size() const noexcept { return size_; }
    std::string_view view() const noexcept { return {data(), size_}; }

    // Access a slice of the mapped file as a string_view. Bounds-checked in debug.
    std::string_view slice(std::size_t offset, std::size_t count) const {
        if (offset + count > size_) throw std::out_of_range("MappedFile slice out of range");
        return {data() + offset, count};
    }

private:
    HANDLE file_{INVALID_HANDLE_VALUE};
    HANDLE mapping_{nullptr};
    LPVOID view_{nullptr};
    std::size_t size_{0};
};

} // namespace xlpp::internal
