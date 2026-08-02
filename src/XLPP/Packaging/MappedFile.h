#pragma once
#include <cstddef>
#include <string_view>
#include <filesystem>
#include <stdexcept>
#include <string>

#ifdef _WIN32
  #ifndef NOMINMAX
  #define NOMINMAX
  #endif
  #include <windows.h>
#else
  #include <fcntl.h>
  #include <sys/mman.h>
  #include <sys/stat.h>
  #include <unistd.h>
#endif

namespace xlpp::internal {

class MappedFile {
public:
    explicit MappedFile(const std::filesystem::path& path) {
#ifdef _WIN32
        file_ = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
                            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file_ == INVALID_HANDLE_VALUE)
            throw std::runtime_error("Cannot open file: " + path.string());

        LARGE_INTEGER li;
        if (!GetFileSizeEx(file_, &li)) { CloseHandle(file_); throw std::runtime_error("Cannot get file size"); }
        size_ = static_cast<std::size_t>(li.QuadPart);

        mapping_ = CreateFileMappingW(file_, nullptr, PAGE_READONLY, li.HighPart, li.LowPart, nullptr);
        if (!mapping_) { CloseHandle(file_); throw std::runtime_error("Cannot create file mapping"); }

        view_ = MapViewOfFile(mapping_, FILE_MAP_READ, 0, 0, 0);
        if (!view_) { CloseHandle(mapping_); CloseHandle(file_); throw std::runtime_error("Cannot map view"); }
        CloseHandle(file_);
#else
        fd_ = ::open(path.c_str(), O_RDONLY);
        if (fd_ < 0) throw std::runtime_error("Cannot open file: " + path.string());

        struct stat st;
        if (::fstat(fd_, &st) != 0) { ::close(fd_); throw std::runtime_error("Cannot stat file"); }
        size_ = static_cast<std::size_t>(st.st_size);

        void* addr = ::mmap(nullptr, size_, PROT_READ, MAP_PRIVATE, fd_, 0);
        if (addr == MAP_FAILED) { ::close(fd_); throw std::runtime_error("Cannot mmap file"); }
        view_ = addr;
        ::close(fd_);
#endif
    }

    ~MappedFile() {
        if (view_) {
#ifdef _WIN32
            UnmapViewOfFile(view_);
            if (mapping_) CloseHandle(mapping_);
#else
            ::munmap(view_, size_);
#endif
        }
    }

    MappedFile(const MappedFile&) = delete;
    MappedFile& operator=(const MappedFile&) = delete;
    MappedFile(MappedFile&& other) noexcept
        : view_(other.view_), size_(other.size_)
#ifdef _WIN32
        , mapping_(other.mapping_)
#endif
    {
        other.view_ = nullptr;
        other.size_ = 0;
#ifdef _WIN32
        other.mapping_ = nullptr;
#endif
    }
    MappedFile& operator=(MappedFile&& other) noexcept {
        if (this != &other) {
            if (view_) {
#ifdef _WIN32
                UnmapViewOfFile(view_);
                if (mapping_) CloseHandle(mapping_);
#else
                ::munmap(view_, size_);
#endif
            }
            view_ = other.view_;
            size_ = other.size_;
#ifdef _WIN32
            mapping_ = other.mapping_;
#endif
            other.view_ = nullptr;
            other.size_ = 0;
#ifdef _WIN32
            other.mapping_ = nullptr;
#endif
        }
        return *this;
    }

    const char* data() const noexcept { return static_cast<const char*>(view_); }
    std::size_t size() const noexcept { return size_; }
    std::string_view view() const noexcept { return {data(), size_}; }

    std::string_view slice(std::size_t offset, std::size_t count) const {
        if (offset + count > size_) throw std::out_of_range("MappedFile slice out of range");
        return {data() + offset, count};
    }

private:
#ifdef _WIN32
    HANDLE file_{INVALID_HANDLE_VALUE};
    HANDLE mapping_{nullptr};
#else
    int fd_{-1};
#endif
    void* view_{nullptr};
    std::size_t size_{0};
};

} // namespace xlpp::internal
