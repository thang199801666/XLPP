#pragma once
#include <cstddef>
#include <filesystem>
#include <limits>
#include <stdexcept>
#include <string>
#include <string_view>

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

// Read-only file mapping used by the ZIP readers. Empty files are represented
// by an empty view instead of attempting an invalid zero-length OS mapping.
class MappedFile {
public:
    explicit MappedFile(const std::filesystem::path& path) {
#ifdef _WIN32
        file_ = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
                            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file_ == INVALID_HANDLE_VALUE)
            throw std::runtime_error("Cannot open file: " + path.string());

        LARGE_INTEGER li{};
        if (!GetFileSizeEx(file_, &li)) {
            CloseHandle(file_);
            file_ = INVALID_HANDLE_VALUE;
            throw std::runtime_error("Cannot get file size");
        }
        if (li.QuadPart < 0 || static_cast<unsigned long long>(li.QuadPart) >
                                 static_cast<unsigned long long>(std::numeric_limits<std::size_t>::max())) {
            CloseHandle(file_);
            file_ = INVALID_HANDLE_VALUE;
            throw std::runtime_error("File is too large to map on this platform");
        }
        size_ = static_cast<std::size_t>(li.QuadPart);
        if (size_ == 0) {
            CloseHandle(file_);
            file_ = INVALID_HANDLE_VALUE;
            return;
        }

        mapping_ = CreateFileMappingW(file_, nullptr, PAGE_READONLY, 0, 0, nullptr);
        if (!mapping_) {
            CloseHandle(file_);
            file_ = INVALID_HANDLE_VALUE;
            throw std::runtime_error("Cannot create file mapping");
        }

        view_ = MapViewOfFile(mapping_, FILE_MAP_READ, 0, 0, 0);
        if (!view_) {
            CloseHandle(mapping_);
            mapping_ = nullptr;
            CloseHandle(file_);
            file_ = INVALID_HANDLE_VALUE;
            throw std::runtime_error("Cannot map view");
        }
        CloseHandle(file_);
        file_ = INVALID_HANDLE_VALUE;
#else
        fd_ = ::open(path.c_str(), O_RDONLY);
        if (fd_ < 0) throw std::runtime_error("Cannot open file: " + path.string());

        struct stat st{};
        if (::fstat(fd_, &st) != 0) {
            ::close(fd_);
            fd_ = -1;
            throw std::runtime_error("Cannot stat file");
        }
        if (st.st_size < 0 || static_cast<unsigned long long>(st.st_size) >
                                static_cast<unsigned long long>(std::numeric_limits<std::size_t>::max())) {
            ::close(fd_);
            fd_ = -1;
            throw std::runtime_error("File is too large to map on this platform");
        }
        size_ = static_cast<std::size_t>(st.st_size);
        if (size_ == 0) {
            ::close(fd_);
            fd_ = -1;
            return;
        }

        void* addr = ::mmap(nullptr, size_, PROT_READ, MAP_PRIVATE, fd_, 0);
        if (addr == MAP_FAILED) {
            ::close(fd_);
            fd_ = -1;
            throw std::runtime_error("Cannot mmap file");
        }
        view_ = addr;
        ::close(fd_);
        fd_ = -1;
#endif
    }

    ~MappedFile() { release(); }

    MappedFile(const MappedFile&) = delete;
    MappedFile& operator=(const MappedFile&) = delete;

    MappedFile(MappedFile&& other) noexcept
        : view_(other.view_), size_(other.size_)
#ifdef _WIN32
        , file_(other.file_), mapping_(other.mapping_)
#else
        , fd_(other.fd_)
#endif
    {
        other.view_ = nullptr;
        other.size_ = 0;
#ifdef _WIN32
        other.file_ = INVALID_HANDLE_VALUE;
        other.mapping_ = nullptr;
#else
        other.fd_ = -1;
#endif
    }

    MappedFile& operator=(MappedFile&& other) noexcept {
        if (this != &other) {
            release();
            view_ = other.view_;
            size_ = other.size_;
#ifdef _WIN32
            file_ = other.file_;
            mapping_ = other.mapping_;
            other.file_ = INVALID_HANDLE_VALUE;
            other.mapping_ = nullptr;
#else
            fd_ = other.fd_;
            other.fd_ = -1;
#endif
            other.view_ = nullptr;
            other.size_ = 0;
        }
        return *this;
    }

    const char* data() const noexcept { return static_cast<const char*>(view_); }
    std::size_t size() const noexcept { return size_; }
    std::string_view view() const noexcept {
        return size_ == 0 ? std::string_view{} : std::string_view(data(), size_);
    }

    std::string_view slice(std::size_t offset, std::size_t count) const {
        if (offset > size_ || count > size_ - offset)
            throw std::out_of_range("MappedFile slice out of range");
        if (count == 0) return {};
        return {data() + offset, count};
    }

private:
    void release() noexcept {
#ifdef _WIN32
        if (view_) UnmapViewOfFile(view_);
        if (mapping_) CloseHandle(mapping_);
        if (file_ != INVALID_HANDLE_VALUE) CloseHandle(file_);
        file_ = INVALID_HANDLE_VALUE;
        mapping_ = nullptr;
#else
        if (view_) ::munmap(view_, size_);
        if (fd_ >= 0) ::close(fd_);
        fd_ = -1;
#endif
        view_ = nullptr;
        size_ = 0;
    }

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
