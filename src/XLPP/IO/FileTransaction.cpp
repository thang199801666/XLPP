#include "IO/FileTransaction.h"

#include <atomic>
#include <cerrno>
#include <random>
#include <stdexcept>
#include <system_error>

#ifdef _WIN32
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#else
#include <fcntl.h>
#include <unistd.h>
#endif

namespace xlpp::internal {
namespace {
std::uint64_t temporaryToken() {
    static std::atomic<std::uint64_t> sequence{0};
    std::random_device random;
    return (static_cast<std::uint64_t>(random()) << 32u) ^
           static_cast<std::uint64_t>(random()) ^
           sequence.fetch_add(1, std::memory_order_relaxed);
}

#ifndef _WIN32
void fsyncDescriptor(int fd, const char* what) {
    if (::fsync(fd) != 0) {
        const int error = errno;
        (void)::close(fd);
        throw std::system_error(error, std::generic_category(), what);
    }
    if (::close(fd) != 0)
        throw std::system_error(errno, std::generic_category(), what);
}

void syncDirectoryToDisk(const std::filesystem::path& directory) {
    const auto dir = directory.empty() ? std::filesystem::current_path() : directory;
#ifdef O_DIRECTORY
    const int fd = ::open(dir.c_str(), O_RDONLY | O_DIRECTORY);
#else
    const int fd = ::open(dir.c_str(), O_RDONLY);
#endif
    if (fd < 0)
        throw std::system_error(errno, std::generic_category(), "Cannot open workbook directory for fsync");
    fsyncDescriptor(fd, "Cannot fsync workbook directory");
}
#endif
} // namespace

std::filesystem::path uniqueTemporaryPath(const std::filesystem::path& directory,
                                          std::string_view prefix) {
    for (int attempt = 0; attempt < 64; ++attempt) {
        auto path = directory / (std::string(prefix) + std::to_string(temporaryToken()) + ".tmp");
        std::error_code error;
        if (!std::filesystem::exists(path, error) && !error) return path;
    }
    throw std::runtime_error("Unable to allocate a temporary XL++ file");
}

std::filesystem::path xlppTemporaryPath(std::string_view purpose) {
    return uniqueTemporaryPath(std::filesystem::temp_directory_path(),
                               "xlpp_" + std::string(purpose) + "_");
}

std::filesystem::path atomicSaveTemporaryPath(const std::filesystem::path& target) {
    auto directory = target.parent_path();
    if (directory.empty()) directory = std::filesystem::current_path();
    return uniqueTemporaryPath(directory, "." + target.filename().string() + ".xlpp_");
}

void syncFileToDisk(const std::filesystem::path& path) {
#ifdef _WIN32
    // FlushFileBuffers requires a writable handle on Windows. The previous
    // read-only handle caused atomic saves to fail with ERROR_ACCESS_DENIED
    // even though the staging file itself had been written successfully.
    HANDLE handle = ::CreateFileW(path.c_str(), GENERIC_READ | GENERIC_WRITE,
                                  FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                  nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (handle == INVALID_HANDLE_VALUE)
        throw std::system_error(static_cast<int>(::GetLastError()), std::system_category(),
                                "Cannot open workbook for durable flush");
    if (!::FlushFileBuffers(handle)) {
        const auto error = static_cast<int>(::GetLastError());
        ::CloseHandle(handle);
        throw std::system_error(error, std::system_category(), "Cannot flush workbook to disk");
    }
    if (!::CloseHandle(handle))
        throw std::system_error(static_cast<int>(::GetLastError()), std::system_category(),
                                "Cannot close workbook after durable flush");
#else
    const int fd = ::open(path.c_str(), O_RDONLY);
    if (fd < 0)
        throw std::system_error(errno, std::generic_category(), "Cannot open workbook for fsync");
    fsyncDescriptor(fd, "Cannot fsync workbook");
#endif
}

void replaceFileAtomically(const std::filesystem::path& temporary,
                           const std::filesystem::path& destination,
                           bool durable) {
    if (durable) syncFileToDisk(temporary);
#ifdef _WIN32
    DWORD flags = MOVEFILE_REPLACE_EXISTING;
    if (durable) flags |= MOVEFILE_WRITE_THROUGH;
    if (!::MoveFileExW(temporary.c_str(), destination.c_str(), flags)) {
        throw std::system_error(static_cast<int>(::GetLastError()), std::system_category(),
                                "Cannot atomically replace workbook");
    }
#else
    if (::rename(temporary.c_str(), destination.c_str()) != 0) {
        throw std::system_error(errno, std::generic_category(),
                                "Cannot atomically replace workbook");
    }
    if (durable) syncDirectoryToDisk(destination.parent_path());
#endif
}

ScopedTemporaryFile::~ScopedTemporaryFile() {
    if (!active_) return;
    std::error_code ignored;
    std::filesystem::remove(path_, ignored);
}

} // namespace xlpp::internal
