#pragma once

#include <atomic>
#include <chrono>
#include <cstdint>
#include <filesystem>
#include <stdexcept>
#include <string>
#include <system_error>
#include <utility>

#ifdef _WIN32
#  ifndef NOMINMAX
#    define NOMINMAX
#  endif
#  include <windows.h>
#endif

namespace xlpp::internal {

inline std::filesystem::path siblingTemporaryPath(const std::filesystem::path& destination) {
    static std::atomic<std::uint64_t> sequence{0};
    const auto parent = destination.has_parent_path() ? destination.parent_path() : std::filesystem::path{"."};
    const auto stem = destination.filename().native();
    const auto clock = static_cast<std::uint64_t>(
        std::chrono::steady_clock::now().time_since_epoch().count());

    for (std::uint64_t attempt = 0; attempt < 64; ++attempt) {
        const auto id = sequence.fetch_add(1, std::memory_order_relaxed) ^ clock ^ (attempt << 32);
        auto temporaryName = stem;
#ifdef _WIN32
        temporaryName += L".xlpp-tmp-" + std::to_wstring(id);
#else
        temporaryName += ".xlpp-tmp-" + std::to_string(id);
#endif
        const auto candidate = parent / std::filesystem::path(temporaryName);
        std::error_code ec;
        if (!std::filesystem::exists(candidate, ec)) return candidate;
    }
    throw std::runtime_error("Unable to allocate temporary output path for: " + destination.string());
}

inline void replaceFileAtomically(const std::filesystem::path& temporary,
                                  const std::filesystem::path& destination) {
#ifdef _WIN32
    if (!MoveFileExW(temporary.c_str(), destination.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        const auto error = static_cast<unsigned long>(GetLastError());
        throw std::runtime_error("Unable to atomically replace output file (Win32 error "
                                 + std::to_string(error) + "): " + destination.string());
    }
#else
    std::error_code ec;
    std::filesystem::rename(temporary, destination, ec);
    if (ec) {
        throw std::runtime_error("Unable to atomically replace output file: " + destination.string()
                                 + " (" + ec.message() + ")");
    }
#endif
}

template <class Writer>
void atomicWriteFile(const std::filesystem::path& destination, Writer&& writer) {
    const auto temporary = siblingTemporaryPath(destination);
    try {
        std::forward<Writer>(writer)(temporary);
        replaceFileAtomically(temporary, destination);
    } catch (...) {
        std::error_code ignored;
        std::filesystem::remove(temporary, ignored);
        throw;
    }
}

} // namespace xlpp::internal
