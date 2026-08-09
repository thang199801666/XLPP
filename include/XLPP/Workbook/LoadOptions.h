#pragma once

#include <cstddef>
#include <functional>
#include <string>
#include <vector>

namespace xlpp {

// Issues collected while loading. With LoadOptions::lenient the load continues
// past recoverable sheet-level failures instead of aborting.
struct LoadDiagnostics {
    std::vector<std::string> warnings;
    std::vector<std::string> errors;
    bool hadErrors() const noexcept { return !errors.empty(); }
};

// Guards applied while opening a package for load; zero means unlimited.
struct LoadOptions {
    bool lenient{false};
    std::size_t maxEntries{0};
    std::size_t maxEntryBytes{0};
    std::size_t maxTotalBytes{0};
    std::size_t maxFileBytes{0};
    std::function<bool()> cancel{};
    std::function<void(std::size_t done, std::size_t total)> progress{};
    std::string password{};
    bool verifyEncryptionIntegrity{true};
};

} // namespace xlpp
