#include <XLPP/Workbook/Workbook.h>
#include "../Encryption/OfficeCrypto.h"

#include <array>
#include <atomic>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <functional>
#include <stdexcept>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>

namespace {

std::filesystem::path uniqueTemporaryPath(std::string_view prefix, std::string_view extension) {
    static std::atomic<std::uint64_t> counter{0};
    const auto stamp = static_cast<std::uint64_t>(
        std::chrono::high_resolution_clock::now().time_since_epoch().count());
    const auto sequence = counter.fetch_add(1, std::memory_order_relaxed);
    const auto threadId = static_cast<std::uint64_t>(std::hash<std::thread::id>{}(std::this_thread::get_id()));
    return std::filesystem::temp_directory_path() /
        (std::string(prefix) + std::to_string(stamp) + "_" + std::to_string(threadId) + "_" +
         std::to_string(sequence) + std::string(extension));
}

class ScopedTemporaryFile final {
public:
    explicit ScopedTemporaryFile(std::filesystem::path path) : path_(std::move(path)) {}
    ScopedTemporaryFile(const ScopedTemporaryFile&) = delete;
    ScopedTemporaryFile& operator=(const ScopedTemporaryFile&) = delete;
    ~ScopedTemporaryFile() noexcept {
        std::error_code ec;
        std::filesystem::remove(path_, ec);
    }
    const std::filesystem::path& path() const noexcept { return path_; }
private:
    std::filesystem::path path_;
};

void copyInputStreamToFile(std::istream& input,
                           const std::filesystem::path& path,
                           std::size_t maxBytes,
                           const std::function<bool()>& cancel) {
    std::ofstream output(path, std::ios::binary | std::ios::trunc);
    if (!output) throw std::runtime_error("Unable to create temporary workbook input file");

    std::array<char, 64u * 1024u> buffer{};
    std::size_t total = 0;
    for (;;) {
        if (cancel && cancel()) throw std::runtime_error("Workbook load cancelled while reading input stream");
        input.read(buffer.data(), static_cast<std::streamsize>(buffer.size()));
        const auto got = input.gcount();
        if (got > 0) {
            const auto count = static_cast<std::size_t>(got);
            if (maxBytes != 0 && (total > maxBytes || count > maxBytes - total))
                throw std::runtime_error("Workbook input stream exceeds configured maxFileBytes limit");
            output.write(buffer.data(), got);
            if (!output) throw std::runtime_error("Unable to write temporary workbook input file");
            total += count;
        }
        if (input.eof()) break;
        if (input.bad() || (input.fail() && got == 0))
            throw std::runtime_error("Failed while reading workbook input stream");
    }
    output.flush();
    if (!output) throw std::runtime_error("Unable to flush temporary workbook input file");
}

void copyFileToOutputStream(const std::filesystem::path& path, std::ostream& output) {
    std::ifstream input(path, std::ios::binary);
    if (!input) throw std::runtime_error("Unable to open temporary workbook output file");

    std::array<char, 64u * 1024u> buffer{};
    for (;;) {
        input.read(buffer.data(), static_cast<std::streamsize>(buffer.size()));
        const auto got = input.gcount();
        if (got > 0) {
            output.write(buffer.data(), got);
            if (!output) throw std::runtime_error("Failed while writing workbook output stream");
        }
        if (input.eof()) break;
        if (input.bad() || (input.fail() && got == 0))
            throw std::runtime_error("Failed while reading temporary workbook output file");
    }
}

} // namespace

namespace xlpp {

bool Workbook::isPasswordEncryptedFile(const std::filesystem::path& path) noexcept {
    return internal::isEncryptedOfficeCompoundFile(path);
}

PackageEncryptionInfo Workbook::inspectPasswordEncryptionFile(const std::filesystem::path& path) {
    return internal::inspectOfficeEncryption(path);
}

void Workbook::load(std::istream& stream) { load(stream, LoadOptions{}); }

void Workbook::load(std::istream& stream, const LoadOptions& options) {
    ScopedTemporaryFile temporary(uniqueTemporaryPath("xlpp_stream_load_", ".tmp"));
    copyInputStreamToFile(stream, temporary.path(), options.maxFileBytes, options.cancel);
    load(temporary.path(), options);
}

void Workbook::save(std::ostream& stream) const { save(stream, SaveOptions{}); }

void Workbook::save(std::ostream& stream, const SaveOptions& options) const {
    ScopedTemporaryFile temporary(uniqueTemporaryPath("xlpp_stream_save_", ".tmp"));
    save(temporary.path(), options);
    copyFileToOutputStream(temporary.path(), stream);
}

} // namespace xlpp
