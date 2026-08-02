#pragma once
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <map>
#include <string>
#include <vector>

namespace xlpp::internal {
class ThreadPool;

// Guards applied while opening a package: 0 means unlimited.
struct ZipOpenLimits {
    std::size_t maxEntries{0};
    std::size_t maxEntryBytes{0};
    std::size_t maxTotalBytes{0};
    std::size_t maxFileBytes{0};
    std::function<bool()> cancel{};
    std::function<void(std::size_t done, std::size_t total)> progress{};
};

// Callbacks honored while writing a package.
struct ZipWriteOptions {
    std::function<bool()> cancel{};
    std::function<void(std::size_t done, std::size_t total)> progress{};
};

class ZipArchive {
public:
    void add(std::string name, std::string data, bool compress = true);
    void addFile(std::string name, std::filesystem::path sourcePath, bool compress = true);
    void save(const std::filesystem::path& path) const;
    void save(const std::filesystem::path& path, const ZipWriteOptions& options) const;
    static ZipArchive open(const std::filesystem::path& path);
    static ZipArchive open(const std::filesystem::path& path, const ZipOpenLimits& limits);
    bool contains(const std::string& name) const;
    const std::string& get(const std::string& name) const;
    // All entry names in the archive, in sorted order.
    std::vector<std::string> entryNames() const;

    // zlib level: -1 = default, 0 = store, 1-9. Strategy: 0 = default,
    // 1 = filtered, 2 = huffman only, 3 = RLE, 4 = fixed.
    void setCompressionLevel(int level) { compressionLevel_ = level; }
    void setCompressionStrategy(int strategy) { compressionStrategy_ = strategy; }
    // 0 = sequential single-threaded; N = up to N threads for entry compression.
    void setParallelWorkers(std::size_t workers) { workers_ = workers; }
    // Test seam: forces the ZIP64 write path (large-entry layout, per-record
    // ZIP64 extra fields, and the EOCD64 record plus locator) even for small
    // archives, so the code path can be exercised without 4 GB payloads.
    void setForceZip64(bool force) { forceZip64_ = force; }

private:
    struct Entry {
        std::string data;
        std::filesystem::path sourcePath;
        bool fromFile{false};
        bool compress{true};
    };
    std::map<std::string, Entry> entries_;
    int compressionLevel_{-1};
    int compressionStrategy_{0};
    std::size_t workers_{0};
    bool forceZip64_{false};
};
}
