#pragma once
#include "MappedFile.h"
#include <array>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <functional>
#include <map>
#include <memory>
#include <string>
#include <vector>

struct z_stream_s;

namespace xlpp::internal {

struct ZipEntryInfo {
    std::string name;
    std::uint32_t localOffset{0};
    std::uint32_t compressedSize{0};
    std::uint32_t uncompressedSize{0};
    std::uint16_t method{8};
    std::uint32_t crc{0};
};

// Pull-based decompression source for a single ZIP entry. Uses memory-mapped
// data when available, otherwise falls back to ifstream.
class ZipEntrySource {
public:
    ZipEntrySource() = default;
    ZipEntrySource(std::filesystem::path path, ZipEntryInfo info);
    // Mapped variant: reads from a MappedFile view instead of opening a file handle.
    ZipEntrySource(const MappedFile* mapped, ZipEntryInfo info);
    ZipEntrySource(const ZipEntrySource&) = delete;
    ZipEntrySource& operator=(const ZipEntrySource&) = delete;
    ZipEntrySource(ZipEntrySource&&) noexcept = default;
    ZipEntrySource& operator=(ZipEntrySource&&) noexcept = default;
    ~ZipEntrySource();

    std::size_t read(unsigned char* out, std::size_t capacity);
    bool complete() const noexcept { return complete_; }

private:
    void open();
    bool refillInput();
    std::size_t readStored(unsigned char* out, std::size_t capacity);
    std::size_t readDeflate(unsigned char* out, std::size_t capacity);

    std::filesystem::path path_;
    ZipEntryInfo info_;
    std::ifstream in_;
    const MappedFile* mapped_{nullptr};
    const char* mappedPtr_{nullptr};
    std::array<unsigned char, 64 * 1024> inbuf_{};
    std::uint32_t crc_{0};
    std::uint32_t remainingCompressed_{0};
    bool opened_{false};
    bool eof_{false};
    bool complete_{false};
    z_stream_s* z_{nullptr};
};

// Direct streaming ZIP reader. Uses memory-mapped I/O for zero-copy central
// directory parsing and fast entry access. Falls back to ifstream for small or
// un-mappable files.
class ZipArchiveReader {
public:
    explicit ZipArchiveReader(const std::filesystem::path& path);

    bool contains(const std::string& name) const;
    std::size_t entryCount() const noexcept { return entries_.size(); }
    std::vector<std::string> names() const;

    std::string readEntry(const std::string& name) const;
    void forEachChunk(const std::string& name,
                      const std::function<void(const char*, std::size_t)>& emit) const;
    ZipEntrySource openEntry(const std::string& name) const;

private:
    std::filesystem::path path_;
    std::map<std::string, ZipEntryInfo> entries_;
    std::shared_ptr<MappedFile> mapped_;
};

} // namespace xlpp::internal
