#include "ZipArchive.h"
#include "Core/Threading/ThreadPool.h"
#include <zlib.h>
#include <algorithm>
#include <array>
#include <atomic>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <iterator>
#include <limits>
#include <stdexcept>
#include <vector>

namespace {
using Bytes = std::vector<unsigned char>;
template<class T> T readLE(const Bytes& b, std::size_t p) {
    std::uint64_t value = 0;
    for (std::size_t i = 0; i < sizeof(T); ++i)
        value |= static_cast<std::uint64_t>(b.at(p + i)) << (8u * i);
    return static_cast<T>(value);
}
template<class T> void writeLE(std::ostream& out,T v){for(size_t i=0;i<sizeof(T);++i){const auto c=static_cast<char>((v>>(8*i))&0xff);out.write(&c,1);}}
void writeLe64String(std::string& s, std::uint64_t v) {
    for (std::size_t i = 0; i < 8; ++i)
        s.push_back(static_cast<char>((v >> (8u * i)) & 0xffu));
}

constexpr std::uint32_t kLocalSig=0x04034b50u;
constexpr std::uint32_t kCentralSig=0x02014b50u;
constexpr std::uint32_t kEocdSig=0x06054b50u;
constexpr std::uint32_t kEocd64Sig=0x06064b50u;
constexpr std::uint32_t kEocd64LocSig=0x07064b50u;
constexpr std::uint16_t kZip64Extra=0x0001u;

struct StreamResult { std::uint32_t crc{0}; std::uint64_t compressed{0}; std::uint64_t uncompressed{0}; };

std::string readFile(const std::filesystem::path& path) {
    std::ifstream source(path, std::ios::binary);
    if (!source) throw std::runtime_error("Cannot open ZIP source file: " + path.string());
    std::string data((std::istreambuf_iterator<char>(source)), {});
    return data;
}

std::uint64_t fileSize(const std::filesystem::path& path) {
    std::error_code ec;
    const auto size = std::filesystem::file_size(path, ec);
    if (ec) throw std::runtime_error("Cannot stat ZIP source file: " + path.string());
    return static_cast<std::uint64_t>(size);
}

template<class Reader>
StreamResult streamDeflate(std::ostream& out, Reader&& reader, bool compress, int level, int strategy) {
    std::array<unsigned char, 64 * 1024> input{};
    std::array<unsigned char, 64 * 1024> output{};
    StreamResult result;
    result.crc = static_cast<std::uint32_t>(crc32(0, Z_NULL, 0));
    z_stream z{};
    if (compress && deflateInit2(&z, level, Z_DEFLATED, -MAX_WBITS, 8, strategy) != Z_OK)
        throw std::runtime_error("deflateInit2 failed");
    for (;;) {
        const std::size_t count = reader(input.data(), input.size());
        if (count) {
            result.crc = static_cast<std::uint32_t>(crc32(result.crc, input.data(), static_cast<uInt>(count)));
            result.uncompressed += static_cast<std::uint64_t>(count);
        }
        if (!compress) {
            if (count) out.write(reinterpret_cast<const char*>(input.data()), static_cast<std::streamsize>(count));
            result.compressed += static_cast<std::uint64_t>(count);
            if (!count) break;
            continue;
        }
        z.next_in = input.data();
        z.avail_in = static_cast<uInt>(count);
        const int flush = count ? Z_NO_FLUSH : Z_FINISH;
        do {
            z.next_out = output.data();
            z.avail_out = static_cast<uInt>(output.size());
            const int rc = deflate(&z, flush);
            if (rc == Z_STREAM_ERROR) { deflateEnd(&z); throw std::runtime_error("deflate failed"); }
            const std::size_t produced = output.size() - z.avail_out;
            if (produced) out.write(reinterpret_cast<const char*>(output.data()), static_cast<std::streamsize>(produced));
            result.compressed += static_cast<std::uint64_t>(produced);
        } while (z.avail_out == 0);
        if (!count) break;
    }
    if (compress) deflateEnd(&z);
    return result;
}

std::string inflateRaw(const unsigned char* p, std::size_t n, std::size_t expected) {
    std::string output;
    output.reserve(expected);
    std::array<unsigned char, 64 * 1024> buffer{};
    z_stream z{};
    if (inflateInit2(&z, -MAX_WBITS) != Z_OK) throw std::runtime_error("inflateInit2 failed");
    std::size_t inputPos = 0;
    int status = Z_OK;
    try {
        while (status != Z_STREAM_END) {
            if (z.avail_in == 0 && inputPos < n) {
                const auto chunk = std::min<std::size_t>(n - inputPos, std::numeric_limits<uInt>::max());
                z.next_in = const_cast<Bytef*>(reinterpret_cast<const Bytef*>(p + inputPos));
                z.avail_in = static_cast<uInt>(chunk);
                inputPos += chunk;
            }
            z.next_out = buffer.data();
            z.avail_out = static_cast<uInt>(buffer.size());
            status = inflate(&z, Z_NO_FLUSH);
            const auto produced = buffer.size() - z.avail_out;
            if (produced) {
                if (output.size() > expected || produced > expected - output.size())
                    throw std::runtime_error("ZIP entry expands beyond declared uncompressed size");
                output.append(reinterpret_cast<const char*>(buffer.data()), produced);
            }
            if (status != Z_OK && status != Z_STREAM_END && status != Z_BUF_ERROR)
                throw std::runtime_error("inflate failed");
            if (status == Z_BUF_ERROR && produced == 0 && z.avail_in == 0 && inputPos == n)
                throw std::runtime_error("Truncated deflate ZIP entry");
        }
    } catch (...) {
        inflateEnd(&z);
        throw;
    }
    inflateEnd(&z);
    if (output.size() != expected) throw std::runtime_error("ZIP uncompressed size mismatch");
    return output;
}

struct MemoryBlob { std::string data; std::uint32_t crc{0}; std::uint64_t compressed{0}; std::uint64_t uncompressed{0}; };

std::uint32_t crcOf(const std::string& data) {
    auto crc = crc32(0, Z_NULL, 0);
    std::size_t position = 0;
    while (position < data.size()) {
        const auto chunk = std::min<std::size_t>(data.size() - position, std::numeric_limits<uInt>::max());
        crc = crc32(crc, reinterpret_cast<const Bytef*>(data.data() + position), static_cast<uInt>(chunk));
        position += chunk;
    }
    return static_cast<std::uint32_t>(crc);
}

MemoryBlob deflateMemory(const std::string& input, int level, int strategy) {
    MemoryBlob result;
    result.crc = crcOf(input);
    result.uncompressed = static_cast<std::uint64_t>(input.size());
    std::string output;
    output.reserve(input.size() / 2 + 64);
    std::array<unsigned char, 64 * 1024> in{};
    std::array<unsigned char, 64 * 1024> out{};
    z_stream z{};
    if (deflateInit2(&z, level, Z_DEFLATED, -MAX_WBITS, 8, strategy) != Z_OK)
        throw std::runtime_error("deflateInit2 failed");
    std::size_t position = 0;
    for (;;) {
        const std::size_t remaining = input.size() - position;
        const std::size_t count = std::min(remaining, in.size());
        if (count) {
            std::memcpy(in.data(), input.data() + position, count);
            position += count;
        }
        z.next_in = in.data();
        z.avail_in = static_cast<uInt>(count);
        const int flush = count ? Z_NO_FLUSH : Z_FINISH;
        do {
            z.next_out = out.data();
            z.avail_out = static_cast<uInt>(out.size());
            const int rc = deflate(&z, flush);
            if (rc == Z_STREAM_ERROR) { deflateEnd(&z); throw std::runtime_error("deflate failed"); }
            const std::size_t produced = out.size() - z.avail_out;
            output.append(reinterpret_cast<const char*>(out.data()), produced);
        } while (z.avail_out == 0);
        if (!count) break;
    }
    deflateEnd(&z);
    result.compressed = static_cast<std::uint64_t>(output.size());
    result.data = std::move(output);
    return result;
}

struct PlannedEntry {
    std::string name;
    std::uint16_t method{0};
    std::uint32_t crc{0};
    std::uint64_t compressed{0};
    std::uint64_t uncompressed{0};
    std::string data;
    std::filesystem::path backingPath;
    bool fileBacked{false};
};

class ScopedTempFiles {
public:
    ~ScopedTempFiles() {
        for (const auto& path : paths_) {
            std::error_code ec;
            std::filesystem::remove(path, ec);
        }
    }
    void add(std::filesystem::path path) { paths_.push_back(std::move(path)); }
private:
    std::vector<std::filesystem::path> paths_;
};

std::filesystem::path makeZipPreparationPath(const std::filesystem::path& destination) {
    static std::atomic<std::uint64_t> sequence{0};
    const auto directory = destination.parent_path().empty() ? std::filesystem::current_path() : destination.parent_path();
    const auto base = destination.filename().string();
    for (unsigned attempt = 0; attempt < 1000; ++attempt) {
        const auto token = sequence.fetch_add(1, std::memory_order_relaxed) + 1;
        auto candidate = directory / (base + ".xlpp-zipprep-" + std::to_string(token) + ".tmp");
        std::error_code ec;
        if (!std::filesystem::exists(candidate, ec)) return candidate;
    }
    throw std::runtime_error("Cannot allocate temporary ZIP preparation file");
}

StreamResult scanStoredFile(const std::filesystem::path& path, const std::function<bool()>& cancel) {
    std::ifstream source(path, std::ios::binary);
    if (!source) throw std::runtime_error("Cannot open ZIP source file: " + path.string());
    std::array<unsigned char, 64 * 1024> buffer{};
    StreamResult result;
    result.crc = static_cast<std::uint32_t>(crc32(0, Z_NULL, 0));
    while (source) {
        if (cancel && cancel()) throw std::runtime_error("Save cancelled");
        source.read(reinterpret_cast<char*>(buffer.data()), static_cast<std::streamsize>(buffer.size()));
        const auto count = static_cast<std::size_t>(source.gcount());
        if (!count) break;
        result.crc = static_cast<std::uint32_t>(crc32(result.crc, buffer.data(), static_cast<uInt>(count)));
        result.uncompressed += static_cast<std::uint64_t>(count);
    }
    if (source.bad()) throw std::runtime_error("Failed reading ZIP source file: " + path.string());
    result.compressed = result.uncompressed;
    return result;
}

void copyFileToStream(std::ostream& out, const std::filesystem::path& path,
                      const std::function<bool()>& cancel) {
    std::ifstream source(path, std::ios::binary);
    if (!source) throw std::runtime_error("Cannot open prepared ZIP data: " + path.string());
    std::array<char, 64 * 1024> buffer{};
    while (source) {
        if (cancel && cancel()) throw std::runtime_error("Save cancelled");
        source.read(buffer.data(), static_cast<std::streamsize>(buffer.size()));
        const auto count = source.gcount();
        if (count <= 0) break;
        out.write(buffer.data(), count);
        if (!out) throw std::runtime_error("Failed writing ZIP entry data");
    }
    if (source.bad()) throw std::runtime_error("Failed reading prepared ZIP data: " + path.string());
}
}


namespace xlpp::internal {

void ZipArchive::add(std::string name, std::string data, bool compress){
    if (entries_.contains(name)) throw std::invalid_argument("Duplicate ZIP entry: " + name);
    entries_.emplace(std::move(name), Entry{std::move(data), {}, false, compress});
}
void ZipArchive::addFile(std::string name, std::filesystem::path sourcePath, bool compress){
    if (entries_.contains(name)) throw std::invalid_argument("Duplicate ZIP entry: " + name);
    entries_.emplace(std::move(name), Entry{{}, std::move(sourcePath), true, compress});
}
void ZipArchive::addUnique(std::string name, std::string data, bool compress){
    if (!entries_.contains(name)) add(std::move(name), std::move(data), compress);
}
void ZipArchive::replace(std::string name, std::string data, bool compress){ entries_[std::move(name)] = Entry{std::move(data), {}, false, compress}; }
bool ZipArchive::contains(const std::string& n)const{return entries_.contains(n);}
const std::string& ZipArchive::get(const std::string& n)const{auto i=entries_.find(n);if(i==entries_.end())throw std::runtime_error("ZIP entry not found: "+n);if(i->second.fromFile)throw std::runtime_error("ZIP file-backed entry is not materialized: "+n);return i->second.data;}
std::vector<std::string> ZipArchive::entryNames() const { std::vector<std::string> names; names.reserve(entries_.size()); for (const auto& [name, entry] : entries_) names.push_back(name); return names; }

void writeZip64Extra(std::string& extra, std::uint64_t us, bool hasUs,
                     std::uint64_t cs, bool hasCs, std::uint64_t off, bool hasOff) {
    if (!hasUs && !hasCs && !hasOff) return;
    std::string data;
    if (hasUs) writeLe64String(data, us);
    if (hasCs) writeLe64String(data, cs);
    if (hasOff) writeLe64String(data, off);
    extra.push_back(static_cast<char>(kZip64Extra & 0xff));
    extra.push_back(static_cast<char>((kZip64Extra >> 8) & 0xff));
    extra.push_back(static_cast<char>(data.size() & 0xff));
    extra.push_back(static_cast<char>((data.size() >> 8) & 0xff));
    extra += data;
}

void ZipArchive::save(const std::filesystem::path& path) const { save(path, ZipWriteOptions{}); }

void ZipArchive::save(const std::filesystem::path& path, const ZipWriteOptions& options) const {
    std::uint64_t totalUncompressed = 0;
    bool large = forceZip64_ || entries_.size() > 0xFFFFu;
    for (const auto& [name, entry] : entries_) {
        const std::uint64_t size = entry.fromFile ? fileSize(entry.sourcePath) : entry.data.size();
        totalUncompressed += size;
        if (size > 0xFFFFFFFFull) large = true;
    }
    if (totalUncompressed > 0xE0000000ull) large = true;

    if (large) {
        // ZIP64 requires sizes and CRC values in the local headers.  Memory
        // entries can be planned in memory, but file-backed entries must not
        // be materialized merely because the archive crosses a ZIP64 limit.
        // Compressed file-backed entries are therefore deflated once into a
        // bounded temporary backing file; stored entries are scanned for CRC
        // and copied directly from their source during the final write.
        ScopedTempFiles temporaryFiles;
        std::vector<PlannedEntry> planned;
        planned.reserve(entries_.size());
        for (const auto& [name, entry] : entries_) {
            if (options.cancel && options.cancel()) throw std::runtime_error("Save cancelled");
            PlannedEntry p;
            p.name = name;
            p.method = static_cast<std::uint16_t>(entry.compress ? 8 : 0);
            if (!entry.fromFile) {
                MemoryBlob produced;
                if (entry.compress) produced = deflateMemory(entry.data, compressionLevel_, compressionStrategy_);
                else {
                    produced.crc = crcOf(entry.data);
                    produced.data = entry.data;
                    produced.compressed = produced.uncompressed = static_cast<std::uint64_t>(entry.data.size());
                }
                p.crc = produced.crc;
                p.compressed = produced.compressed;
                p.uncompressed = produced.uncompressed;
                p.data = std::move(produced.data);
            } else if (entry.compress) {
                const auto preparedPath = makeZipPreparationPath(path);
                temporaryFiles.add(preparedPath);
                std::ifstream source(entry.sourcePath, std::ios::binary);
                if (!source) throw std::runtime_error("Cannot open ZIP source file: " + entry.sourcePath.string());
                std::ofstream prepared(preparedPath, std::ios::binary | std::ios::trunc);
                if (!prepared) throw std::runtime_error("Cannot create ZIP preparation file: " + preparedPath.string());
                const auto produced = streamDeflate(
                    prepared,
                    [&](unsigned char* buffer, std::size_t capacity) {
                        if (options.cancel && options.cancel()) throw std::runtime_error("Save cancelled");
                        source.read(reinterpret_cast<char*>(buffer), static_cast<std::streamsize>(capacity));
                        return static_cast<std::size_t>(source.gcount());
                    },
                    true, compressionLevel_, compressionStrategy_);
                if (source.bad()) throw std::runtime_error("Failed reading ZIP source file: " + entry.sourcePath.string());
                prepared.flush();
                if (!prepared) throw std::runtime_error("Failed writing ZIP preparation file: " + preparedPath.string());
                p.crc = produced.crc;
                p.compressed = produced.compressed;
                p.uncompressed = produced.uncompressed;
                p.backingPath = preparedPath;
                p.fileBacked = true;
            } else {
                const auto produced = scanStoredFile(entry.sourcePath, options.cancel);
                p.crc = produced.crc;
                p.compressed = produced.compressed;
                p.uncompressed = produced.uncompressed;
                p.backingPath = entry.sourcePath;
                p.fileBacked = true;
            }
            planned.push_back(std::move(p));
        }

        std::ofstream out(path, std::ios::binary | std::ios::trunc);
        if (!out) throw std::runtime_error("Cannot create XLSX");
        struct Rec { std::uint64_t localOffset; PlannedEntry* entry; bool off64; bool size64; };
        std::vector<Rec> records;
        records.reserve(planned.size());
        std::uint64_t offset = 0;
        const std::size_t total = planned.size();
        for (std::size_t i = 0; i < planned.size(); ++i) {
            if (options.cancel && options.cancel()) throw std::runtime_error("Save cancelled");
            auto& p = planned[i];
            const bool size64 = forceZip64_ || p.compressed > 0xFFFFFFFFull || p.uncompressed > 0xFFFFFFFFull;
            std::string extra;
            writeZip64Extra(extra, p.uncompressed, size64, p.compressed, size64, 0, false);
            writeLE(out, kLocalSig); writeLE<std::uint16_t>(out, 45); writeLE<std::uint16_t>(out, 0x0000);
            writeLE(out, p.method); writeLE<std::uint16_t>(out, 0); writeLE<std::uint16_t>(out, 0);
            writeLE(out, p.crc);
            writeLE(out, static_cast<std::uint32_t>(size64 ? 0xFFFFFFFFu : p.compressed));
            writeLE(out, static_cast<std::uint32_t>(size64 ? 0xFFFFFFFFu : p.uncompressed));
            writeLE<std::uint16_t>(out, static_cast<std::uint16_t>(p.name.size()));
            writeLE<std::uint16_t>(out, static_cast<std::uint16_t>(extra.size()));
            out.write(p.name.data(), static_cast<std::streamsize>(p.name.size()));
            out.write(extra.data(), static_cast<std::streamsize>(extra.size()));
            if (p.fileBacked) copyFileToStream(out, p.backingPath, options.cancel);
            else if (!p.data.empty()) out.write(p.data.data(), static_cast<std::streamsize>(p.data.size()));
            if (!out) throw std::runtime_error("Failed writing ZIP64 entry");
            records.push_back({offset, &p, false, size64});
            offset += 30 + p.name.size() + extra.size() + p.compressed;
            if (options.progress) options.progress(i + 1, total);
        }
        const std::uint64_t centralOffset = offset;
        std::uint64_t centralSize = 0;
        for (auto& r : records) {
            const bool off64 = forceZip64_ || r.localOffset > 0xFFFFFFFFull;
            r.off64 = off64;
            const bool size64 = r.size64;
            std::string extra;
            writeZip64Extra(extra, r.entry->uncompressed, size64, r.entry->compressed, size64, r.localOffset, off64);
            writeLE(out, kCentralSig); writeLE<std::uint16_t>(out, 45); writeLE<std::uint16_t>(out, 45); writeLE<std::uint16_t>(out, 0x0000);
            writeLE(out, r.entry->method); writeLE<std::uint16_t>(out, 0); writeLE<std::uint16_t>(out, 0);
            writeLE(out, r.entry->crc);
            writeLE(out, static_cast<std::uint32_t>(size64 ? 0xFFFFFFFFu : r.entry->compressed));
            writeLE(out, static_cast<std::uint32_t>(size64 ? 0xFFFFFFFFu : r.entry->uncompressed));
            writeLE<std::uint16_t>(out, static_cast<std::uint16_t>(r.entry->name.size()));
            writeLE<std::uint16_t>(out, static_cast<std::uint16_t>(extra.size()));
            writeLE<std::uint16_t>(out, 0); writeLE<std::uint16_t>(out, 0);
            writeLE<std::uint16_t>(out, 0); writeLE<std::uint32_t>(out, 0);
            writeLE(out, static_cast<std::uint32_t>(off64 ? 0xFFFFFFFFu : r.localOffset));
            out.write(r.entry->name.data(), static_cast<std::streamsize>(r.entry->name.size()));
            out.write(extra.data(), static_cast<std::streamsize>(extra.size()));
            centralSize += 46 + r.entry->name.size() + extra.size();
        }
        const std::uint64_t centralEnd = centralOffset + centralSize;
        const bool eocd64 = forceZip64_ || entries_.size() > 0xFFFFu || centralSize > 0xFFFFFFFFull || centralOffset > 0xFFFFFFFFull;
        if (eocd64) {
            const std::uint64_t eocd64Pos = centralEnd;
            writeLE(out, kEocd64Sig); writeLE(out, static_cast<std::uint64_t>(44));
            writeLE<std::uint16_t>(out, 45); writeLE<std::uint16_t>(out, 45);
            writeLE<std::uint32_t>(out, 0); writeLE<std::uint32_t>(out, 0);
            writeLE(out, static_cast<std::uint64_t>(records.size()));
            writeLE(out, static_cast<std::uint64_t>(records.size()));
            writeLE(out, centralSize); writeLE(out, centralOffset);
            writeLE(out, kEocd64LocSig); writeLE<std::uint32_t>(out, 0); writeLE(out, eocd64Pos); writeLE<std::uint32_t>(out, 1);
        }
        writeLE(out, kEocdSig); writeLE<std::uint16_t>(out, 0); writeLE<std::uint16_t>(out, 0);
        writeLE<std::uint16_t>(out, static_cast<std::uint16_t>(eocd64 ? 0xFFFFu : records.size()));
        writeLE<std::uint16_t>(out, static_cast<std::uint16_t>(eocd64 ? 0xFFFFu : records.size()));
        writeLE(out, static_cast<std::uint32_t>(eocd64 ? 0xFFFFFFFFu : centralSize));
        writeLE(out, static_cast<std::uint32_t>(eocd64 ? 0xFFFFFFFFu : centralOffset));
        writeLE<std::uint16_t>(out, 0);
        out.flush();
        if (!out) throw std::runtime_error("Failed finalizing ZIP64 archive");
        return;
    }

    struct Rec{std::string name;std::uint32_t crc,compressed,uncompressed,offset;std::uint16_t method;};
    std::vector<Rec> records;
    records.reserve(entries_.size());
    std::ofstream out(path, std::ios::binary);
    if(!out) throw std::runtime_error("Cannot create XLSX");

    const bool parallel = workers_ > 1 && entries_.size() > 1;
    if (parallel) {
        struct Blob { std::string name; std::uint32_t crc, compressed, uncompressed; std::uint16_t method; std::string data; };
        std::vector<std::string> order;
        order.reserve(entries_.size());
        for (const auto& [name, entry] : entries_) order.push_back(name);
        std::vector<Blob> blobs(order.size());
        {
            ThreadPool pool(std::min(workers_, entries_.size()));
            pool.parallelFor(0, entries_.size(), [&](std::size_t i) {
                const auto& name = order[i];
                const auto& entry = entries_.at(name);
                const std::string in = entry.fromFile ? readFile(entry.sourcePath) : entry.data;
                MemoryBlob produced;
                if (entry.compress) produced = deflateMemory(in, compressionLevel_, compressionStrategy_);
                else { produced.crc = crcOf(in); produced.data = in; produced.compressed = produced.uncompressed = static_cast<std::uint64_t>(in.size()); }
                blobs[i] = Blob{name, produced.crc, static_cast<std::uint32_t>(produced.compressed), static_cast<std::uint32_t>(produced.uncompressed),
                                static_cast<std::uint16_t>(entry.compress ? 8 : 0), std::move(produced.data)};
            });
        }
        const std::size_t total = blobs.size();
        for (std::size_t i = 0; i < blobs.size(); ++i) {
            if (options.cancel && options.cancel()) throw std::runtime_error("Save cancelled");
            auto& blob = blobs[i];
            const auto offset = static_cast<std::uint32_t>(out.tellp());
            writeLE(out,0x04034b50u); writeLE<std::uint16_t>(out,20); writeLE<std::uint16_t>(out,0x0008);
            writeLE(out,blob.method); writeLE<std::uint16_t>(out,0); writeLE<std::uint16_t>(out,0);
            writeLE<std::uint32_t>(out,0); writeLE<std::uint32_t>(out,0); writeLE<std::uint32_t>(out,0);
            writeLE<std::uint16_t>(out,static_cast<std::uint16_t>(blob.name.size())); writeLE<std::uint16_t>(out,0); out.write(blob.name.data(),static_cast<std::streamsize>(blob.name.size()));
            if (!blob.data.empty()) out.write(blob.data.data(), static_cast<std::streamsize>(blob.data.size()));
            writeLE(out,0x08074b50u); writeLE(out,blob.crc); writeLE(out,blob.compressed); writeLE(out,blob.uncompressed);
            records.push_back({blob.name,blob.crc,blob.compressed,blob.uncompressed,offset,blob.method});
            if (options.progress) options.progress(i + 1, total);
        }
    } else {
        const std::size_t total = entries_.size();
        std::size_t done = 0;
        for (const auto& [name, entry] : entries_) {
            if (options.cancel && options.cancel()) throw std::runtime_error("Save cancelled");
            const auto offset = static_cast<std::uint32_t>(out.tellp());
            const std::uint16_t method = entry.compress ? 8 : 0;
            writeLE(out,0x04034b50u); writeLE<std::uint16_t>(out,20); writeLE<std::uint16_t>(out,0x0008);
            writeLE(out,method); writeLE<std::uint16_t>(out,0); writeLE<std::uint16_t>(out,0);
            writeLE<std::uint32_t>(out,0); writeLE<std::uint32_t>(out,0); writeLE<std::uint32_t>(out,0);
            writeLE<std::uint16_t>(out,static_cast<std::uint16_t>(name.size())); writeLE<std::uint16_t>(out,0); out.write(name.data(),static_cast<std::streamsize>(name.size()));

            StreamResult result;
            if (entry.fromFile) {
                std::ifstream source(entry.sourcePath, std::ios::binary);
                if (!source) throw std::runtime_error("Cannot open ZIP source file: " + entry.sourcePath.string());
                result = streamDeflate(out,[&](unsigned char* buffer,std::size_t capacity){source.read(reinterpret_cast<char*>(buffer),static_cast<std::streamsize>(capacity));return static_cast<std::size_t>(source.gcount());},entry.compress,compressionLevel_,compressionStrategy_);
            } else {
                std::size_t position=0;
                result = streamDeflate(out,[&](unsigned char* buffer,std::size_t capacity){const auto remaining=entry.data.size()-position;const auto count=std::min(capacity,remaining);if(count){std::memcpy(buffer,entry.data.data()+position,count);position+=count;}return count;},entry.compress,compressionLevel_,compressionStrategy_);
            }
            writeLE(out,0x08074b50u); writeLE(out,result.crc);
            writeLE(out,static_cast<std::uint32_t>(result.compressed));
            writeLE(out,static_cast<std::uint32_t>(result.uncompressed));
            records.push_back({name,result.crc,static_cast<std::uint32_t>(result.compressed),
                               static_cast<std::uint32_t>(result.uncompressed),offset,method});
            ++done;
            if (options.progress) options.progress(done, total);
        }
    }
    const auto centralOffset=static_cast<std::uint32_t>(out.tellp());
    for(const auto& r:records){writeLE(out,0x02014b50u);writeLE<std::uint16_t>(out,20);writeLE<std::uint16_t>(out,20);writeLE<std::uint16_t>(out,0x0008);writeLE(out,r.method);writeLE<std::uint16_t>(out,0);writeLE<std::uint16_t>(out,0);writeLE(out,r.crc);writeLE(out,r.compressed);writeLE(out,r.uncompressed);writeLE<std::uint16_t>(out,static_cast<std::uint16_t>(r.name.size()));writeLE<std::uint16_t>(out,0);writeLE<std::uint16_t>(out,0);writeLE<std::uint16_t>(out,0);writeLE<std::uint16_t>(out,0);writeLE<std::uint32_t>(out,0);writeLE(out,r.offset);out.write(r.name.data(),static_cast<std::streamsize>(r.name.size()));}
    const auto centralEnd=static_cast<std::uint32_t>(out.tellp());
    writeLE(out,0x06054b50u);writeLE<std::uint16_t>(out,0);writeLE<std::uint16_t>(out,0);writeLE<std::uint16_t>(out,static_cast<std::uint16_t>(records.size()));writeLE<std::uint16_t>(out,static_cast<std::uint16_t>(records.size()));writeLE(out,centralEnd-centralOffset);writeLE(out,centralOffset);writeLE<std::uint16_t>(out,0);
}

ZipArchive ZipArchive::open(const std::filesystem::path& path) { return open(path, ZipOpenLimits{}); }

ZipArchive ZipArchive::open(const std::filesystem::path& path, const ZipOpenLimits& limits) {
    std::ifstream f(path, std::ios::binary);
    if (!f) throw std::runtime_error("Cannot open XLSX");
    Bytes b((std::istreambuf_iterator<char>(f)), {});
    if (limits.maxFileBytes && b.size() > limits.maxFileBytes) throw std::runtime_error("XLSX exceeds maximum file size");
    if (b.size() < 22) throw std::runtime_error("Invalid ZIP: file too small");

    // EOCD can only live in the last 22 + 65,535 bytes. Validate its comment
    // length so a signature embedded inside a ZIP comment cannot be mistaken
    // for the real end record.
    const std::size_t lower = b.size() > 22u + 65535u ? b.size() - (22u + 65535u) : 0u;
    std::size_t e = b.size() - 22u;
    bool foundEocd = false;
    for (;;) {
        if (readLE<std::uint32_t>(b, e) == kEocdSig) {
            const auto commentLen = readLE<std::uint16_t>(b, e + 20);
            if (e + 22u + static_cast<std::size_t>(commentLen) == b.size()) {
                foundEocd = true;
                break;
            }
        }
        if (e == lower) break;
        --e;
    }
    if (!foundEocd) throw std::runtime_error("Invalid ZIP: end of central directory not found");

    const auto disk16 = readLE<std::uint16_t>(b, e + 4);
    const auto centralDisk16 = readLE<std::uint16_t>(b, e + 6);
    const auto entriesOnDisk16 = readLE<std::uint16_t>(b, e + 8);
    const auto count16 = readLE<std::uint16_t>(b, e + 10);
    const auto cdSize32 = readLE<std::uint32_t>(b, e + 12);
    const auto cdOffset32 = readLE<std::uint32_t>(b, e + 16);
    if (disk16 != 0 || centralDisk16 != 0) throw std::runtime_error("Multi-disk ZIP packages are not supported");

    std::uint64_t count = count16;
    std::uint64_t entriesOnDisk = entriesOnDisk16;
    std::uint64_t centralSize = cdSize32;
    std::uint64_t centralOffset = cdOffset32;
    const bool zip64 = count16 == 0xFFFFu || entriesOnDisk16 == 0xFFFFu ||
                       cdSize32 == 0xFFFFFFFFu || cdOffset32 == 0xFFFFFFFFu;
    if (zip64) {
        if (e < 20) throw std::runtime_error("ZIP64 locator is missing");
        const auto loc = e - 20;
        if (readLE<std::uint32_t>(b, loc) != kEocd64LocSig) throw std::runtime_error("ZIP64 locator is missing");
        const auto eocdDisk = readLE<std::uint32_t>(b, loc + 4);
        const auto eocd64Pos = readLE<std::uint64_t>(b, loc + 8);
        const auto disks = readLE<std::uint32_t>(b, loc + 16);
        if (eocdDisk != 0 || disks != 1) throw std::runtime_error("Multi-disk ZIP64 packages are not supported");
        if (eocd64Pos > b.size() || b.size() - static_cast<std::size_t>(eocd64Pos) < 56)
            throw std::runtime_error("Invalid ZIP64 end record");
        const auto z64 = static_cast<std::size_t>(eocd64Pos);
        if (readLE<std::uint32_t>(b, z64) != kEocd64Sig) throw std::runtime_error("Invalid ZIP64 end record");
        const auto recordSize = readLE<std::uint64_t>(b, z64 + 4);
        if (recordSize < 44 || recordSize > b.size() - z64 - 12u) throw std::runtime_error("Invalid ZIP64 end-record size");
        const auto disk = readLE<std::uint32_t>(b, z64 + 16);
        const auto centralDisk = readLE<std::uint32_t>(b, z64 + 20);
        entriesOnDisk = readLE<std::uint64_t>(b, z64 + 24);
        count = readLE<std::uint64_t>(b, z64 + 32);
        centralSize = readLE<std::uint64_t>(b, z64 + 40);
        centralOffset = readLE<std::uint64_t>(b, z64 + 48);
        if (disk != 0 || centralDisk != 0 || entriesOnDisk != count)
            throw std::runtime_error("Multi-disk ZIP64 packages are not supported");
    } else if (entriesOnDisk != count) {
        throw std::runtime_error("Multi-disk ZIP packages are not supported");
    }

    if (limits.maxEntries && count > limits.maxEntries) throw std::runtime_error("XLSX exceeds maximum entry count");
    if (centralOffset > b.size() || centralSize > b.size() - static_cast<std::size_t>(centralOffset))
        throw std::runtime_error("Invalid ZIP: central directory exceeds archive bounds");
    if (count > centralSize / 46u + 1u) throw std::runtime_error("Invalid ZIP: impossible central-directory entry count");
    const auto centralEnd64 = centralOffset + centralSize;

    ZipArchive z;
    std::uint64_t totalBytes = 0;
    std::uint64_t cursor = centralOffset;
    for (std::uint64_t i = 0; i < count; ++i) {
        if (limits.cancel && limits.cancel()) throw std::runtime_error("Open cancelled");
        if (cursor > centralEnd64 || centralEnd64 - cursor < 46 || cursor > std::numeric_limits<std::size_t>::max())
            throw std::runtime_error("Invalid central directory");
        const auto pos = static_cast<std::size_t>(cursor);
        if (pos > b.size() || b.size() - pos < 46 || readLE<std::uint32_t>(b, pos) != kCentralSig)
            throw std::runtime_error("Invalid central directory");

        const auto flags = readLE<std::uint16_t>(b, pos + 8);
        const auto method = readLE<std::uint16_t>(b, pos + 10);
        const auto expectedCrc = readLE<std::uint32_t>(b, pos + 16);
        const auto cs32 = readLE<std::uint32_t>(b, pos + 20);
        const auto us32 = readLE<std::uint32_t>(b, pos + 24);
        const auto nameLen = readLE<std::uint16_t>(b, pos + 28);
        const auto extraLen = readLE<std::uint16_t>(b, pos + 30);
        const auto commentLen = readLE<std::uint16_t>(b, pos + 32);
        const auto diskStart16 = readLE<std::uint16_t>(b, pos + 34);
        const auto localOffset32 = readLE<std::uint32_t>(b, pos + 42);
        if ((flags & 0x0001u) != 0) throw std::runtime_error("Traditional ZIP encryption is not supported");
        if (method != 0 && method != 8) throw std::runtime_error("Unsupported ZIP compression method");

        const std::uint64_t recordSize = 46u + static_cast<std::uint64_t>(nameLen) + extraLen + commentLen;
        if (recordSize > centralEnd64 - cursor || recordSize > b.size() - pos)
            throw std::runtime_error("Truncated ZIP central-directory record");
        const auto extraBegin = pos + 46u + static_cast<std::size_t>(nameLen);
        const auto extraEnd = extraBegin + static_cast<std::size_t>(extraLen);

        std::uint64_t cs = cs32, us = us32, localOffset = localOffset32;
        std::uint32_t diskStart = diskStart16;
        const bool needUs = us32 == 0xFFFFFFFFu;
        const bool needCs = cs32 == 0xFFFFFFFFu;
        const bool needOffset = localOffset32 == 0xFFFFFFFFu;
        const bool needDisk = diskStart16 == 0xFFFFu;
        if (needUs || needCs || needOffset || needDisk) {
            bool found = false;
            std::size_t p = extraBegin;
            while (p + 4 <= extraEnd) {
                const auto id = readLE<std::uint16_t>(b, p);
                const auto len = readLE<std::uint16_t>(b, p + 2);
                const auto fieldBegin = p + 4u;
                if (fieldBegin > extraEnd || len > extraEnd - fieldBegin)
                    throw std::runtime_error("Truncated ZIP extra field");
                if (id == kZip64Extra) {
                    std::size_t q = fieldBegin;
                    const auto fieldEnd = fieldBegin + len;
                    auto take64 = [&]() {
                        if (q + 8 > fieldEnd) throw std::runtime_error("Truncated ZIP64 extra field");
                        const auto value = readLE<std::uint64_t>(b, q); q += 8; return value;
                    };
                    auto take32 = [&]() {
                        if (q + 4 > fieldEnd) throw std::runtime_error("Truncated ZIP64 extra field");
                        const auto value = readLE<std::uint32_t>(b, q); q += 4; return value;
                    };
                    if (needUs) us = take64();
                    if (needCs) cs = take64();
                    if (needOffset) localOffset = take64();
                    if (needDisk) diskStart = take32();
                    found = true;
                    break;
                }
                p = fieldBegin + len;
            }
            if (!found) throw std::runtime_error("ZIP64 sentinel is missing the ZIP64 extra field");
        }
        if (diskStart != 0) throw std::runtime_error("Multi-disk ZIP entries are not supported");

        if (localOffset > std::numeric_limits<std::size_t>::max() || cs > std::numeric_limits<std::size_t>::max() || us > std::numeric_limits<std::size_t>::max())
            throw std::runtime_error("ZIP entry exceeds platform address space");
        const auto localPos = static_cast<std::size_t>(localOffset);
        if (localPos > b.size() || b.size() - localPos < 30 || readLE<std::uint32_t>(b, localPos) != kLocalSig)
            throw std::runtime_error("Invalid local header");
        const auto localFlags = readLE<std::uint16_t>(b, localPos + 6);
        const auto localMethod = readLE<std::uint16_t>(b, localPos + 8);
        if (localFlags != flags || localMethod != method) throw std::runtime_error("ZIP local/central header mismatch");
        const auto lNameLen = readLE<std::uint16_t>(b, localPos + 26);
        const auto lExtraLen = readLE<std::uint16_t>(b, localPos + 28);
        const auto localVariable = static_cast<std::uint64_t>(lNameLen) + lExtraLen;
        if (localVariable > b.size() - (localPos + 30u)) throw std::runtime_error("Malformed ZIP: local header exceeds archive");
        const auto dp = localPos + 30u + static_cast<std::size_t>(localVariable);
        if (static_cast<std::size_t>(cs) > b.size() - dp) throw std::runtime_error("Malformed ZIP: entry data exceeds archive");

        std::string n(reinterpret_cast<const char*>(b.data() + pos + 46u), nameLen);
        if (n.empty()) throw std::runtime_error("ZIP entry name cannot be empty");
        const std::string localName(reinterpret_cast<const char*>(b.data() + localPos + 30u), lNameLen);
        if (localName != n) throw std::runtime_error("ZIP local/central entry name mismatch: " + n);
        if (z.entries_.contains(n)) throw std::runtime_error("Duplicate ZIP entry: " + n);

        if (method != 0 && cs < us && cs <= (std::numeric_limits<std::uint64_t>::max() - 4096u) / 1032u &&
            us > cs * 1032u + 4096u)
            throw std::runtime_error("Malformed ZIP: implausible uncompressed size for " + n);
        if (limits.maxEntryBytes && us > limits.maxEntryBytes) throw std::runtime_error("XLSX entry exceeds maximum size: " + n);
        if (limits.maxTotalBytes && (us > limits.maxTotalBytes || totalBytes > limits.maxTotalBytes - us))
            throw std::runtime_error("XLSX exceeds maximum total decompressed size");
        totalBytes += us;

        std::string d = method == 0
            ? std::string(reinterpret_cast<const char*>(b.data() + dp), static_cast<std::size_t>(cs))
            : inflateRaw(b.data() + dp, static_cast<std::size_t>(cs), static_cast<std::size_t>(us));
        if (d.size() != static_cast<std::size_t>(us)) throw std::runtime_error("ZIP uncompressed size mismatch: " + n);
        if (crcOf(d) != expectedCrc) throw std::runtime_error("ZIP CRC mismatch: " + n);
        z.entries_.emplace(std::move(n), Entry{std::move(d), {}, false, true});
        cursor += recordSize;
        if (limits.progress) limits.progress(static_cast<std::size_t>(i + 1), static_cast<std::size_t>(count));
    }
    return z;
}

}


