#include "ZipArchive.h"
#include "../Threading/ThreadPool.h"
#include <zlib.h>
#include <algorithm>
#include <array>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <iterator>
#include <stdexcept>
#include <vector>

namespace {
using Bytes = std::vector<unsigned char>;
template<class T> T readLE(const Bytes& b,size_t p){T v=0;for(size_t i=0;i<sizeof(T);++i)v|=static_cast<T>(b.at(p+i))<<(8*i);return v;}
template<class T> void writeLE(std::ostream& out,T v){for(size_t i=0;i<sizeof(T);++i){const auto c=static_cast<char>((v>>(8*i))&0xff);out.write(&c,1);}}
void writeLe64String(std::string& s,std::uint64_t v){for(int i=0;i<8;++i)s.push_back(static_cast<char>((v>>(8*i))&0xff));}

constexpr std::uint32_t kLocalSig=0x04034b50u;
constexpr std::uint32_t kCentralSig=0x02014b50u;
constexpr std::uint32_t kEocdSig=0x06054b50u;
constexpr std::uint32_t kEocd64Sig=0x06064b50u;
constexpr std::uint32_t kEocd64LocSig=0x07064b50u;
constexpr std::uint16_t kZip64Extra=0x0001u;

struct StreamResult { std::uint32_t crc{0}; std::uint32_t compressed{0}; std::uint32_t uncompressed{0}; };

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
    result.crc = crc32(0, Z_NULL, 0);
    z_stream z{};
    if (compress && deflateInit2(&z, level, Z_DEFLATED, -MAX_WBITS, 8, strategy) != Z_OK)
        throw std::runtime_error("deflateInit2 failed");
    for (;;) {
        const std::size_t count = reader(input.data(), input.size());
        if (count) {
            result.crc = crc32(result.crc, input.data(), static_cast<uInt>(count));
            result.uncompressed += static_cast<std::uint32_t>(count);
        }
        if (!compress) {
            if (count) out.write(reinterpret_cast<const char*>(input.data()), static_cast<std::streamsize>(count));
            result.compressed += static_cast<std::uint32_t>(count);
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
            result.compressed += static_cast<std::uint32_t>(produced);
        } while (z.avail_out == 0);
        if (!count) break;
    }
    if (compress) deflateEnd(&z);
    return result;
}

std::string inflateRaw(const unsigned char* p,size_t n,size_t expected){std::string o(expected,'\0');z_stream z{};if(inflateInit2(&z,-MAX_WBITS)!=Z_OK)throw std::runtime_error("inflateInit2 failed");z.next_in=(Bytef*)p;z.avail_in=(uInt)n;z.next_out=(Bytef*)o.data();z.avail_out=(uInt)o.size();int rc=inflate(&z,Z_FINISH);inflateEnd(&z);if(rc!=Z_STREAM_END)throw std::runtime_error("inflate failed");o.resize(z.total_out);return o;}

struct MemoryBlob { std::string data; std::uint32_t crc{0}; std::uint32_t compressed{0}; std::uint32_t uncompressed{0}; };

std::uint32_t crcOf(const std::string& data) {
    if (data.empty()) return crc32(0, Z_NULL, 0);
    return crc32(crc32(0, Z_NULL, 0), reinterpret_cast<const Bytef*>(data.data()), static_cast<uInt>(data.size()));
}

MemoryBlob deflateMemory(const std::string& input, int level, int strategy) {
    MemoryBlob result;
    result.crc = crcOf(input);
    result.uncompressed = static_cast<std::uint32_t>(input.size());
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
    result.compressed = static_cast<std::uint32_t>(output.size());
    result.data = std::move(output);
    return result;
}

struct InputEntry { std::string name; std::string data; bool compress; };
struct PlannedEntry {
    std::string name;
    std::uint16_t method{0};
    std::uint32_t crc{0};
    std::uint64_t compressed{0};
    std::uint64_t uncompressed{0};
    std::string data;
};

// In-memory compression of every entry, spread across workers using ThreadPool.
std::vector<PlannedEntry> planEntries(const std::vector<InputEntry>& inputs, int level, int strategy,
                                      std::size_t workers) {
    std::vector<PlannedEntry> planned(inputs.size());
    const auto processEntry = [&](std::size_t i) {
        const auto& in = inputs[i];
        MemoryBlob produced;
        if (in.compress) produced = deflateMemory(in.data, level, strategy);
        else { produced.crc = crcOf(in.data); produced.data = in.data; produced.compressed = produced.uncompressed = static_cast<std::uint32_t>(in.data.size()); }
        auto& p = planned[i];
        p.name = in.name;
        p.method = static_cast<std::uint16_t>(in.compress ? 8 : 0);
        p.crc = produced.crc;
        p.compressed = produced.compressed;
        p.uncompressed = produced.uncompressed;
        p.data = std::move(produced.data);
    };
    if (workers > 1 && inputs.size() > 1) {
        xlpp::internal::ThreadPool pool(std::min(workers, inputs.size()));
        pool.parallelFor(0, inputs.size(), processEntry);
    } else {
        for (std::size_t i = 0; i < inputs.size(); ++i) processEntry(i);
    }
    return planned;
}
}


namespace xlpp::internal {

void ZipArchive::add(std::string name, std::string data, bool compress){ entries_[std::move(name)] = Entry{std::move(data), {}, false, compress}; }
void ZipArchive::addFile(std::string name, std::filesystem::path sourcePath, bool compress){ entries_[std::move(name)] = Entry{{}, std::move(sourcePath), true, compress}; }
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
        // ZIP64 layout: sizes and offsets are known up front, so local headers
        // carry real values (no data descriptor) and per-record extra fields.
        std::vector<InputEntry> inputs;
        inputs.reserve(entries_.size());
        for (const auto& [name, entry] : entries_) {
            InputEntry in;
            in.name = name;
            in.data = entry.fromFile ? readFile(entry.sourcePath) : entry.data;
            in.compress = entry.compress;
            inputs.push_back(std::move(in));
        }
        auto planned = planEntries(inputs, compressionLevel_, compressionStrategy_, workers_);
        std::ofstream out(path, std::ios::binary);
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
            writeLE(out, kLocalSig); writeLE<std::uint16_t>(out, 20); writeLE<std::uint16_t>(out, 0x0000);
            writeLE(out, p.method); writeLE<std::uint16_t>(out, 0); writeLE<std::uint16_t>(out, 0);
            writeLE(out, p.crc);
            writeLE(out, static_cast<std::uint32_t>(size64 ? 0xFFFFFFFFu : p.compressed));
            writeLE(out, static_cast<std::uint32_t>(size64 ? 0xFFFFFFFFu : p.uncompressed));
            writeLE<std::uint16_t>(out, static_cast<std::uint16_t>(p.name.size()));
            writeLE<std::uint16_t>(out, static_cast<std::uint16_t>(extra.size()));
            out.write(p.name.data(), static_cast<std::streamsize>(p.name.size()));
            out.write(extra.data(), static_cast<std::streamsize>(extra.size()));
            if (!p.data.empty()) out.write(p.data.data(), static_cast<std::streamsize>(p.data.size()));
            records.push_back({offset, &p, false, size64});
            offset += 30 + p.name.size() + extra.size() + p.data.size();
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
            writeLE(out, kCentralSig); writeLE<std::uint16_t>(out, 20); writeLE<std::uint16_t>(out, 20); writeLE<std::uint16_t>(out, 0x0000);
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
                else { produced.crc = crcOf(in); produced.data = in; produced.compressed = produced.uncompressed = static_cast<std::uint32_t>(in.size()); }
                blobs[i] = Blob{name, produced.crc, produced.compressed, produced.uncompressed,
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
            writeLE(out,0x08074b50u); writeLE(out,result.crc); writeLE(out,result.compressed); writeLE(out,result.uncompressed);
            records.push_back({name,result.crc,result.compressed,result.uncompressed,offset,method});
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

    size_t e = b.size() > 22 ? b.size() - 22 : 0;
    while (e > 0 && readLE<std::uint32_t>(b, e) != kEocdSig) --e;
    if (readLE<std::uint32_t>(b, e) != kEocdSig) throw std::runtime_error("Invalid ZIP: end of central directory not found");
    const std::uint16_t count16 = readLE<std::uint16_t>(b, e + 10);
    const std::uint32_t cdSize32 = readLE<std::uint32_t>(b, e + 12);
    const std::uint32_t cdOffset32 = readLE<std::uint32_t>(b, e + 16);
    std::uint64_t count = count16;
    std::uint64_t centralSize = cdSize32;
    std::uint64_t cdOffset = cdOffset32;
    if (count16 == 0xFFFF || cdSize32 == 0xFFFFFFFF || cdOffset32 == 0xFFFFFFFF) {
        if (e >= 20) {
            const size_t loc = e - 20;
            if (readLE<std::uint32_t>(b, loc) == kEocd64LocSig) {
                const std::uint64_t eocd64Pos = readLE<std::uint64_t>(b, loc + 8);
                if (eocd64Pos <= b.size() && eocd64Pos + 56 <= b.size() &&
                    readLE<std::uint32_t>(b, static_cast<size_t>(eocd64Pos)) == kEocd64Sig) {
                    count = readLE<std::uint64_t>(b, static_cast<size_t>(eocd64Pos) + 32);
                    centralSize = readLE<std::uint64_t>(b, static_cast<size_t>(eocd64Pos) + 40);
                    cdOffset = readLE<std::uint64_t>(b, static_cast<size_t>(eocd64Pos) + 48);
                }
            }
        }
    }
    if (limits.maxEntries && count > limits.maxEntries) throw std::runtime_error("XLSX exceeds maximum entry count");

    ZipArchive z;
    std::uint64_t totalBytes = 0;
    for (std::uint64_t i = 0; i < count; ++i) {
        if (limits.cancel && limits.cancel()) throw std::runtime_error("Open cancelled");
        const size_t pos = static_cast<size_t>(cdOffset);
        if (pos > b.size() || b.size() - pos < 46 || readLE<std::uint32_t>(b, pos) != kCentralSig) throw std::runtime_error("Invalid central directory");
        const std::uint16_t method = readLE<std::uint16_t>(b, pos + 10);
        const std::uint32_t cs32 = readLE<std::uint32_t>(b, pos + 20);
        const std::uint32_t us32 = readLE<std::uint32_t>(b, pos + 24);
        const std::uint16_t nameLen = readLE<std::uint16_t>(b, pos + 28);
        const std::uint16_t extraLen = readLE<std::uint16_t>(b, pos + 30);
        const std::uint16_t commentLen = readLE<std::uint16_t>(b, pos + 32);
        const std::uint32_t localOffset32 = readLE<std::uint32_t>(b, pos + 42);
        std::uint64_t cs = cs32, us = us32, localOffset = localOffset32;
        if (cs32 == 0xFFFFFFFFu || us32 == 0xFFFFFFFFu || localOffset32 == 0xFFFFFFFFu) {
            const size_t extraBegin = pos + 46 + nameLen;
            const size_t extraEnd = extraBegin + extraLen;
            size_t p = extraBegin;
            while (p + 4 <= extraEnd && extraEnd <= b.size()) {
                const std::uint16_t id = readLE<std::uint16_t>(b, p);
                const std::uint16_t len = readLE<std::uint16_t>(b, p + 2);
                if (id == kZip64Extra) {
                    size_t q = p + 4;
                    if (us32 == 0xFFFFFFFFu) { if (q + 8 > extraEnd) break; us = readLE<std::uint64_t>(b, q); q += 8; }
                    if (cs32 == 0xFFFFFFFFu) { if (q + 8 > extraEnd) break; cs = readLE<std::uint64_t>(b, q); q += 8; }
                    if (localOffset32 == 0xFFFFFFFFu) { if (q + 8 > extraEnd) break; localOffset = readLE<std::uint64_t>(b, q); q += 8; }
                    break;
                }
                p = p + 4 + len;
            }
        }
        const size_t localPos = static_cast<size_t>(localOffset);
        if (localPos > b.size() || b.size() - localPos < 30 || readLE<std::uint32_t>(b, localPos) != kLocalSig) throw std::runtime_error("Invalid local header");
        const std::uint16_t lNameLen = readLE<std::uint16_t>(b, localPos + 26);
        const std::uint16_t lExtraLen = readLE<std::uint16_t>(b, localPos + 28);
        const size_t dp = localPos + 30 + lNameLen + lExtraLen;
        if (dp > b.size() || static_cast<size_t>(cs) > b.size() - dp) throw std::runtime_error("Malformed ZIP: entry data exceeds archive");
        std::string n(reinterpret_cast<char*>(&b[pos + 46]), nameLen);
        if (method != 0 && static_cast<std::uint64_t>(cs) < static_cast<std::uint64_t>(us) &&
            static_cast<std::uint64_t>(us) > static_cast<std::uint64_t>(cs) * 1032u + 4096u)
            throw std::runtime_error("Malformed ZIP: implausible uncompressed size for " + n);
        if (limits.maxEntryBytes && us > limits.maxEntryBytes) throw std::runtime_error("XLSX entry exceeds maximum size: " + n);
        if (limits.maxTotalBytes && totalBytes + us > limits.maxTotalBytes) throw std::runtime_error("XLSX exceeds maximum total decompressed size");
        totalBytes += us;
        std::string d = method == 0 ? std::string(reinterpret_cast<char*>(&b[dp]), static_cast<size_t>(cs)) : inflateRaw(&b[dp], static_cast<size_t>(cs), static_cast<size_t>(us));
        z.entries_[std::move(n)] = Entry{std::move(d), {}, false, true};
        cdOffset += 46 + nameLen + extraLen + commentLen;
        if (limits.progress) limits.progress(static_cast<std::size_t>(i + 1), static_cast<std::size_t>(count));
    }
    return z;
}

}


