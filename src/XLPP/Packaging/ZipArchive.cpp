#include "ZipArchive.h"
#include "AtomicFile.h"
#include "MappedFile.h"
#include "../Threading/ThreadPool.h"
#include <zlib.h>
#include <algorithm>
#include <array>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <iterator>
#include <limits>
#include <span>
#include <stdexcept>
#include <sstream>
#include <streambuf>
#include <type_traits>
#include <vector>

namespace {
template<class T>
T readLE(std::span<const unsigned char> bytes, std::size_t position) {
    static_assert(std::is_unsigned_v<T>, "ZIP little-endian reads require an unsigned integer type");
    if (position > bytes.size() || bytes.size() - position < sizeof(T))
        throw std::runtime_error("Invalid ZIP: truncated little-endian field");
    T value = 0;
    for (std::size_t i = 0; i < sizeof(T); ++i)
        value |= static_cast<T>(bytes[position + i]) << (8 * i);
    return value;
}

template<class T>
void writeLE(std::ostream& out, T value) {
    static_assert(std::is_unsigned_v<T>, "ZIP little-endian writes require an unsigned integer type");
    std::array<char, sizeof(T)> bytes{};
    for (std::size_t i = 0; i < sizeof(T); ++i)
        bytes[i] = static_cast<char>((value >> (8 * i)) & 0xffu);
    out.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
}

void writeLe64String(std::string& s,std::uint64_t v){for(int i=0;i<8;++i)s.push_back(static_cast<char>((v>>(8*i))&0xff));}

constexpr std::uint32_t kLocalSig=0x04034b50u;
constexpr std::uint32_t kCentralSig=0x02014b50u;
constexpr std::uint32_t kEocdSig=0x06054b50u;
constexpr std::uint32_t kEocd64Sig=0x06064b50u;
constexpr std::uint32_t kEocd64LocSig=0x07064b50u;
constexpr std::uint16_t kZip64Extra=0x0001u;

struct StreamResult { std::uint32_t crc{0}; std::uint32_t compressed{0}; std::uint32_t uncompressed{0}; };

std::string readFile(const std::filesystem::path& path) {
    std::ifstream source(path, std::ios::binary | std::ios::ate);
    if (!source) throw std::runtime_error("Cannot open ZIP source file: " + path.string());
    const auto end = source.tellg();
    if (end < 0) throw std::runtime_error("Cannot stat ZIP source file: " + path.string());
    const auto size = static_cast<std::uint64_t>(end);
    if (size > static_cast<std::uint64_t>(std::numeric_limits<std::size_t>::max()))
        throw std::runtime_error("ZIP source file is too large for this platform: " + path.string());
    std::string data(static_cast<std::size_t>(size), '\0');
    source.seekg(0, std::ios::beg);
    if (!data.empty()) source.read(data.data(), static_cast<std::streamsize>(data.size()));
    if (!source && !data.empty()) throw std::runtime_error("Cannot read ZIP source file: " + path.string());
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

std::uint32_t crcOf(std::string_view data) {
    const auto initial = crc32(0, Z_NULL, 0);
    if (data.empty()) return initial;
    return static_cast<std::uint32_t>(crc32_z(initial,
        reinterpret_cast<const Bytef*>(data.data()), static_cast<z_size_t>(data.size())));
}

std::string inflateRaw(const unsigned char* input, std::size_t compressedSize,
                       std::size_t expectedSize, std::uint32_t expectedCrc,
                       std::string_view entryName) {
    std::string output(expectedSize, '\0');
    z_stream z{};
    if (inflateInit2(&z, -MAX_WBITS) != Z_OK)
        throw std::runtime_error("inflateInit2 failed");

    std::size_t inputPosition = 0;
    std::size_t outputPosition = 0;
    unsigned char overflowByte = 0;
    int status = Z_OK;
    try {
        while (status != Z_STREAM_END) {
            if (z.avail_in == 0 && inputPosition < compressedSize) {
                const auto chunk = static_cast<uInt>(std::min<std::size_t>(
                    compressedSize - inputPosition, std::numeric_limits<uInt>::max()));
                z.next_in = const_cast<Bytef*>(input + inputPosition);
                z.avail_in = chunk;
                inputPosition += chunk;
            }

            const bool writingExpected = outputPosition < expectedSize;
            if (z.avail_out == 0) {
                if (writingExpected) {
                    const auto chunk = static_cast<uInt>(std::min<std::size_t>(
                        expectedSize - outputPosition, std::numeric_limits<uInt>::max()));
                    z.next_out = reinterpret_cast<Bytef*>(output.data() + outputPosition);
                    z.avail_out = chunk;
                } else {
                    z.next_out = &overflowByte;
                    z.avail_out = 1;
                }
            }

            const auto beforeOut = z.avail_out;
            status = inflate(&z, Z_NO_FLUSH);
            const auto produced = static_cast<std::size_t>(beforeOut - z.avail_out);
            if (writingExpected) {
                outputPosition += produced;
            } else if (produced != 0) {
                throw std::runtime_error("Malformed ZIP: uncompressed size mismatch for " + std::string(entryName));
            }

            if (status == Z_STREAM_END) break;
            if (status != Z_OK && status != Z_BUF_ERROR)
                throw std::runtime_error("Malformed ZIP: deflate stream is invalid for " + std::string(entryName));
            if (status == Z_BUF_ERROR && z.avail_in == 0 && inputPosition == compressedSize)
                throw std::runtime_error("Malformed ZIP: truncated deflate stream for " + std::string(entryName));
            if (z.avail_in == 0 && inputPosition == compressedSize && produced == 0)
                throw std::runtime_error("Malformed ZIP: truncated deflate stream for " + std::string(entryName));
        }

        const auto consumed = inputPosition - static_cast<std::size_t>(z.avail_in);
        if (consumed != compressedSize || outputPosition != expectedSize)
            throw std::runtime_error("Malformed ZIP: compressed/uncompressed size mismatch for " + std::string(entryName));
    } catch (...) {
        inflateEnd(&z);
        throw;
    }
    inflateEnd(&z);

    if (crcOf(output) != expectedCrc)
        throw std::runtime_error("ZIP CRC mismatch: " + std::string(entryName));
    return output;
}

struct MemoryBlob { std::string data; std::uint32_t crc{0}; std::uint32_t compressed{0}; std::uint32_t uncompressed{0}; };

MemoryBlob deflateMemory(std::string_view input, int level, int strategy) {
    if (input.size() > std::numeric_limits<std::uint32_t>::max())
        throw std::runtime_error("In-memory ZIP entry exceeds non-ZIP64 staging limit");

    MemoryBlob result;
    result.crc = crcOf(input);
    result.uncompressed = static_cast<std::uint32_t>(input.size());
    std::string output;
    output.reserve(input.size() / 2 + 64);
    std::array<unsigned char, 64 * 1024> out{};
    z_stream z{};
    if (deflateInit2(&z, level, Z_DEFLATED, -MAX_WBITS, 8, strategy) != Z_OK)
        throw std::runtime_error("deflateInit2 failed");

    std::size_t position = 0;
    try {
        while (position < input.size()) {
            const auto chunk = static_cast<uInt>(std::min<std::size_t>(
                input.size() - position, std::numeric_limits<uInt>::max()));
            z.next_in = reinterpret_cast<Bytef*>(const_cast<char*>(input.data() + position));
            z.avail_in = chunk;
            position += chunk;
            do {
                z.next_out = out.data();
                z.avail_out = static_cast<uInt>(out.size());
                const int rc = deflate(&z, Z_NO_FLUSH);
                if (rc != Z_OK) throw std::runtime_error("deflate failed");
                const std::size_t produced = out.size() - z.avail_out;
                if (produced) output.append(reinterpret_cast<const char*>(out.data()), produced);
            } while (z.avail_in != 0 || z.avail_out == 0);
        }

        int rc = Z_OK;
        do {
            z.next_out = out.data();
            z.avail_out = static_cast<uInt>(out.size());
            rc = deflate(&z, Z_FINISH);
            if (rc != Z_OK && rc != Z_STREAM_END) throw std::runtime_error("deflate finish failed");
            const std::size_t produced = out.size() - z.avail_out;
            if (produced) output.append(reinterpret_cast<const char*>(out.data()), produced);
        } while (rc != Z_STREAM_END);
    } catch (...) {
        deflateEnd(&z);
        throw;
    }
    deflateEnd(&z);

    if (output.size() > std::numeric_limits<std::uint32_t>::max())
        throw std::runtime_error("Compressed ZIP entry exceeds non-ZIP64 staging limit");
    result.compressed = static_cast<std::uint32_t>(output.size());
    result.data = std::move(output);
    return result;
}

struct InputEntry {
    std::string name;
    const std::string* borrowedData{nullptr};
    std::string ownedData;
    bool compress{true};
    std::string_view bytes() const noexcept { return borrowedData ? std::string_view(*borrowedData) : std::string_view(ownedData); }
};
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
        const auto bytes = in.bytes();
        MemoryBlob produced;
        if (in.compress) produced = deflateMemory(bytes, level, strategy);
        else {
            if (bytes.size() > std::numeric_limits<std::uint32_t>::max())
                throw std::runtime_error("Stored ZIP entry exceeds staging limit");
            produced.crc = crcOf(bytes);
            produced.data.assign(bytes.data(), bytes.size());
            produced.compressed = produced.uncompressed = static_cast<std::uint32_t>(bytes.size());
        }
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

class VectorStreamBuf final : public std::streambuf {
public:
    explicit VectorStreamBuf(std::vector<unsigned char>& buffer) : buffer_(buffer) {}

protected:
    std::streamsize xsputn(const char* data, std::streamsize count) override {
        if (count <= 0) return 0;
        const auto oldSize = buffer_.size();
        const auto add = static_cast<std::size_t>(count);
        if (add > buffer_.max_size() - oldSize)
            throw std::length_error("ZIP memory output exceeds vector capacity");
        buffer_.resize(oldSize + add);
        std::memcpy(buffer_.data() + oldSize, data, add);
        return count;
    }

    int_type overflow(int_type ch) override {
        if (traits_type::eq_int_type(ch, traits_type::eof())) return traits_type::not_eof(ch);
        buffer_.push_back(static_cast<unsigned char>(traits_type::to_char_type(ch)));
        return ch;
    }

    pos_type seekoff(off_type offset, std::ios_base::seekdir direction,
                     std::ios_base::openmode mode) override {
        if ((mode & std::ios_base::out) == 0 || direction != std::ios_base::cur || offset != 0)
            return pos_type(off_type(-1));
        return pos_type(static_cast<off_type>(buffer_.size()));
    }

private:
    std::vector<unsigned char>& buffer_;
};

std::size_t checkedSizeT(std::uint64_t value, std::string_view field) {
    if (value > static_cast<std::uint64_t>(std::numeric_limits<std::size_t>::max()))
        throw std::runtime_error("Invalid ZIP: " + std::string(field) + " exceeds platform address space");
    return static_cast<std::size_t>(value);
}

xlpp::internal::ZipArchive parseArchive(std::span<const unsigned char> bytes,
                                        const xlpp::internal::ZipOpenLimits& limits) {
    using xlpp::internal::ZipArchive;
    if (limits.maxFileBytes && bytes.size() > limits.maxFileBytes)
        throw std::runtime_error("XLSX exceeds maximum file size");
    if (bytes.size() < 22) throw std::runtime_error("Invalid ZIP: file too small");

    constexpr std::size_t maxEocdSearch = 22u + 65535u;
    const auto searchBegin = bytes.size() > maxEocdSearch ? bytes.size() - maxEocdSearch : 0u;
    std::size_t eocd = std::string_view::npos;
    for (std::size_t candidate = bytes.size() - 22;; --candidate) {
        if (readLE<std::uint32_t>(bytes, candidate) == kEocdSig) {
            const auto commentLength = readLE<std::uint16_t>(bytes, candidate + 20);
            if (candidate + 22u + commentLength == bytes.size()) {
                eocd = candidate;
                break;
            }
        }
        if (candidate == searchBegin) break;
    }
    if (eocd == std::string_view::npos)
        throw std::runtime_error("Invalid ZIP: end of central directory not found");

    const auto diskNumber = readLE<std::uint16_t>(bytes, eocd + 4);
    const auto centralDisk = readLE<std::uint16_t>(bytes, eocd + 6);
    const auto diskEntries = readLE<std::uint16_t>(bytes, eocd + 8);
    const auto count16 = readLE<std::uint16_t>(bytes, eocd + 10);
    const auto cdSize32 = readLE<std::uint32_t>(bytes, eocd + 12);
    const auto cdOffset32 = readLE<std::uint32_t>(bytes, eocd + 16);
    if (diskNumber != 0 || centralDisk != 0 || (diskEntries != count16 && count16 != 0xFFFFu))
        throw std::runtime_error("Unsupported ZIP: multi-disk archives are not supported");

    std::uint64_t count = count16;
    std::uint64_t centralSize = cdSize32;
    std::uint64_t centralOffset = cdOffset32;
    if (count16 == 0xFFFFu || cdSize32 == 0xFFFFFFFFu || cdOffset32 == 0xFFFFFFFFu) {
        if (eocd < 20) throw std::runtime_error("Invalid ZIP64: EOCD64 locator is missing");
        const auto locator = eocd - 20;
        if (readLE<std::uint32_t>(bytes, locator) != kEocd64LocSig)
            throw std::runtime_error("Invalid ZIP64: EOCD64 locator is missing");
        if (readLE<std::uint32_t>(bytes, locator + 4) != 0 || readLE<std::uint32_t>(bytes, locator + 16) != 1)
            throw std::runtime_error("Unsupported ZIP64: multi-disk archives are not supported");
        const auto eocd64Offset = readLE<std::uint64_t>(bytes, locator + 8);
        const auto eocd64 = checkedSizeT(eocd64Offset, "EOCD64 offset");
        if (eocd64 > bytes.size() || bytes.size() - eocd64 < 56 ||
            readLE<std::uint32_t>(bytes, eocd64) != kEocd64Sig)
            throw std::runtime_error("Invalid ZIP64: EOCD64 record is truncated or missing");
        const auto recordSize = readLE<std::uint64_t>(bytes, eocd64 + 4);
        if (recordSize < 44 || recordSize > bytes.size() - eocd64 - 12)
            throw std::runtime_error("Invalid ZIP64: EOCD64 record size is invalid");
        if (readLE<std::uint32_t>(bytes, eocd64 + 16) != 0 || readLE<std::uint32_t>(bytes, eocd64 + 20) != 0)
            throw std::runtime_error("Unsupported ZIP64: multi-disk archives are not supported");
        count = readLE<std::uint64_t>(bytes, eocd64 + 32);
        centralSize = readLE<std::uint64_t>(bytes, eocd64 + 40);
        centralOffset = readLE<std::uint64_t>(bytes, eocd64 + 48);
    }

    if (limits.maxEntries && count > limits.maxEntries)
        throw std::runtime_error("XLSX exceeds maximum entry count");
    const auto centralBegin = checkedSizeT(centralOffset, "central-directory offset");
    const auto centralBytes = checkedSizeT(centralSize, "central-directory size");
    if (centralBegin > bytes.size() || centralBytes > bytes.size() - centralBegin)
        throw std::runtime_error("Invalid ZIP: central directory exceeds archive bounds");
    const auto centralEnd = centralBegin + centralBytes;

    ZipArchive archive;
    std::uint64_t totalBytes = 0;
    std::size_t cursor = centralBegin;
    for (std::uint64_t index = 0; index < count; ++index) {
        if (limits.cancel && limits.cancel()) throw std::runtime_error("Open cancelled");
        if (cursor > centralEnd || centralEnd - cursor < 46 ||
            readLE<std::uint32_t>(bytes, cursor) != kCentralSig)
            throw std::runtime_error("Invalid ZIP: central-directory record is truncated or malformed");

        const auto flags = readLE<std::uint16_t>(bytes, cursor + 8);
        const auto method = readLE<std::uint16_t>(bytes, cursor + 10);
        const auto expectedCrc = readLE<std::uint32_t>(bytes, cursor + 16);
        const auto compressed32 = readLE<std::uint32_t>(bytes, cursor + 20);
        const auto uncompressed32 = readLE<std::uint32_t>(bytes, cursor + 24);
        const auto nameLength = readLE<std::uint16_t>(bytes, cursor + 28);
        const auto extraLength = readLE<std::uint16_t>(bytes, cursor + 30);
        const auto commentLength = readLE<std::uint16_t>(bytes, cursor + 32);
        const auto localOffset32 = readLE<std::uint32_t>(bytes, cursor + 42);
        const auto recordSize = static_cast<std::size_t>(46u + nameLength + extraLength + commentLength);
        if (recordSize > centralEnd - cursor)
            throw std::runtime_error("Invalid ZIP: central-directory variable fields exceed archive bounds");
        if ((flags & 0x0001u) != 0 || (flags & 0x0040u) != 0 || (flags & 0x2000u) != 0)
            throw std::runtime_error("Unsupported ZIP: encrypted entries are not supported");
        if (method != 0 && method != 8)
            throw std::runtime_error("Unsupported ZIP compression method: " + std::to_string(method));

        std::uint64_t compressed = compressed32;
        std::uint64_t uncompressed = uncompressed32;
        std::uint64_t localOffset = localOffset32;
        const bool needCompressed64 = compressed32 == 0xFFFFFFFFu;
        const bool needUncompressed64 = uncompressed32 == 0xFFFFFFFFu;
        const bool needOffset64 = localOffset32 == 0xFFFFFFFFu;
        if (needCompressed64 || needUncompressed64 || needOffset64) {
            const auto extraBegin = cursor + 46u + nameLength;
            const auto extraEnd = extraBegin + extraLength;
            bool foundZip64 = false;
            for (std::size_t extra = extraBegin; extra + 4 <= extraEnd;) {
                const auto id = readLE<std::uint16_t>(bytes, extra);
                const auto length = readLE<std::uint16_t>(bytes, extra + 2);
                const auto payload = extra + 4u;
                if (payload > extraEnd || length > extraEnd - payload)
                    throw std::runtime_error("Invalid ZIP: malformed extra field");
                if (id == kZip64Extra) {
                    std::size_t value = payload;
                    const auto valueEnd = payload + length;
                    const auto next64 = [&]() {
                        if (value > valueEnd || valueEnd - value < 8)
                            throw std::runtime_error("Invalid ZIP64: required extra-field value is missing");
                        const auto result = readLE<std::uint64_t>(bytes, value);
                        value += 8;
                        return result;
                    };
                    if (needUncompressed64) uncompressed = next64();
                    if (needCompressed64) compressed = next64();
                    if (needOffset64) localOffset = next64();
                    foundZip64 = true;
                    break;
                }
                extra = payload + length;
            }
            if (!foundZip64) throw std::runtime_error("Invalid ZIP64: ZIP64 extra field is missing");
        }

        const std::string name(reinterpret_cast<const char*>(bytes.data() + cursor + 46u), nameLength);
        if (name.empty()) throw std::runtime_error("Invalid ZIP: empty entry name");
        if (method == 0 && compressed != uncompressed)
            throw std::runtime_error("Malformed ZIP: stored entry size mismatch for " + name);
        if (method == 8 && compressed < uncompressed && uncompressed > 4096u &&
            (compressed == 0 || (uncompressed - 4096u) / 1032u > compressed))
            throw std::runtime_error("Malformed ZIP: implausible uncompressed size for " + name);
        if (limits.maxEntryBytes && uncompressed > limits.maxEntryBytes)
            throw std::runtime_error("XLSX entry exceeds maximum size: " + name);
        if (limits.maxTotalBytes && (uncompressed > limits.maxTotalBytes || totalBytes > limits.maxTotalBytes - uncompressed))
            throw std::runtime_error("XLSX exceeds maximum total decompressed size");
        totalBytes += uncompressed;

        const auto local = checkedSizeT(localOffset, "local-header offset");
        if (local > bytes.size() || bytes.size() - local < 30 || readLE<std::uint32_t>(bytes, local) != kLocalSig)
            throw std::runtime_error("Invalid ZIP: local header is missing for " + name);
        const auto localFlags = readLE<std::uint16_t>(bytes, local + 6);
        const auto localMethod = readLE<std::uint16_t>(bytes, local + 8);
        const auto localNameLength = readLE<std::uint16_t>(bytes, local + 26);
        const auto localExtraLength = readLE<std::uint16_t>(bytes, local + 28);
        if (localMethod != method || ((localFlags ^ flags) & ~0x0008u) != 0)
            throw std::runtime_error("Malformed ZIP: central/local header mismatch for " + name);
        const auto dataPosition = local + 30u + localNameLength + localExtraLength;
        if (dataPosition > bytes.size()) throw std::runtime_error("Malformed ZIP: local header exceeds archive bounds for " + name);
        if (localNameLength != nameLength ||
            std::memcmp(bytes.data() + local + 30u, name.data(), nameLength) != 0)
            throw std::runtime_error("Malformed ZIP: central/local entry name mismatch for " + name);

        const auto compressedSize = checkedSizeT(compressed, "compressed entry size");
        const auto uncompressedSize = checkedSizeT(uncompressed, "uncompressed entry size");
        if (compressedSize > bytes.size() - dataPosition)
            throw std::runtime_error("Malformed ZIP: entry data exceeds archive bounds for " + name);

        std::string data;
        if (method == 0) {
            data.assign(reinterpret_cast<const char*>(bytes.data() + dataPosition), compressedSize);
            if (crcOf(data) != expectedCrc) throw std::runtime_error("ZIP CRC mismatch: " + name);
        } else {
            data = inflateRaw(bytes.data() + dataPosition, compressedSize, uncompressedSize, expectedCrc, name);
        }
        archive.add(name, std::move(data), method == 8);
        cursor += recordSize;
        if (limits.progress) limits.progress(static_cast<std::size_t>(index + 1), checkedSizeT(count, "entry count"));
    }
    if (cursor > centralEnd)
        throw std::runtime_error("Invalid ZIP: central-directory records exceed declared size");
    return archive;
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
    atomicWriteFile(path, [&](const std::filesystem::path& temporary) {
        std::ofstream out(temporary, std::ios::binary | std::ios::trunc);
        if (!out) throw std::runtime_error("Cannot create XLSX output: " + temporary.string());
        saveToStream(out, options);
        out.flush();
        if (!out) throw std::runtime_error("Unable to flush XLSX output: " + temporary.string());
        out.close();
        if (!out) throw std::runtime_error("Unable to close XLSX output: " + temporary.string());
    });
}

std::vector<unsigned char> ZipArchive::saveToBytes() const { return saveToBytes(ZipWriteOptions{}); }

std::vector<unsigned char> ZipArchive::saveToBytes(const ZipWriteOptions& options) const {
    std::vector<unsigned char> bytes;
    bytes.reserve(1024u * 1024u);
    VectorStreamBuf buffer(bytes);
    std::ostream out(&buffer);
    saveToStream(out, options);
    out.flush();
    if (!out) throw std::runtime_error("Unable to serialize ZIP package to memory");
    return bytes;
}

void ZipArchive::saveToStream(std::ostream& out, const ZipWriteOptions& options) const {
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
            if (entry.fromFile) in.ownedData = readFile(entry.sourcePath);
            else in.borrowedData = &entry.data;
            in.compress = entry.compress;
            inputs.push_back(std::move(in));
        }
        auto planned = planEntries(inputs, compressionLevel_, compressionStrategy_, workers_);
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
                std::string fileData;
                const std::string_view input = entry.fromFile
                    ? std::string_view(fileData = readFile(entry.sourcePath))
                    : std::string_view(entry.data);
                MemoryBlob produced;
                if (entry.compress) produced = deflateMemory(input, compressionLevel_, compressionStrategy_);
                else {
                    if (input.size() > std::numeric_limits<std::uint32_t>::max())
                        throw std::runtime_error("Stored ZIP entry exceeds staging limit");
                    produced.crc = crcOf(input);
                    produced.data.assign(input.data(), input.size());
                    produced.compressed = produced.uncompressed = static_cast<std::uint32_t>(input.size());
                }
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
    try {
        MappedFile mapped(path);
        const auto view = mapped.view();
        return parseArchive(std::span<const unsigned char>(
            reinterpret_cast<const unsigned char*>(view.data()), view.size()), limits);
    } catch (const std::runtime_error& error) {
        const std::string message = error.what();
        if (message.rfind("Cannot mmap file", 0) != 0 &&
            message.rfind("Cannot create file mapping", 0) != 0 &&
            message.rfind("Cannot map view", 0) != 0)
            throw;
    }

    // Mapping can fail on constrained/network filesystems. Fall back to one
    // exact-size buffered read without the former iterator + second-vector copy.
    const auto data = readFile(path);
    return parseArchive(std::span<const unsigned char>(
        reinterpret_cast<const unsigned char*>(data.data()), data.size()), limits);
}

ZipArchive ZipArchive::open(const std::vector<unsigned char>& bytes) { return open(bytes, ZipOpenLimits{}); }

ZipArchive ZipArchive::open(const std::vector<unsigned char>& bytes, const ZipOpenLimits& limits) {
    return parseArchive(std::span<const unsigned char>(bytes.data(), bytes.size()), limits);
}

}


