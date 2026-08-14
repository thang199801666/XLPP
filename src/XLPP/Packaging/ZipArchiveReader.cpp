#include "ZipArchiveReader.h"
#include <zlib.h>
#include <algorithm>
#include <array>
#include <cstring>
#include <limits>
#include <span>
#include <stdexcept>
#include <string_view>
#include <type_traits>
#include <utility>

namespace {
constexpr std::uint32_t kLocalSig = 0x04034b50u;
constexpr std::uint32_t kCentralSig = 0x02014b50u;
constexpr std::uint32_t kEocdSig = 0x06054b50u;
constexpr std::uint32_t kEocd64Sig = 0x06064b50u;
constexpr std::uint32_t kEocd64LocSig = 0x07064b50u;
constexpr std::uint16_t kZip64Extra = 0x0001u;

using ByteSpan = std::span<const unsigned char>;

template<class T>
T readLE(ByteSpan bytes, std::size_t position) {
    static_assert(std::is_unsigned_v<T>);
    if (position > bytes.size() || bytes.size() - position < sizeof(T))
        throw std::runtime_error("Invalid ZIP: truncated little-endian field");
    T value = 0;
    for (std::size_t i = 0; i < sizeof(T); ++i)
        value |= static_cast<T>(bytes[position + i]) << (8 * i);
    return value;
}

std::size_t checkedSize(std::uint64_t value, std::string_view field) {
    if (value > static_cast<std::uint64_t>(std::numeric_limits<std::size_t>::max()))
        throw std::runtime_error("Invalid ZIP: " + std::string(field) + " exceeds platform address space");
    return static_cast<std::size_t>(value);
}

std::vector<unsigned char> readWholeFile(const std::filesystem::path& path) {
    std::ifstream in(path, std::ios::binary | std::ios::ate);
    if (!in) throw std::runtime_error("Cannot open XLSX package: " + path.string());
    const auto end = in.tellg();
    if (end < 0) throw std::runtime_error("Cannot determine XLSX package size: " + path.string());
    const auto size64 = static_cast<std::uint64_t>(end);
    const auto size = checkedSize(size64, "file size");
    std::vector<unsigned char> bytes(size);
    in.seekg(0, std::ios::beg);
    if (size != 0) in.read(reinterpret_cast<char*>(bytes.data()), static_cast<std::streamsize>(size));
    if (!in && size != 0) throw std::runtime_error("Cannot read XLSX package: " + path.string());
    return bytes;
}

struct CentralDirectoryInfo {
    std::uint64_t count{0};
    std::uint64_t offset{0};
    std::uint64_t size{0};
};

CentralDirectoryInfo locateCentralDirectory(ByteSpan bytes) {
    if (bytes.size() < 22) throw std::runtime_error("Invalid ZIP package: file too small");

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
        throw std::runtime_error("Invalid ZIP package: end of central directory not found");

    const auto diskNumber = readLE<std::uint16_t>(bytes, eocd + 4);
    const auto centralDisk = readLE<std::uint16_t>(bytes, eocd + 6);
    const auto diskEntries = readLE<std::uint16_t>(bytes, eocd + 8);
    const auto count16 = readLE<std::uint16_t>(bytes, eocd + 10);
    const auto size32 = readLE<std::uint32_t>(bytes, eocd + 12);
    const auto offset32 = readLE<std::uint32_t>(bytes, eocd + 16);
    if (diskNumber != 0 || centralDisk != 0 || (diskEntries != count16 && count16 != 0xffffu))
        throw std::runtime_error("Unsupported ZIP: multi-disk archives are not supported");

    CentralDirectoryInfo result{count16, offset32, size32};
    if (count16 == 0xffffu || size32 == 0xffffffffu || offset32 == 0xffffffffu) {
        if (eocd < 20) throw std::runtime_error("Invalid ZIP64: EOCD64 locator is missing");
        const auto locator = eocd - 20;
        if (readLE<std::uint32_t>(bytes, locator) != kEocd64LocSig)
            throw std::runtime_error("Invalid ZIP64: EOCD64 locator is missing");
        if (readLE<std::uint32_t>(bytes, locator + 4) != 0 ||
            readLE<std::uint32_t>(bytes, locator + 16) != 1)
            throw std::runtime_error("Unsupported ZIP64: multi-disk archives are not supported");

        const auto eocd64 = checkedSize(readLE<std::uint64_t>(bytes, locator + 8), "EOCD64 offset");
        if (eocd64 > bytes.size() || bytes.size() - eocd64 < 56 ||
            readLE<std::uint32_t>(bytes, eocd64) != kEocd64Sig)
            throw std::runtime_error("Invalid ZIP64: EOCD64 record is truncated or missing");
        const auto recordSize = readLE<std::uint64_t>(bytes, eocd64 + 4);
        if (recordSize < 44 || recordSize > bytes.size() - eocd64 - 12)
            throw std::runtime_error("Invalid ZIP64: EOCD64 record size is invalid");
        if (readLE<std::uint32_t>(bytes, eocd64 + 16) != 0 ||
            readLE<std::uint32_t>(bytes, eocd64 + 20) != 0)
            throw std::runtime_error("Unsupported ZIP64: multi-disk archives are not supported");

        const auto diskCount = readLE<std::uint64_t>(bytes, eocd64 + 24);
        const auto totalCount = readLE<std::uint64_t>(bytes, eocd64 + 32);
        if (diskCount != totalCount)
            throw std::runtime_error("Unsupported ZIP64: multi-disk archives are not supported");
        result.count = totalCount;
        result.size = readLE<std::uint64_t>(bytes, eocd64 + 40);
        result.offset = readLE<std::uint64_t>(bytes, eocd64 + 48);
    }
    return result;
}

void applyZip64Extra(ByteSpan bytes, std::size_t extraBegin, std::size_t extraLength,
                     bool needUncompressed, bool needCompressed, bool needOffset,
                     std::uint64_t& uncompressed, std::uint64_t& compressed,
                     std::uint64_t& localOffset) {
    if (!needUncompressed && !needCompressed && !needOffset) return;
    const auto extraEnd = extraBegin + extraLength;
    bool found = false;
    for (std::size_t cursor = extraBegin; cursor + 4 <= extraEnd;) {
        const auto id = readLE<std::uint16_t>(bytes, cursor);
        const auto length = readLE<std::uint16_t>(bytes, cursor + 2);
        const auto payload = cursor + 4u;
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
            if (needUncompressed) uncompressed = next64();
            if (needCompressed) compressed = next64();
            if (needOffset) localOffset = next64();
            found = true;
            break;
        }
        cursor = payload + length;
    }
    if (!found) throw std::runtime_error("Invalid ZIP64: ZIP64 extra field is missing");
}

std::map<std::string, xlpp::internal::ZipEntryInfo> parseEntries(ByteSpan bytes, const xlpp::internal::ZipOpenLimits& limits) {
    if (limits.maxFileBytes && bytes.size() > limits.maxFileBytes)
        throw std::runtime_error("ZIP package exceeds configured file-size limit");
    const auto central = locateCentralDirectory(bytes);
    const auto centralBegin = checkedSize(central.offset, "central-directory offset");
    const auto centralSize = checkedSize(central.size, "central-directory size");
    if (centralBegin > bytes.size() || centralSize > bytes.size() - centralBegin)
        throw std::runtime_error("Invalid ZIP: central directory exceeds archive bounds");
    const auto centralEnd = centralBegin + centralSize;

    if (limits.maxEntries && central.count > limits.maxEntries)
        throw std::runtime_error("ZIP package exceeds configured entry-count limit");
    std::map<std::string, xlpp::internal::ZipEntryInfo> entries;
    std::uint64_t totalUncompressed = 0;
    std::size_t cursor = centralBegin;
    for (std::uint64_t index = 0; index < central.count; ++index) {
        if (cursor > centralEnd || centralEnd - cursor < 46 ||
            readLE<std::uint32_t>(bytes, cursor) != kCentralSig)
            throw std::runtime_error("Invalid ZIP central directory");

        const auto flags = readLE<std::uint16_t>(bytes, cursor + 8);
        const auto method = readLE<std::uint16_t>(bytes, cursor + 10);
        const auto crc = readLE<std::uint32_t>(bytes, cursor + 16);
        const auto compressed32 = readLE<std::uint32_t>(bytes, cursor + 20);
        const auto uncompressed32 = readLE<std::uint32_t>(bytes, cursor + 24);
        const auto nameLen = readLE<std::uint16_t>(bytes, cursor + 28);
        const auto extraLen = readLE<std::uint16_t>(bytes, cursor + 30);
        const auto commentLen = readLE<std::uint16_t>(bytes, cursor + 32);
        const auto localOffset32 = readLE<std::uint32_t>(bytes, cursor + 42);
        const auto recordSize = static_cast<std::size_t>(46u + nameLen + extraLen + commentLen);
        if (recordSize > centralEnd - cursor)
            throw std::runtime_error("Invalid ZIP: central-directory variable fields exceed archive bounds");
        if ((flags & 0x0001u) != 0 || (flags & 0x0040u) != 0 || (flags & 0x2000u) != 0)
            throw std::runtime_error("Unsupported ZIP: encrypted entries are not supported");
        if (method != 0 && method != 8)
            throw std::runtime_error("Unsupported ZIP compression method: " + std::to_string(method));

        std::uint64_t compressed = compressed32;
        std::uint64_t uncompressed = uncompressed32;
        std::uint64_t localOffset = localOffset32;
        const bool needCompressed = compressed32 == 0xffffffffu;
        const bool needUncompressed = uncompressed32 == 0xffffffffu;
        const bool needOffset = localOffset32 == 0xffffffffu;
        applyZip64Extra(bytes, cursor + 46u + nameLen, extraLen,
                        needUncompressed, needCompressed, needOffset,
                        uncompressed, compressed, localOffset);

        const std::string name(reinterpret_cast<const char*>(bytes.data() + cursor + 46u), nameLen);
        if (name.empty()) throw std::runtime_error("Invalid ZIP: empty entry name");
        if (entries.contains(name)) throw std::runtime_error("Duplicate ZIP entry: " + name);
        if (limits.maxEntryBytes && uncompressed > limits.maxEntryBytes)
            throw std::runtime_error("ZIP entry exceeds configured size limit: " + name);
        if (uncompressed > std::numeric_limits<std::uint64_t>::max() - totalUncompressed)
            throw std::runtime_error("ZIP total uncompressed size overflow");
        totalUncompressed += uncompressed;
        if (limits.maxTotalBytes && totalUncompressed > limits.maxTotalBytes)
            throw std::runtime_error("ZIP package exceeds configured total-size limit");
        if (limits.cancel && limits.cancel())
            throw std::runtime_error("ZIP open cancelled");
        if (limits.progress && central.count <= static_cast<std::uint64_t>(std::numeric_limits<std::size_t>::max()))
            limits.progress(static_cast<std::size_t>(index + 1), static_cast<std::size_t>(central.count));
        if (method == 0 && compressed != uncompressed)
            throw std::runtime_error("Malformed ZIP: stored entry size mismatch for " + name);
        if (method == 8 && compressed < uncompressed && uncompressed > 4096u &&
            (compressed == 0 || (uncompressed - 4096u) / 1032u > compressed))
            throw std::runtime_error("Malformed ZIP: implausible uncompressed size for " + name);

        const auto local = checkedSize(localOffset, "local-header offset");
        if (local > bytes.size() || bytes.size() - local < 30 ||
            readLE<std::uint32_t>(bytes, local) != kLocalSig)
            throw std::runtime_error("Invalid ZIP: local header is missing for " + name);
        const auto localFlags = readLE<std::uint16_t>(bytes, local + 6);
        const auto localMethod = readLE<std::uint16_t>(bytes, local + 8);
        const auto localNameLen = readLE<std::uint16_t>(bytes, local + 26);
        const auto localExtraLen = readLE<std::uint16_t>(bytes, local + 28);
        if (localMethod != method || ((localFlags ^ flags) & ~0x0008u) != 0)
            throw std::runtime_error("Malformed ZIP: central/local header mismatch for " + name);
        const auto dataOffset64 = localOffset + 30u + localNameLen + localExtraLen;
        const auto dataOffset = checkedSize(dataOffset64, "entry data offset");
        if (dataOffset > bytes.size())
            throw std::runtime_error("Malformed ZIP: local header exceeds archive bounds for " + name);
        if (localNameLen != nameLen || local + 30u > bytes.size() ||
            nameLen > bytes.size() - (local + 30u) ||
            std::memcmp(bytes.data() + local + 30u, name.data(), nameLen) != 0)
            throw std::runtime_error("Malformed ZIP: central/local entry name mismatch for " + name);
        const auto compressedSize = checkedSize(compressed, "compressed entry size");
        if (compressedSize > bytes.size() - dataOffset)
            throw std::runtime_error("Malformed ZIP: entry data exceeds archive bounds for " + name);

        entries.emplace(name, xlpp::internal::ZipEntryInfo{
            name, localOffset, dataOffset64, compressed, uncompressed, method, flags, crc});
        cursor += recordSize;
    }
    if (cursor > centralEnd)
        throw std::runtime_error("Invalid ZIP: central-directory records exceed declared size");
    return entries;
}

std::uint32_t updateCrc(std::uint32_t crc, const unsigned char* bytes, std::size_t count) {
    return static_cast<std::uint32_t>(crc32_z(crc, bytes, static_cast<z_size_t>(count)));
}

} // namespace

namespace xlpp::internal {

ZipEntrySource::ZipEntrySource(std::filesystem::path path, ZipEntryInfo info)
    : path_(std::move(path)), info_(std::move(info)) {}

ZipEntrySource::ZipEntrySource(const MappedFile* mapped, ZipEntryInfo info)
    : info_(std::move(info)), mapped_(mapped) {}

ZipEntrySource::ZipEntrySource(ZipEntrySource&& other) noexcept
    : path_(std::move(other.path_)), info_(std::move(other.info_)), in_(std::move(other.in_)),
      mapped_(other.mapped_), mappedPtr_(other.mappedPtr_), inbuf_(other.inbuf_), crc_(other.crc_),
      remainingCompressed_(other.remainingCompressed_), produced_(other.produced_), opened_(other.opened_),
      eof_(other.eof_), complete_(other.complete_), z_(other.z_) {
    if (z_ && z_->next_in) {
        const auto inputOffset = static_cast<std::size_t>(z_->next_in - other.inbuf_.data());
        if (inputOffset <= inbuf_.size()) z_->next_in = inbuf_.data() + inputOffset;
    }
    other.mapped_ = nullptr;
    other.mappedPtr_ = nullptr;
    other.z_ = nullptr;
    other.remainingCompressed_ = 0;
    other.produced_ = 0;
    other.opened_ = false;
    other.eof_ = true;
    other.complete_ = false;
}

ZipEntrySource& ZipEntrySource::operator=(ZipEntrySource&& other) noexcept {
    if (this != &other) {
        releaseInflater();
        const auto inputOffset = (other.z_ && other.z_->next_in)
            ? static_cast<std::size_t>(other.z_->next_in - other.inbuf_.data())
            : std::size_t{0};
        const bool hasInputPointer = other.z_ && other.z_->next_in;
        path_ = std::move(other.path_);
        info_ = std::move(other.info_);
        in_ = std::move(other.in_);
        mapped_ = other.mapped_;
        mappedPtr_ = other.mappedPtr_;
        inbuf_ = other.inbuf_;
        crc_ = other.crc_;
        remainingCompressed_ = other.remainingCompressed_;
        produced_ = other.produced_;
        opened_ = other.opened_;
        eof_ = other.eof_;
        complete_ = other.complete_;
        z_ = other.z_;
        if (hasInputPointer && inputOffset <= inbuf_.size()) z_->next_in = inbuf_.data() + inputOffset;

        other.mapped_ = nullptr;
        other.mappedPtr_ = nullptr;
        other.z_ = nullptr;
        other.remainingCompressed_ = 0;
        other.produced_ = 0;
        other.opened_ = false;
        other.eof_ = true;
        other.complete_ = false;
    }
    return *this;
}

ZipEntrySource::~ZipEntrySource() { releaseInflater(); }

void ZipEntrySource::releaseInflater() noexcept {
    if (z_) {
        inflateEnd(z_);
        delete z_;
        z_ = nullptr;
    }
}

void ZipEntrySource::open() {
    if (opened_) return;
    remainingCompressed_ = info_.compressedSize;
    produced_ = 0;
    crc_ = crc32(0, Z_NULL, 0);

    if (mapped_) {
        const auto dataOffset = checkedSize(info_.dataOffset, "entry data offset");
        const auto compressed = checkedSize(info_.compressedSize, "compressed entry size");
        if (dataOffset > mapped_->size() || compressed > mapped_->size() - dataOffset)
            throw std::runtime_error("ZIP entry data exceeds mapped file: " + info_.name);
        mappedPtr_ = reinterpret_cast<const unsigned char*>(mapped_->data()) + dataOffset;
    } else {
        in_.open(path_, std::ios::binary);
        if (!in_) throw std::runtime_error("Cannot open XLSX package: " + path_.string());
        if (info_.dataOffset > static_cast<std::uint64_t>(std::numeric_limits<std::streamoff>::max()))
            throw std::runtime_error("ZIP entry offset exceeds stream range: " + info_.name);
        in_.seekg(static_cast<std::streamoff>(info_.dataOffset), std::ios::beg);
        if (!in_) throw std::runtime_error("Cannot seek to ZIP entry: " + info_.name);
    }

    if (info_.method == 8) {
        z_ = new z_stream{};
        if (inflateInit2(z_, -MAX_WBITS) != Z_OK) {
            delete z_;
            z_ = nullptr;
            throw std::runtime_error("inflateInit2 failed");
        }
    } else if (info_.method != 0) {
        throw std::runtime_error("Unsupported ZIP compression method");
    }
    opened_ = true;
}

bool ZipEntrySource::refillInput() {
    if (remainingCompressed_ == 0) return false;
    const auto count = static_cast<std::size_t>(
        std::min<std::uint64_t>(remainingCompressed_, inbuf_.size()));
    if (mappedPtr_) {
        std::memcpy(inbuf_.data(), mappedPtr_, count);
        mappedPtr_ += count;
    } else {
        in_.read(reinterpret_cast<char*>(inbuf_.data()), static_cast<std::streamsize>(count));
        if (static_cast<std::size_t>(in_.gcount()) != count)
            throw std::runtime_error("Truncated ZIP entry data: " + info_.name);
    }
    z_->next_in = inbuf_.data();
    z_->avail_in = static_cast<uInt>(count);
    remainingCompressed_ -= count;
    return true;
}

void ZipEntrySource::validateCompletion() {
    if (produced_ != info_.uncompressedSize)
        throw std::runtime_error("ZIP uncompressed-size mismatch: " + info_.name);
    if (crc_ != info_.crc)
        throw std::runtime_error("ZIP CRC mismatch: " + info_.name);
    eof_ = true;
    complete_ = true;
}

std::size_t ZipEntrySource::readStored(unsigned char* out, std::size_t capacity) {
    if (capacity == 0) return 0;
    const auto count = static_cast<std::size_t>(
        std::min<std::uint64_t>(remainingCompressed_, capacity));
    if (count != 0) {
        if (mappedPtr_) {
            std::memcpy(out, mappedPtr_, count);
            mappedPtr_ += count;
        } else {
            in_.read(reinterpret_cast<char*>(out), static_cast<std::streamsize>(count));
            if (static_cast<std::size_t>(in_.gcount()) != count)
                throw std::runtime_error("Truncated stored ZIP entry: " + info_.name);
        }
        crc_ = updateCrc(crc_, out, count);
        remainingCompressed_ -= count;
        produced_ += count;
        if (produced_ > info_.uncompressedSize)
            throw std::runtime_error("ZIP uncompressed-size mismatch: " + info_.name);
    }
    if (remainingCompressed_ == 0) validateCompletion();
    return count;
}

std::size_t ZipEntrySource::readDeflate(unsigned char* out, std::size_t capacity) {
    if (capacity == 0) return 0;
    const auto outputCapacity = static_cast<uInt>(
        std::min<std::size_t>(capacity, std::numeric_limits<uInt>::max()));
    for (;;) {
        if (z_->avail_in == 0 && remainingCompressed_ > 0) refillInput();
        z_->next_out = out;
        z_->avail_out = outputCapacity;
        const int status = inflate(z_, Z_NO_FLUSH);
        const std::size_t produced = outputCapacity - z_->avail_out;
        if (produced != 0) {
            crc_ = updateCrc(crc_, out, produced);
            produced_ += produced;
            if (produced_ > info_.uncompressedSize)
                throw std::runtime_error("ZIP uncompressed-size mismatch: " + info_.name);
        }
        if (status == Z_STREAM_END) {
            if (remainingCompressed_ != 0 || z_->avail_in != 0)
                throw std::runtime_error("ZIP compressed-size mismatch: " + info_.name);
            validateCompletion();
            return produced;
        }
        if (status != Z_OK && status != Z_BUF_ERROR)
            throw std::runtime_error("Invalid deflate ZIP entry: " + info_.name);
        if (produced != 0) return produced;
        if (z_->avail_in == 0 && remainingCompressed_ == 0)
            throw std::runtime_error("Truncated deflate ZIP entry: " + info_.name);
    }
}

std::size_t ZipEntrySource::read(unsigned char* out, std::size_t capacity) {
    if (eof_ || capacity == 0) return 0;
    if (!opened_) open();
    return info_.method == 0 ? readStored(out, capacity) : readDeflate(out, capacity);
}

ZipArchiveReader::ZipArchiveReader(const std::filesystem::path& path)
    : ZipArchiveReader(path, ZipOpenLimits{}) {}

ZipArchiveReader::ZipArchiveReader(const std::filesystem::path& path, const ZipOpenLimits& limits) : path_(path) {
    if (limits.maxFileBytes) {
        std::error_code ec;
        const auto size = std::filesystem::file_size(path_, ec);
        if (!ec && size > limits.maxFileBytes)
            throw std::runtime_error("ZIP package exceeds configured file-size limit");
    }
    try {
        mapped_ = std::make_shared<MappedFile>(path);
    } catch (const std::runtime_error&) {
        mapped_.reset();
    }

    if (mapped_) {
        const auto view = mapped_->view();
        entries_ = parseEntries(ByteSpan(reinterpret_cast<const unsigned char*>(view.data()), view.size()), limits);
    } else {
        // Mapping can be unavailable on network/constrained filesystems. Keep
        // the rare fallback correct and ZIP64-aware; payload reads still use
        // the source stream after the constructor returns.
        const auto bytes = readWholeFile(path_);
        entries_ = parseEntries(ByteSpan(bytes.data(), bytes.size()), limits);
    }
}

bool ZipArchiveReader::contains(const std::string& name) const { return entries_.contains(name); }

std::vector<std::string> ZipArchiveReader::names() const {
    std::vector<std::string> result;
    result.reserve(entries_.size());
    for (const auto& [name, _] : entries_) result.push_back(name);
    return result;
}

std::string ZipArchiveReader::readEntry(const std::string& name) const {
    const auto it = entries_.find(name);
    if (it == entries_.end()) throw std::runtime_error("ZIP entry not found: " + name);
    if (it->second.uncompressedSize > static_cast<std::uint64_t>(std::numeric_limits<std::size_t>::max()))
        throw std::runtime_error("ZIP entry is too large to materialize: " + name);

    auto source = openEntry(name);
    std::string result;
    result.reserve(static_cast<std::size_t>(it->second.uncompressedSize));
    std::array<unsigned char, 64 * 1024> buffer{};
    for (;;) {
        const auto count = source.read(buffer.data(), buffer.size());
        if (!count) break;
        result.append(reinterpret_cast<const char*>(buffer.data()), count);
    }
    return result;
}

void ZipArchiveReader::forEachChunk(const std::string& name,
                                    const std::function<void(const char*, std::size_t)>& emit) const {
    auto source = openEntry(name);
    std::array<unsigned char, 64 * 1024> buffer{};
    for (;;) {
        const auto count = source.read(buffer.data(), buffer.size());
        if (!count) break;
        if (emit) emit(reinterpret_cast<const char*>(buffer.data()), count);
    }
}

ZipEntrySource ZipArchiveReader::openEntry(const std::string& name) const {
    const auto it = entries_.find(name);
    if (it == entries_.end()) throw std::runtime_error("ZIP entry not found: " + name);
    if (mapped_) return ZipEntrySource(mapped_.get(), it->second);
    return ZipEntrySource(path_, it->second);
}

} // namespace xlpp::internal
