#include "ZipArchiveReader.h"
#include <array>
#include <zlib.h>
#include <algorithm>
#include <cstring>
#include <stdexcept>
#include <limits>
#include <utility>

namespace {
constexpr std::uint32_t kLocalSig = 0x04034b50u;
constexpr std::uint32_t kCentralSig = 0x02014b50u;
constexpr std::array<unsigned char, 4> kEocdSig = {'\x50', '\x4b', '\x05', '\x06'};


std::uint32_t readLE32(const unsigned char*& p) {
    const auto v = static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8) |
                   (static_cast<std::uint32_t>(p[2]) << 16) | (static_cast<std::uint32_t>(p[3]) << 24);
    p += 4;
    return v;
}

std::uint16_t readLE16(std::istream& in) {
    std::array<unsigned char, 2> bytes{};
    in.read(reinterpret_cast<char*>(bytes.data()), 2);
    if (in.gcount() != 2) throw std::runtime_error("Truncated ZIP file");
    return static_cast<std::uint16_t>(bytes[0] | (bytes[1] << 8));
}

std::uint32_t readLE32(std::istream& in) {
    std::array<unsigned char, 4> bytes{};
    in.read(reinterpret_cast<char*>(bytes.data()), 4);
    if (in.gcount() != 4) throw std::runtime_error("Truncated ZIP file");
    return static_cast<std::uint32_t>(bytes[0]) | (static_cast<std::uint32_t>(bytes[1]) << 8) |
           (static_cast<std::uint32_t>(bytes[2]) << 16) | (static_cast<std::uint32_t>(bytes[3]) << 24);
}

std::uint64_t readLE64(std::istream& in) {
    std::array<unsigned char, 8> bytes{};
    in.read(reinterpret_cast<char*>(bytes.data()), 8);
    if (in.gcount() != 8) throw std::runtime_error("Truncated ZIP file");
    std::uint64_t value = 0;
    for (int i = 7; i >= 0; --i) value = (value << 8) | bytes[static_cast<std::size_t>(i)];
    return value;
}

std::uint16_t le16At(std::string_view data, std::size_t off) {
    if (off > data.size() || data.size() - off < 2) throw std::runtime_error("Truncated ZIP structure");
    return static_cast<std::uint16_t>(static_cast<unsigned char>(data[off]) |
        (static_cast<std::uint16_t>(static_cast<unsigned char>(data[off + 1])) << 8));
}

std::uint32_t le32At(std::string_view data, std::size_t off) {
    if (off > data.size() || data.size() - off < 4) throw std::runtime_error("Truncated ZIP structure");
    return static_cast<std::uint32_t>(static_cast<unsigned char>(data[off])) |
        (static_cast<std::uint32_t>(static_cast<unsigned char>(data[off + 1])) << 8) |
        (static_cast<std::uint32_t>(static_cast<unsigned char>(data[off + 2])) << 16) |
        (static_cast<std::uint32_t>(static_cast<unsigned char>(data[off + 3])) << 24);
}

std::uint64_t le64At(std::string_view data, std::size_t off) {
    if (off > data.size() || data.size() - off < 8) throw std::runtime_error("Truncated ZIP structure");
    std::uint64_t value = 0;
    for (int i = 7; i >= 0; --i)
        value = (value << 8) | static_cast<unsigned char>(data[off + static_cast<std::size_t>(i)]);
    return value;
}

std::size_t findValidEocd(std::string_view tail) {
    constexpr std::string_view signature("\x50\x4b\x05\x06", 4);
    if (tail.size() < 22) return std::string_view::npos;
    std::size_t searchEnd = tail.size();
    while (searchEnd >= 4) {
        const auto pos = tail.substr(0, searchEnd).rfind(signature);
        if (pos == std::string_view::npos) return pos;
        if (pos + 22 <= tail.size()) {
            const auto commentLen = le16At(tail, pos + 20);
            if (pos + 22u + static_cast<std::size_t>(commentLen) == tail.size()) return pos;
        }
        if (pos == 0) break;
        searchEnd = pos;
    }
    return std::string_view::npos;
}

void parseZip64Extra(std::string_view extra,
                     bool needUncompressed, bool needCompressed, bool needOffset, bool needDisk,
                     std::uint64_t& uncompressed, std::uint64_t& compressed,
                     std::uint64_t& localOffset, std::uint32_t& diskStart) {
    std::size_t pos = 0;
    while (pos + 4 <= extra.size()) {
        const auto id = le16At(extra, pos);
        const auto len = le16At(extra, pos + 2);
        pos += 4;
        if (len > extra.size() - pos) throw std::runtime_error("Truncated ZIP extra field");
        if (id == 0x0001u) {
            const auto field = extra.substr(pos, len);
            std::size_t cursor = 0;
            auto take64 = [&]() {
                if (cursor + 8 > field.size()) throw std::runtime_error("Truncated ZIP64 extra field");
                const auto value = le64At(field, cursor); cursor += 8; return value;
            };
            auto take32 = [&]() {
                if (cursor + 4 > field.size()) throw std::runtime_error("Truncated ZIP64 extra field");
                const auto value = le32At(field, cursor); cursor += 4; return value;
            };
            if (needUncompressed) uncompressed = take64();
            if (needCompressed) compressed = take64();
            if (needOffset) localOffset = take64();
            if (needDisk) diskStart = take32();
            return;
        }
        pos += len;
    }
    if (needUncompressed || needCompressed || needOffset || needDisk)
        throw std::runtime_error("ZIP64 sentinel is missing the ZIP64 extra field");
}

} // namespace

namespace xlpp::internal {

ZipEntrySource::ZipEntrySource(std::filesystem::path path, ZipEntryInfo info)
    : path_(std::move(path)), info_(info) {}

ZipEntrySource::ZipEntrySource(const MappedFile* mapped, ZipEntryInfo info)
    : info_(std::move(info)), mapped_(mapped) {}

void ZipEntrySource::releaseInflater() noexcept {
    if (z_) {
        inflateEnd(z_);
        delete z_;
        z_ = nullptr;
    }
}

void ZipEntrySource::moveFrom(ZipEntrySource&& other) noexcept {
    path_ = std::move(other.path_);
    info_ = std::move(other.info_);
    in_ = std::move(other.in_);
    mapped_ = other.mapped_;
    mappedPtr_ = other.mappedPtr_;
    inbuf_ = other.inbuf_;
    crc_ = other.crc_;
    remainingCompressed_ = other.remainingCompressed_;
    producedUncompressed_ = other.producedUncompressed_;
    opened_ = other.opened_;
    eof_ = other.eof_;
    complete_ = other.complete_;
    z_ = std::exchange(other.z_, nullptr);

    // zlib may retain next_in inside the object's refill buffer. Rebase that
    // pointer when an already-open source is moved. Mapped input points into
    // the shared MappedFile and requires no rebasing.
    if (z_ && z_->next_in) {
        const auto* begin = other.inbuf_.data();
        const auto* end = begin + other.inbuf_.size();
        if (z_->next_in >= begin && z_->next_in <= end) {
            const auto offset = static_cast<std::size_t>(z_->next_in - begin);
            z_->next_in = inbuf_.data() + offset;
        }
    }

    other.mapped_ = nullptr;
    other.mappedPtr_ = nullptr;
    other.remainingCompressed_ = 0;
    other.producedUncompressed_ = 0;
    other.opened_ = false;
    other.eof_ = true;
    other.complete_ = false;
}

ZipEntrySource::ZipEntrySource(ZipEntrySource&& other) noexcept { moveFrom(std::move(other)); }

ZipEntrySource& ZipEntrySource::operator=(ZipEntrySource&& other) noexcept {
    if (this != &other) {
        releaseInflater();
        moveFrom(std::move(other));
    }
    return *this;
}

ZipEntrySource::~ZipEntrySource() { releaseInflater(); }

void ZipEntrySource::open() {
    if (mapped_) {
        const auto size = mapped_->size();
        const auto localOffset = static_cast<std::size_t>(info_.localOffset);
        if (localOffset > size || size - localOffset < 30)
            throw std::runtime_error("Truncated ZIP local header: " + info_.name);
        const auto* p = reinterpret_cast<const unsigned char*>(mapped_->data()) + localOffset;
        const auto* header = p;
        if (readLE32(p) != kLocalSig) throw std::runtime_error("Invalid ZIP local header: " + info_.name);
        const auto localMethod = static_cast<std::uint16_t>(header[8] | (static_cast<std::uint16_t>(header[9]) << 8));
        if (localMethod != info_.method) throw std::runtime_error("ZIP local/central compression method mismatch: " + info_.name);
        const auto nameLen = static_cast<std::uint16_t>(header[26] | (static_cast<std::uint16_t>(header[27]) << 8));
        const auto extraLen = static_cast<std::uint16_t>(header[28] | (static_cast<std::uint16_t>(header[29]) << 8));
        const auto dataOffset = localOffset + 30u + static_cast<std::size_t>(nameLen) + static_cast<std::size_t>(extraLen);
        if (dataOffset > size || info_.compressedSize > size - dataOffset)
            throw std::runtime_error("ZIP entry data exceeds archive bounds: " + info_.name);
        const std::string_view localName(reinterpret_cast<const char*>(header + 30), nameLen);
        if (localName != info_.name) throw std::runtime_error("ZIP local/central entry name mismatch: " + info_.name);
        mappedPtr_ = mapped_->data() + dataOffset;
        remainingCompressed_ = info_.compressedSize;
    } else {
        in_.open(path_, std::ios::binary);
        if (!in_) throw std::runtime_error("Cannot open XLSX package: " + path_.string());
        in_.seekg(0, std::ios::end);
        const auto fileSize = in_.tellg();
        if (fileSize < 0 || static_cast<std::uint64_t>(fileSize) < static_cast<std::uint64_t>(info_.localOffset) + 30u)
            throw std::runtime_error("Truncated ZIP local header: " + info_.name);
        if (info_.localOffset > static_cast<std::uint64_t>(std::numeric_limits<std::streamoff>::max()))
            throw std::runtime_error("ZIP local header offset exceeds platform stream limits: " + info_.name);
        in_.seekg(static_cast<std::streamoff>(info_.localOffset));
        if (readLE32(in_) != kLocalSig) throw std::runtime_error("Invalid ZIP local header: " + info_.name);
        (void)readLE16(in_); // version
        (void)readLE16(in_); // flags
        const auto localMethod = readLE16(in_);
        if (localMethod != info_.method) throw std::runtime_error("ZIP local/central compression method mismatch: " + info_.name);
        in_.seekg(14, std::ios::cur); // time/date/crc/sizes
        const auto nameLen = readLE16(in_);
        const auto extraLen = readLE16(in_);
        std::string localName(nameLen, '\0');
        if (nameLen) {
            in_.read(localName.data(), static_cast<std::streamsize>(nameLen));
            if (static_cast<std::size_t>(in_.gcount()) != nameLen) throw std::runtime_error("Truncated ZIP local entry name");
        }
        if (localName != info_.name) throw std::runtime_error("ZIP local/central entry name mismatch: " + info_.name);
        in_.seekg(extraLen, std::ios::cur);
        const auto dataOffset = static_cast<std::uint64_t>(info_.localOffset) + 30u + nameLen + extraLen;
        if (dataOffset > static_cast<std::uint64_t>(fileSize) || info_.compressedSize > static_cast<std::uint64_t>(fileSize) - dataOffset)
            throw std::runtime_error("ZIP entry data exceeds archive bounds: " + info_.name);
        remainingCompressed_ = info_.compressedSize;
    }
    crc_ = static_cast<std::uint32_t>(crc32(0, Z_NULL, 0));
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
        std::min<std::uint64_t>(remainingCompressed_, static_cast<std::uint64_t>(inbuf_.size())));
    if (mappedPtr_) {
        std::memcpy(inbuf_.data(), mappedPtr_, count);
        mappedPtr_ += count;
    } else {
        in_.read(reinterpret_cast<char*>(inbuf_.data()), static_cast<std::streamsize>(count));
        if (static_cast<std::size_t>(in_.gcount()) != count)
            throw std::runtime_error("Truncated ZIP entry data");
    }
    z_->next_in = inbuf_.data();
    z_->avail_in = static_cast<uInt>(count);
    remainingCompressed_ -= static_cast<std::uint64_t>(count);
    return true;
}

std::size_t ZipEntrySource::readStored(unsigned char* out, std::size_t capacity) {
    if (capacity == 0) return 0;
    const auto count = static_cast<std::size_t>(
        std::min<std::uint64_t>(remainingCompressed_, static_cast<std::uint64_t>(capacity)));
    if (mappedPtr_) {
        std::memcpy(out, mappedPtr_, count);
        mappedPtr_ += count;
    } else {
        in_.read(reinterpret_cast<char*>(out), static_cast<std::streamsize>(count));
        if (static_cast<std::size_t>(in_.gcount()) != count)
            throw std::runtime_error("Truncated stored ZIP entry");
    }
    if (count > info_.uncompressedSize - std::min(producedUncompressed_, info_.uncompressedSize))
        throw std::runtime_error("ZIP entry expands beyond declared uncompressed size: " + info_.name);
    producedUncompressed_ += static_cast<std::uint64_t>(count);
    crc_ = static_cast<std::uint32_t>(crc32(crc_, out, static_cast<uInt>(count)));
    remainingCompressed_ -= static_cast<std::uint64_t>(count);
    if (remainingCompressed_ == 0) {
        eof_ = true;
        if (producedUncompressed_ != info_.uncompressedSize)
            throw std::runtime_error("ZIP uncompressed size mismatch: " + info_.name);
        if (crc_ != info_.crc) throw std::runtime_error("ZIP CRC mismatch: " + info_.name);
        complete_ = true;
    }
    return count;
}

std::size_t ZipEntrySource::readDeflate(unsigned char* out, std::size_t capacity) {
    if (capacity == 0) return 0;
    for (;;) {
        if (z_->avail_in == 0 && remainingCompressed_ > 0) refillInput();
        z_->next_out = out;
        z_->avail_out = static_cast<uInt>(capacity);
        const int status = inflate(z_, Z_NO_FLUSH);
        const std::size_t produced = capacity - z_->avail_out;
        if (produced) {
            if (producedUncompressed_ > info_.uncompressedSize ||
                static_cast<std::uint64_t>(produced) > info_.uncompressedSize - producedUncompressed_)
                throw std::runtime_error("ZIP entry expands beyond declared uncompressed size: " + info_.name);
            producedUncompressed_ += static_cast<std::uint64_t>(produced);
            crc_ = static_cast<std::uint32_t>(crc32(crc_, out, static_cast<uInt>(produced)));
        }
        if (status == Z_STREAM_END) {
            eof_ = true;
            if (producedUncompressed_ != info_.uncompressedSize)
                throw std::runtime_error("ZIP uncompressed size mismatch: " + info_.name);
            if (crc_ != info_.crc) throw std::runtime_error("ZIP CRC mismatch: " + info_.name);
            complete_ = true;
            return produced;
        }
        if (status != Z_OK && status != Z_BUF_ERROR)
            throw std::runtime_error("inflate failed");
        if (produced) return produced;
        if (z_->avail_in == 0 && remainingCompressed_ == 0)
            throw std::runtime_error("Truncated deflate ZIP entry");
    }
}

std::size_t ZipEntrySource::read(unsigned char* out, std::size_t capacity) {
    if (eof_ || capacity == 0) return 0;
    if (!opened_) open();
    return info_.method == 0 ? readStored(out, capacity) : readDeflate(out, capacity);
}

ZipArchiveReader::ZipArchiveReader(const std::filesystem::path& path)
    : ZipArchiveReader(path, ZipArchiveReaderLimits{}) {}

ZipArchiveReader::ZipArchiveReader(const std::filesystem::path& path, const ZipArchiveReaderLimits& limits) : path_(path) {
    // Keep the mapped view for zero-copy entry reads. Central-directory parsing
    // intentionally uses one unified seek-based path so mapped and fallback
    // modes have identical ZIP/ZIP64 validation semantics.
    try {
        mapped_ = std::make_shared<MappedFile>(path);
    } catch (...) {
        mapped_.reset();
    }

    std::ifstream in(path_, std::ios::binary);
    if (!in) throw std::runtime_error("Cannot open XLSX package: " + path_.string());
    in.seekg(0, std::ios::end);
    const auto endOff = in.tellg();
    if (endOff < 22) throw std::runtime_error("Invalid ZIP package: file too small");
    const auto fileSize = static_cast<std::uint64_t>(endOff);
    if (limits.maxFileBytes && fileSize > limits.maxFileBytes)
        throw std::runtime_error("ZIP package exceeds configured file-size limit");
    const auto tailSize64 = std::min<std::uint64_t>(fileSize, 22u + 65535u);
    const auto tailStart = fileSize - tailSize64;
    in.seekg(static_cast<std::streamoff>(tailStart));
    std::string tail(static_cast<std::size_t>(tailSize64), '\0');
    in.read(tail.data(), static_cast<std::streamsize>(tail.size()));
    if (static_cast<std::size_t>(in.gcount()) != tail.size()) throw std::runtime_error("Truncated ZIP tail");

    const auto eocdPos = findValidEocd(tail);
    if (eocdPos == std::string_view::npos)
        throw std::runtime_error("Invalid ZIP package: end of central directory not found");
    const auto eocdOffset = tailStart + eocdPos;
    const auto tailView = std::string_view(tail);
    const auto disk16 = le16At(tailView, eocdPos + 4);
    const auto centralDisk16 = le16At(tailView, eocdPos + 6);
    const auto entriesOnDisk16 = le16At(tailView, eocdPos + 8);
    const auto count16 = le16At(tailView, eocdPos + 10);
    const auto size32 = le32At(tailView, eocdPos + 12);
    const auto offset32 = le32At(tailView, eocdPos + 16);
    if (disk16 != 0 || centralDisk16 != 0)
        throw std::runtime_error("Multi-disk ZIP packages are not supported");

    std::uint64_t cdCount = count16;
    std::uint64_t entriesOnDisk = entriesOnDisk16;
    std::uint64_t cdSize = size32;
    std::uint64_t cdOffset = offset32;
    const bool zip64 = count16 == 0xFFFFu || entriesOnDisk16 == 0xFFFFu ||
                       size32 == 0xFFFFFFFFu || offset32 == 0xFFFFFFFFu;
    if (zip64) {
        if (eocdOffset < 20) throw std::runtime_error("ZIP64 locator is missing");
        in.clear();
        in.seekg(static_cast<std::streamoff>(eocdOffset - 20));
        if (readLE32(in) != 0x07064b50u) throw std::runtime_error("ZIP64 locator is missing");
        const auto eocd64Disk = readLE32(in);
        const auto eocd64Offset = readLE64(in);
        const auto diskCount = readLE32(in);
        if (eocd64Disk != 0 || diskCount != 1) throw std::runtime_error("Multi-disk ZIP64 packages are not supported");
        if (eocd64Offset > fileSize || fileSize - eocd64Offset < 56)
            throw std::runtime_error("ZIP64 end record exceeds archive bounds");
        if (eocd64Offset > static_cast<std::uint64_t>(std::numeric_limits<std::streamoff>::max()))
            throw std::runtime_error("ZIP64 offset exceeds platform stream limits");
        in.clear();
        in.seekg(static_cast<std::streamoff>(eocd64Offset));
        if (readLE32(in) != 0x06064b50u) throw std::runtime_error("Invalid ZIP64 end record");
        const auto recordSize = readLE64(in);
        if (recordSize < 44 || recordSize > fileSize - eocd64Offset - 12)
            throw std::runtime_error("Invalid ZIP64 end-record size");
        (void)readLE16(in); (void)readLE16(in); // versions
        const auto disk = readLE32(in);
        const auto centralDisk = readLE32(in);
        entriesOnDisk = readLE64(in);
        cdCount = readLE64(in);
        cdSize = readLE64(in);
        cdOffset = readLE64(in);
        if (disk != 0 || centralDisk != 0 || entriesOnDisk != cdCount)
            throw std::runtime_error("Multi-disk ZIP64 packages are not supported");
    } else if (entriesOnDisk != cdCount) {
        throw std::runtime_error("Multi-disk ZIP packages are not supported");
    }

    if (cdOffset > fileSize || cdSize > fileSize - cdOffset)
        throw std::runtime_error("Invalid ZIP package: central directory exceeds archive bounds");
    if (cdCount > cdSize / 46u + 1u)
        throw std::runtime_error("Invalid ZIP package: central-directory entry count is impossible");
    if (limits.maxEntries && cdCount > limits.maxEntries)
        throw std::runtime_error("ZIP package exceeds configured entry-count limit");
    if (cdOffset > static_cast<std::uint64_t>(std::numeric_limits<std::streamoff>::max()))
        throw std::runtime_error("ZIP central-directory offset exceeds platform stream limits");
    if (cdCount == 0) return;

    in.clear();
    in.seekg(static_cast<std::streamoff>(cdOffset));
    std::uint64_t consumed = 0;
    std::uint64_t totalDeclaredUncompressed = 0;
    for (std::uint64_t index = 0; index < cdCount; ++index) {
        if (consumed > cdSize || cdSize - consumed < 46) throw std::runtime_error("Truncated ZIP central directory");
        if (readLE32(in) != kCentralSig) throw std::runtime_error("Invalid ZIP central directory");
        (void)readLE16(in); (void)readLE16(in); // versions
        const auto flags = readLE16(in);
        const auto method = readLE16(in);
        (void)readLE16(in); (void)readLE16(in); // time/date
        const auto crc = readLE32(in);
        const auto compressed32 = readLE32(in);
        const auto uncompressed32 = readLE32(in);
        const auto nameLen = readLE16(in);
        const auto extraLen = readLE16(in);
        const auto commentLen = readLE16(in);
        const auto diskStart16 = readLE16(in);
        (void)readLE16(in); (void)readLE32(in); // attributes
        const auto localOffset32 = readLE32(in);
        if ((flags & 0x0001u) != 0) throw std::runtime_error("Traditional ZIP encryption is not supported");
        if (method != 0 && method != 8) throw std::runtime_error("Unsupported ZIP compression method");
        const auto recordSize = 46u + static_cast<std::uint64_t>(nameLen) + extraLen + commentLen;
        if (recordSize > cdSize - consumed) throw std::runtime_error("Truncated ZIP central-directory record");

        std::string name(nameLen, '\0');
        if (nameLen) {
            in.read(name.data(), static_cast<std::streamsize>(nameLen));
            if (static_cast<std::size_t>(in.gcount()) != nameLen) throw std::runtime_error("Truncated ZIP entry name");
        }
        std::string extra(extraLen, '\0');
        if (extraLen) {
            in.read(extra.data(), static_cast<std::streamsize>(extraLen));
            if (static_cast<std::size_t>(in.gcount()) != extraLen) throw std::runtime_error("Truncated ZIP entry extra field");
        }
        if (commentLen) in.seekg(static_cast<std::streamoff>(commentLen), std::ios::cur);
        if (!in) throw std::runtime_error("Truncated ZIP central-directory record");

        std::uint64_t compressed = compressed32;
        std::uint64_t uncompressed = uncompressed32;
        std::uint64_t localOffset = localOffset32;
        std::uint32_t diskStart = diskStart16;
        parseZip64Extra(extra,
                        uncompressed32 == 0xFFFFFFFFu, compressed32 == 0xFFFFFFFFu,
                        localOffset32 == 0xFFFFFFFFu, diskStart16 == 0xFFFFu,
                        uncompressed, compressed, localOffset, diskStart);
        if (diskStart != 0) throw std::runtime_error("Multi-disk ZIP entries are not supported");
        if (limits.maxEntryBytes && uncompressed > limits.maxEntryBytes)
            throw std::runtime_error("ZIP entry exceeds configured uncompressed-size limit: " + name);
        if (limits.maxTotalBytes &&
            (uncompressed > limits.maxTotalBytes || totalDeclaredUncompressed > limits.maxTotalBytes - uncompressed))
            throw std::runtime_error("ZIP package exceeds configured total uncompressed-size limit");
        totalDeclaredUncompressed += uncompressed;
        if (method != 0 && compressed < uncompressed &&
            compressed <= (std::numeric_limits<std::uint64_t>::max() - 4096u) / 1032u &&
            uncompressed > compressed * 1032u + 4096u)
            throw std::runtime_error("ZIP entry has an implausible compression ratio: " + name);
        if (localOffset > fileSize || fileSize - localOffset < 30)
            throw std::runtime_error("ZIP local header offset exceeds archive bounds");
        if (name.empty()) throw std::runtime_error("ZIP entry name cannot be empty");
        if (entries_.contains(name)) throw std::runtime_error("Duplicate ZIP entry: " + name);
        entries_.emplace(name, ZipEntryInfo{name, localOffset, compressed, uncompressed, method, crc});
        consumed += recordSize;
    }
    if (consumed > cdSize) throw std::runtime_error("Invalid ZIP central-directory size");
}

bool ZipArchiveReader::contains(const std::string& name) const {
    return entries_.contains(name);
}

std::vector<std::string> ZipArchiveReader::names() const {
    std::vector<std::string> result;
    result.reserve(entries_.size());
    for (const auto& [name, _] : entries_) result.push_back(name);
    return result;
}

std::string ZipArchiveReader::readEntry(const std::string& name) const {
    auto source = openEntry(name);
    std::string result;
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
    auto info = it->second;
    info.name = name;
    if (mapped_)
        return ZipEntrySource(mapped_.get(), std::move(info));
    return ZipEntrySource(path_, std::move(info));
}

} // namespace xlpp::internal
