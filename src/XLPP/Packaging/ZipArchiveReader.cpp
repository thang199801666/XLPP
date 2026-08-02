#include "ZipArchiveReader.h"
#include <zlib.h>
#include <algorithm>
#include <cstring>
#include <stdexcept>

namespace {
constexpr std::uint32_t kLocalSig = 0x04034b50u;
constexpr std::uint32_t kCentralSig = 0x02014b50u;
constexpr std::array<unsigned char, 4> kEocdSig = {'\x50', '\x4b', '\x05', '\x06'};

std::uint16_t readLE16(const unsigned char*& p) {
    const auto v = static_cast<std::uint16_t>(p[0] | (p[1] << 8));
    p += 2;
    return v;
}

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

} // namespace

namespace xlpp::internal {

ZipEntrySource::ZipEntrySource(std::filesystem::path path, ZipEntryInfo info)
    : path_(std::move(path)), info_(info) {}

ZipEntrySource::ZipEntrySource(const MappedFile* mapped, ZipEntryInfo info)
    : info_(std::move(info)), mapped_(mapped) {}

ZipEntrySource::~ZipEntrySource() {
    if (z_) inflateEnd(z_);
    delete z_;
}

void ZipEntrySource::open() {
    if (mapped_) {
        const auto* p = reinterpret_cast<const unsigned char*>(mapped_->data()) + info_.localOffset;
        if (readLE32(p) != kLocalSig) throw std::runtime_error("Invalid ZIP local header: " + info_.name);
        p += 22; // skip to name length
        const auto nameLen = readLE16(p);
        const auto extraLen = readLE16(p);
        mappedPtr_ = reinterpret_cast<const char*>(p + nameLen + extraLen);
        remainingCompressed_ = info_.compressedSize;
    } else {
        in_.open(path_, std::ios::binary);
        if (!in_) throw std::runtime_error("Cannot open XLSX package: " + path_.string());
        in_.seekg(info_.localOffset + 26);
        const auto nameLen = readLE16(in_);
        const auto extraLen = readLE16(in_);
        in_.seekg(nameLen + extraLen, std::ios::cur);
        remainingCompressed_ = info_.compressedSize;
    }
    crc_ = crc32(0, Z_NULL, 0);
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
        std::min<std::uint32_t>(remainingCompressed_, static_cast<std::uint32_t>(inbuf_.size())));
    if (mappedPtr_) {
        std::memcpy(inbuf_.data(), mappedPtr_, count);
        mappedPtr_ += count;
    } else {
        in_.read(reinterpret_cast<char*>(inbuf_.data()), count);
        if (static_cast<std::size_t>(in_.gcount()) != count)
            throw std::runtime_error("Truncated ZIP entry data");
    }
    z_->next_in = inbuf_.data();
    z_->avail_in = static_cast<uInt>(count);
    remainingCompressed_ -= static_cast<std::uint32_t>(count);
    return true;
}

std::size_t ZipEntrySource::readStored(unsigned char* out, std::size_t capacity) {
    if (capacity == 0) return 0;
    const auto count = static_cast<std::size_t>(
        std::min<std::uint32_t>(remainingCompressed_, static_cast<std::uint32_t>(capacity)));
    if (mappedPtr_) {
        std::memcpy(out, mappedPtr_, count);
        mappedPtr_ += count;
    } else {
        in_.read(reinterpret_cast<char*>(out), count);
        if (static_cast<std::size_t>(in_.gcount()) != count)
            throw std::runtime_error("Truncated stored ZIP entry");
    }
    crc_ = crc32(crc_, out, static_cast<uInt>(count));
    remainingCompressed_ -= static_cast<std::uint32_t>(count);
    if (remainingCompressed_ == 0) {
        eof_ = true;
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
        if (produced) crc_ = crc32(crc_, out, static_cast<uInt>(produced));
        if (status == Z_STREAM_END) {
            eof_ = true;
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

ZipArchiveReader::ZipArchiveReader(const std::filesystem::path& path) : path_(path) {
    // Try memory-mapped I/O first
    try {
        mapped_ = std::make_shared<MappedFile>(path);
    } catch (...) {
        mapped_.reset();
    }

    if (mapped_) {
        const auto view = mapped_->view();
        const auto* data = reinterpret_cast<const unsigned char*>(view.data());
        const auto size = view.size();
        if (size < 22) throw std::runtime_error("Invalid ZIP package: file too small");

        // Find EOCD in last 65KB
        const auto tail = std::min<std::size_t>(size, 65535 + 22);
        const auto* tailStart = data + size - tail;
        const auto tailEnd = data + size;
        const auto eocdSig = std::string_view(reinterpret_cast<const char*>(kEocdSig.data()), 4);
        const auto tailView = std::string_view(reinterpret_cast<const char*>(tailStart), tail);
        const auto pos = tailView.rfind(eocdSig);
        if (pos == std::string_view::npos)
            throw std::runtime_error("Invalid ZIP package: end of central directory not found");

        const auto* eocdPtr = tailStart + pos;
        const auto cdCount = static_cast<std::uint16_t>(eocdPtr[10] | (static_cast<std::uint16_t>(eocdPtr[11]) << 8));
        const auto cdOffset = static_cast<std::uint32_t>(
            eocdPtr[16] | (static_cast<std::uint32_t>(eocdPtr[17]) << 8) |
            (static_cast<std::uint32_t>(eocdPtr[18]) << 16) | (static_cast<std::uint32_t>(eocdPtr[19]) << 24));
        if (cdCount == 0) return;

        const auto* cd = data + cdOffset;
        for (std::uint16_t idx = 0; idx < cdCount; ++idx) {
            if (readLE32(cd) != kCentralSig) throw std::runtime_error("Invalid ZIP central directory");
            (void)readLE16(cd); // version made by
            (void)readLE16(cd); // version needed
            (void)readLE16(cd); // flags
            const auto method = readLE16(cd);
            (void)readLE16(cd); // mod time
            (void)readLE16(cd); // mod date
            const auto crc = readLE32(cd);
            const auto compressedSize = readLE32(cd);
            const auto uncompressedSize = readLE32(cd);
            const auto nameLen = readLE16(cd);
            const auto extraLen = readLE16(cd);
            const auto commentLen = readLE16(cd);
            (void)readLE16(cd); // disk number
            (void)readLE16(cd); // internal attrs
            (void)readLE32(cd); // external attrs
            const auto localOffset = readLE32(cd);
            std::string name(reinterpret_cast<const char*>(cd), nameLen);
            cd += nameLen + extraLen + commentLen;
            entries_[name] = ZipEntryInfo{name, localOffset, compressedSize, uncompressedSize, method, crc};
        }
    } else {
        // Fallback: ifstream
        std::ifstream in(path_, std::ios::binary);
        if (!in) throw std::runtime_error("Cannot open XLSX package: " + path_.string());
        in.seekg(0, std::ios::end);
        const auto fileSize = in.tellg();
        if (fileSize < 22) throw std::runtime_error("Invalid ZIP package: file too small");
        const auto tail = std::min<std::streamoff>(fileSize, 22 + 65535);
        in.seekg(fileSize - tail);
        std::string tailData;
        tailData.resize(static_cast<std::size_t>(tail));
        in.read(tailData.data(), static_cast<std::streamsize>(tail));
        const auto endSignature = std::string("\x50\x4b\x05\x06", 4);
        const auto pos = tailData.rfind(endSignature);
        if (pos == std::string::npos) throw std::runtime_error("Invalid ZIP package: EOCD not found");
        const auto eocd = pos + 16;
        const auto cdOffset = static_cast<std::uint32_t>(
            static_cast<unsigned char>(tailData[eocd]) |
            (static_cast<std::uint32_t>(static_cast<unsigned char>(tailData[eocd + 1])) << 8) |
            (static_cast<std::uint32_t>(static_cast<unsigned char>(tailData[eocd + 2])) << 16) |
            (static_cast<std::uint32_t>(static_cast<unsigned char>(tailData[eocd + 3])) << 24));
        const auto cdCount = static_cast<std::uint16_t>(
            static_cast<unsigned char>(tailData[pos + 10]) |
            (static_cast<std::uint16_t>(static_cast<unsigned char>(tailData[pos + 11])) << 8));
        if (cdCount == 0) return;
        in.seekg(cdOffset);
        for (std::uint16_t index = 0; index < cdCount; ++index) {
            if (readLE32(in) != kCentralSig) throw std::runtime_error("Invalid ZIP central directory");
            (void)readLE16(in); (void)readLE16(in); (void)readLE16(in);
            const auto method = readLE16(in);
            (void)readLE16(in); (void)readLE16(in);
            const auto crc = readLE32(in);
            const auto compressedSize = readLE32(in);
            const auto uncompressedSize = readLE32(in);
            const auto nameLen = readLE16(in);
            const auto extraLen = readLE16(in);
            const auto commentLen = readLE16(in);
            (void)readLE16(in); (void)readLE16(in); (void)readLE32(in);
            const auto localOffset = readLE32(in);
            std::string name(nameLen, '\0');
            if (nameLen) in.read(name.data(), nameLen);
            in.seekg(extraLen + commentLen, std::ios::cur);
            entries_[name] = ZipEntryInfo{name, localOffset, compressedSize, uncompressedSize, method, crc};
        }
    }
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
