#include "XlsBinaryReader.h"
#include <XLPP/Workbook/Workbook.h>
#include <cstring>
#include <stdexcept>
#include <string>
#include <vector>
#include <cstdint>
#include <limits>

namespace xlpp {
namespace internal {

bool isOle2CompoundFile(std::string_view bytes) noexcept {
    return bytes.size() >= Ole2CompoundMagic.size() &&
           std::memcmp(bytes.data(), Ole2CompoundMagic.data(), Ole2CompoundMagic.size()) == 0;
}

namespace {

std::uint16_t le16(const std::vector<unsigned char>& d, std::size_t off) {
    if (off + 2 > d.size()) throw std::runtime_error("XLS: truncated 16-bit field");
    return static_cast<std::uint16_t>(d[off]) | (static_cast<std::uint16_t>(d[off + 1]) << 8);
}

std::uint32_t le32(const std::vector<unsigned char>& d, std::size_t off) {
    if (off + 4 > d.size()) throw std::runtime_error("XLS: truncated 32-bit field");
    return static_cast<std::uint32_t>(d[off]) | (static_cast<std::uint32_t>(d[off + 1]) << 8) |
           (static_cast<std::uint32_t>(d[off + 2]) << 16) | (static_cast<std::uint32_t>(d[off + 3]) << 24);
}

std::uint64_t le64(const std::vector<unsigned char>& d, std::size_t off) {
    if (off + 8 > d.size()) throw std::runtime_error("XLS: truncated 64-bit field");
    std::uint64_t value = 0;
    for (int i = 7; i >= 0; --i) value = (value << 8) | d[off + static_cast<std::size_t>(i)];
    return value;
}

constexpr std::uint32_t FatSector = 0xFFFFFFFD;
constexpr std::uint32_t EndOfChain = 0xFFFFFFFE;
constexpr std::uint32_t FreeSector = 0xFFFFFFFF;

// ---------------------------------------------------------------------------
// OLE2 compound file (CFB) container
// ---------------------------------------------------------------------------

struct CfbFile {
    std::vector<unsigned char> data;
    std::size_t sectorSize = 512;
    std::size_t miniSectorSize = 64;
    std::size_t miniStreamCutoff = 4096;
    std::vector<std::uint32_t> fat;
    std::vector<std::uint32_t> miniFat;
    std::uint32_t firstDirSector = EndOfChain;
    std::uint32_t rootStart = EndOfChain;
    std::uint64_t rootSize = 0;
};

std::size_t sectorOffset(const CfbFile& cfb, std::uint32_t sector) {
    if (sector == EndOfChain || sector == FreeSector)
        throw std::runtime_error("XLS: chain references a terminal sector");
    const auto offset = 512 + static_cast<std::size_t>(sector) * cfb.sectorSize;
    if (offset > cfb.data.size())
        throw std::runtime_error("XLS: sector reference outside the file");
    return offset;
}

CfbFile parseCfb(const std::vector<unsigned char>& data) {
    if (data.size() < 512 || !isOle2CompoundFile(std::string_view(
            reinterpret_cast<const char*>(data.data()), data.size())))
        throw std::runtime_error("XLS: not an OLE2 compound file");
    CfbFile cfb;
    cfb.data = data;
    const auto sectorShift = static_cast<std::size_t>(data[0x1E]);
    const auto miniShift = static_cast<std::size_t>(data[0x1F]);
    if (sectorShift < 7 || sectorShift > 12) throw std::runtime_error("XLS: unsupported sector shift");
    if (miniShift < 6 || miniShift > 7) throw std::runtime_error("XLS: unsupported mini sector shift");
    cfb.sectorSize = std::size_t{1} << sectorShift;
    cfb.miniSectorSize = std::size_t{1} << miniShift;
    const auto numFatSectors = le32(data, 0x2C);
    cfb.firstDirSector = le32(data, 0x30);
    cfb.miniStreamCutoff = le32(data, 0x38);
    const auto firstMiniFatSector = le32(data, 0x3C);
    const auto numMiniFatSectors = le32(data, 0x40);
    const auto firstDifatSector = le32(data, 0x44);
    const auto numDifatSectors = le32(data, 0x48);

    // Hardening: sector references must point inside the file and the number of
    // FAT/mini-FAT sectors must be plausible, otherwise a malformed container
    // could force huge allocations or out-of-bounds reads.
    const auto fileSectors = cfb.data.size() / cfb.sectorSize;
    const auto maxSectorPerFat = cfb.data.size() / 4; // each FAT entry needs 4 bytes
    if (numFatSectors > fileSectors || numFatSectors > maxSectorPerFat)
        throw std::runtime_error("XLS: implausible FAT sector count");
    if (numMiniFatSectors > fileSectors)
        throw std::runtime_error("XLS: implausible mini FAT sector count");
    if (numDifatSectors > fileSectors)
        throw std::runtime_error("XLS: implausible DIFAT sector count");

    // Build the FAT from the DIFAT table (109 entries inline + chained sectors).
    auto fatSectorList = std::vector<std::uint32_t>();
    for (std::size_t i = 0; i < 109; ++i) {
        const auto entry = le32(data, 0x4C + i * 4);
        if (entry != FreeSector) fatSectorList.push_back(entry);
    }
    auto difat = firstDifatSector;
    for (std::uint32_t chain = 0; chain < numDifatSectors; ++chain) {
        if (difat == EndOfChain || difat == FreeSector || difat >= fileSectors)
            throw std::runtime_error("XLS: corrupt DIFAT chain");
        const auto base = sectorOffset(cfb, difat);
        const auto entriesPerSector = cfb.sectorSize / 4 - 1;
        for (std::size_t i = 0; i < entriesPerSector; ++i) {
            const auto entry = le32(data, base + i * 4);
            if (entry != FreeSector) fatSectorList.push_back(entry);
        }
        difat = le32(data, base + cfb.sectorSize - 4);
    }
    cfb.fat.assign(static_cast<std::size_t>(numFatSectors) * (cfb.sectorSize / 4), FreeSector);
    std::size_t written = 0;
    for (const auto fatSector : fatSectorList) {
        if (fatSector >= fileSectors) throw std::runtime_error("XLS: FAT sector outside the file");
        const auto base = sectorOffset(cfb, fatSector);
        for (std::size_t i = 0; i < cfb.sectorSize / 4; ++i) {
            if (written < cfb.fat.size()) cfb.fat[written] = le32(data, base + i * 4);
            ++written;
        }
    }

    // Mini FAT (chain of mini-FAT sectors through the main FAT).
    if (firstMiniFatSector != EndOfChain && numMiniFatSectors != 0) {
        cfb.miniFat.clear();
        auto sector = firstMiniFatSector;
        for (std::uint32_t i = 0; i < numMiniFatSectors; ++i) {
            if (sector == EndOfChain || sector == FreeSector || sector >= cfb.fat.size())
                throw std::runtime_error("XLS: corrupt mini FAT chain");
            const auto base = sectorOffset(cfb, sector);
            for (std::size_t j = 0; j < cfb.sectorSize / 4; ++j) cfb.miniFat.push_back(le32(data, base + j * 4));
            sector = cfb.fat[sector];
        }
    }

    // Root entry: locate the mini stream (root storage's child stream) so small
    // streams can be read through the mini FAT.
    if (cfb.firstDirSector != EndOfChain) {
        // Directory occupies the chain starting at firstDirSector.
        auto sector = cfb.firstDirSector;
        std::vector<unsigned char> dir;
        for (std::size_t guard = 0; sector != EndOfChain && sector != FreeSector; ++guard) {
            if (guard > cfb.fat.size() || sector >= cfb.fat.size())
                throw std::runtime_error("XLS: corrupt directory chain");
            const auto base = sectorOffset(cfb, sector);
            const auto end = std::min(base + cfb.sectorSize, data.size());
            dir.insert(dir.end(), data.begin() + static_cast<std::ptrdiff_t>(base),
                       data.begin() + static_cast<std::ptrdiff_t>(end));
            sector = cfb.fat[sector];
        }
        if (dir.size() < 128) throw std::runtime_error("XLS: directory stream is empty");
        // Root entry is the first 128-byte entry.
        cfb.rootStart = le32(dir, 0x74);
        cfb.rootSize = le64(dir, 0x78);
    }
    return cfb;
}

std::vector<std::uint32_t> sectorChain(const CfbFile& cfb, std::uint32_t start) {
    std::vector<std::uint32_t> chain;
    auto sector = start;
    for (std::size_t guard = 0; guard < cfb.fat.size() + 1; ++guard) {
        if (sector == EndOfChain) break;
        if (sector == FreeSector || sector >= cfb.fat.size()) throw std::runtime_error("XLS: invalid FAT sector in chain");
        chain.push_back(sector);
        sector = cfb.fat[sector];
    }
    return chain;
}

std::vector<unsigned char> readStreamBytes(const CfbFile& cfb, std::uint32_t start, std::uint64_t size) {
    if (size == 0) return {};
    if (size < cfb.miniStreamCutoff) {
        // Small stream stored in the mini stream via the mini FAT.
        if (cfb.miniFat.empty()) throw std::runtime_error("XLS: mini stream referenced but mini FAT is empty");
        auto miniChain = sectorChain(cfb, start);
        // miniChain entries are mini-sector indices; build from the mini stream.
        std::vector<unsigned char> out;
        out.reserve(static_cast<std::size_t>(size));
        for (const auto miniSector : miniChain) {
            if (miniSector * cfb.miniSectorSize + cfb.miniSectorSize > cfb.rootSize)
                throw std::runtime_error("XLS: mini sector outside mini stream");
            // The mini stream bytes live in the root entry's main-FAT chain.
            const auto rootChain = sectorChain(cfb, cfb.rootStart);
            const auto miniOffset = miniSector * cfb.miniSectorSize;
            const auto mainSectorIdx = miniOffset / cfb.sectorSize;
            const auto withinSector = miniOffset % cfb.sectorSize;
            if (mainSectorIdx >= rootChain.size()) throw std::runtime_error("XLS: mini stream truncated");
            const auto base = sectorOffset(cfb, rootChain[mainSectorIdx]) + withinSector;
            for (std::size_t i = 0; i < cfb.miniSectorSize && base + i < cfb.data.size(); ++i) {
                if (out.size() >= size) break;
                out.push_back(cfb.data[base + i]);
            }
            if (out.size() >= size) break;
        }
        return out;
    }
    const auto chain = sectorChain(cfb, start);
    std::vector<unsigned char> out;
    out.reserve(static_cast<std::size_t>(size));
    for (const auto sector : chain) {
        const auto base = sectorOffset(cfb, sector);
        for (std::size_t i = 0; i < cfb.sectorSize && base + i < cfb.data.size(); ++i) {
            if (out.size() >= size) break;
            out.push_back(cfb.data[base + i]);
        }
        if (out.size() >= size) break;
    }
    if (out.size() < size) throw std::runtime_error("XLS: stream is shorter than its declared size");
    return out;
}

struct DirectoryEntry {
    std::u16string name;
    std::uint8_t type = 0;
    std::uint32_t start = EndOfChain;
    std::uint64_t size = 0;
};

std::vector<DirectoryEntry> readDirectory(const CfbFile& cfb) {
    const auto chain = sectorChain(cfb, cfb.firstDirSector);
    std::vector<DirectoryEntry> entries;
    for (const auto sector : chain) {
        const auto base = sectorOffset(cfb, sector);
        for (std::size_t off = 0; off + 128 <= cfb.sectorSize; off += 128) {
            const auto nameLen = le16(cfb.data, base + off + 0x40);
            if (nameLen < 4 || nameLen > 64) continue;
            std::u16string name;
            // nameLen includes the UTF-16 null terminator.
            for (std::size_t i = 0; i + 2 < nameLen; i += 2)
                name.push_back(static_cast<char16_t>(le16(cfb.data, base + off + i)));
            DirectoryEntry entry;
            entry.name = std::move(name);
            entry.type = cfb.data[base + off + 0x42];
            entry.start = le32(cfb.data, base + off + 0x74);
            entry.size = le64(cfb.data, base + off + 0x78);
            entries.push_back(std::move(entry));
        }
    }
    return entries;
}

std::string utf16ToUtf8(const std::u16string& text) {
    std::string out;
    for (const auto unit : text) {
        std::uint32_t cp = unit;
        if (unit >= 0xD800 && unit <= 0xDBFF) {
            // surrogate pair (unlikely in stream names; handled conservatively)
            cp = 0xFFFD;
        }
        if (cp < 0x80) out.push_back(static_cast<char>(cp));
        else if (cp < 0x800) {
            out.push_back(static_cast<char>(0xC0 | (cp >> 6)));
            out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
        } else {
            out.push_back(static_cast<char>(0xE0 | (cp >> 12)));
            out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
            out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
        }
    }
    return out;
}

// ---------------------------------------------------------------------------
// BIFF8 record layer
// ---------------------------------------------------------------------------

struct Record {
    std::uint16_t type = 0;
    std::string_view payload;
};

// Record ids used by the basic reader.
constexpr std::uint16_t RecordBof = 0x0809;
constexpr std::uint16_t RecordWorkbookBof = 0x0809;
constexpr std::uint16_t RecordEof = 0x000A;
constexpr std::uint16_t RecordBoundsheet = 0x0085;
constexpr std::uint16_t RecordSst = 0x00FC;
constexpr std::uint16_t RecordContinue = 0x003C;
constexpr std::uint16_t RecordDimension = 0x0200;
constexpr std::uint16_t RecordRow = 0x0208;
constexpr std::uint16_t RecordBlank = 0x0201;
constexpr std::uint16_t RecordMulBlank = 0x00BE;
constexpr std::uint16_t RecordNumber = 0x0203;
constexpr std::uint16_t RecordRk = 0x027E;
constexpr std::uint16_t RecordMulRk = 0x00BD;
constexpr std::uint16_t RecordLabelSst = 0x00FD;
constexpr std::uint16_t RecordBoolErr = 0x0205;
constexpr std::uint16_t RecordFormula = 0x0006;
constexpr std::uint16_t RecordString = 0x0207;
constexpr std::uint16_t RecordDate1904 = 0x0022;
constexpr std::uint16_t RecordLabel = 0x0204;
constexpr std::uint16_t RecordMergedCells = 0x00E5;
constexpr std::uint16_t RecordColInfo = 0x007D;

std::uint16_t recU16(const Record& r, std::size_t off) {
    if (off + 2 > r.payload.size()) throw std::runtime_error("XLS: record truncated reading u16");
    const auto* p = reinterpret_cast<const unsigned char*>(r.payload.data());
    return static_cast<std::uint16_t>(p[off]) | (static_cast<std::uint16_t>(p[off + 1]) << 8);
}

std::uint32_t recU32(const Record& r, std::size_t off) {
    if (off + 4 > r.payload.size()) throw std::runtime_error("XLS: record truncated reading u32");
    const auto* p = reinterpret_cast<const unsigned char*>(r.payload.data());
    return static_cast<std::uint32_t>(p[off]) | (static_cast<std::uint32_t>(p[off + 1]) << 8) |
           (static_cast<std::uint32_t>(p[off + 2]) << 16) | (static_cast<std::uint32_t>(p[off + 3]) << 24);
}

double decodeRk(std::uint32_t rk) {
    const bool fNum = (rk & 0x01) != 0;
    const bool fInt = (rk & 0x02) != 0;
    double value;
    if (fInt) {
        // 30-bit two's-complement integer, sign-extended from bit 29.
        std::int32_t v = static_cast<std::int32_t>(rk >> 2);
        if (rk & 0x80000000) v |= 0xC0000000; // sign-extend bits 31..30
        value = static_cast<double>(v);
    } else {
        const std::uint32_t bits = rk & ~0x03u;
        float f;
        std::memcpy(&f, &bits, sizeof(f));
        value = static_cast<double>(f);
    }
    if (fNum) value /= 100.0;
    return value;
}

xlpp::CellError biffErrorCode(std::uint8_t code) {
    switch (code) {
        case 0x00: return xlpp::CellError::Null;
        case 0x07: return xlpp::CellError::DivisionByZero;
        case 0x0F: return xlpp::CellError::Value;
        case 0x17: return xlpp::CellError::Reference;
        case 0x1D: return xlpp::CellError::Name;
        case 0x24: return xlpp::CellError::Number;
        case 0x2A: return xlpp::CellError::NotAvailable;
        case 0x2B: return xlpp::CellError::GettingData;
        default: return xlpp::CellError::Value;
    }
}

// Shared-string continuation reader. Handles the BIFF8 rule that a Unicode
// string split across records starts each continuation record with a fresh
// 1-byte grbit.
struct SstReader {
    const std::vector<std::string_view>& payloads;
    std::size_t rec = 0;
    std::size_t off = 0;

    std::uint8_t readU8() {
        if (rec >= payloads.size()) throw std::runtime_error("XLS: shared string table truncated");
        if (off >= payloads[rec].size()) {
            ++rec; off = 0;
            if (rec >= payloads.size()) throw std::runtime_error("XLS: shared string table truncated");
        }
        return static_cast<std::uint8_t>(payloads[rec][off++]);
    }
    std::uint16_t readU16() { return static_cast<std::uint16_t>(readU8()) | (static_cast<std::uint16_t>(readU8()) << 8); }
    std::uint32_t readU32() { return static_cast<std::uint32_t>(readU16()) | (static_cast<std::uint32_t>(readU16()) << 16); }

    std::u16string readChars(std::size_t count, bool unicode) {
        std::u16string out;
        out.reserve(count);
        std::size_t remaining = count;
        while (remaining > 0) {
            if (rec >= payloads.size()) throw std::runtime_error("XLS: shared string table truncated");
            if (off >= payloads[rec].size()) {
                ++rec; off = 0;
                if (rec >= payloads.size()) throw std::runtime_error("XLS: shared string table truncated");
                if (unicode) {
                    // Continuation records of a Unicode string begin with grbit.
                    ++off;
                    if (off > payloads[rec].size()) throw std::runtime_error("XLS: malformed string continuation");
                }
                continue;
            }
            const auto avail = payloads[rec].size() - off;
            if (unicode) {
                const auto countHere = std::min(remaining, avail / 2);
                for (std::size_t i = 0; i < countHere; ++i) {
                    const auto lo = static_cast<std::uint8_t>(payloads[rec][off++]);
                    const auto hi = static_cast<std::uint8_t>(payloads[rec][off++]);
                    out.push_back(static_cast<char16_t>((static_cast<std::uint16_t>(hi) << 8) | lo));
                }
                remaining -= countHere;
            } else {
                const auto countHere = std::min(remaining, avail);
                for (std::size_t i = 0; i < countHere; ++i)
                    out.push_back(static_cast<char16_t>(static_cast<std::uint8_t>(payloads[rec][off++])));
                remaining -= countHere;
            }
        }
        return out;
    }
};

std::u16string parseSstString(SstReader& reader) {
    std::uint16_t cch = reader.readU16();
    if (cch & 0x8000) cch = static_cast<std::uint16_t>(((cch & 0x7FFF) << 16) | reader.readU16());
    const auto grbit = reader.readU8();
    const bool unicode = (grbit & 0x01) != 0;
    if (grbit & 0x08) reader.readU16(); // cRun (rich-text runs)
    if (grbit & 0x04) reader.readU32(); // cbExtRst
    return reader.readChars(cch, unicode);
}

// ---------------------------------------------------------------------------
// Worksheet materialization
// ---------------------------------------------------------------------------

struct SheetDescriptor {
    std::string name;
    std::uint8_t type = 0;
    std::uint8_t visibility = 0;
};

} // namespace

void readLegacyXls(const std::vector<unsigned char>& bytes, xlpp::Workbook& workbook) {
    const char* stage = "header";
    try {
    if (!isOle2CompoundFile(std::string_view(reinterpret_cast<const char*>(bytes.data()), bytes.size())))
        throw std::runtime_error("XLS: input is not a legacy binary workbook");

    stage = "cfb"; const auto cfb = parseCfb(bytes);
    stage = "directory"; const auto directory = readDirectory(cfb);

    std::string workbookStreamName;
    for (const auto& entry : directory) {
        if (entry.type != 2) continue;
        if (entry.name == u"Workbook" || entry.name == u"Book") { workbookStreamName = utf16ToUtf8(entry.name); break; }
    }
    if (workbookStreamName.empty()) throw std::runtime_error("XLS: no Workbook stream found in compound file");
    stage = "stream"; std::vector<unsigned char> stream;
    for (const auto& entry : directory) {
        if (entry.type != 2) continue;
        if (entry.name == u"Workbook" || entry.name == u"Book") {
            stream = readStreamBytes(cfb, entry.start, entry.size);
            break;
        }
    }
    if (stream.empty()) throw std::runtime_error("XLS: Workbook stream is empty");

    stage = "records";
    // Split the workbook stream into BIFF8 records.
    std::vector<Record> records;
    {
        std::size_t off = 0;
        while (off + 4 <= stream.size()) {
            Record record;
            record.type = static_cast<std::uint16_t>(stream[off]) | (static_cast<std::uint16_t>(stream[off + 1]) << 8);
            const auto len = static_cast<std::size_t>(stream[off + 2]) | (static_cast<std::size_t>(stream[off + 3]) << 8);
            if (off + 4 + len > stream.size()) throw std::runtime_error("XLS: BIFF8 record extends past the stream");
            record.payload = std::string_view(reinterpret_cast<const char*>(stream.data()) + off + 4, len);
            records.push_back(record);
            off += 4 + len;
        }
    }

    stage = "boundsheets";
    // Collect bound sheets (worksheet names in document order).
    std::vector<SheetDescriptor> sheets;
    for (const auto& r : records) {
        if (r.type == RecordBoundsheet && r.payload.size() >= 8) {
            SheetDescriptor descriptor;
            descriptor.type = static_cast<std::uint8_t>(r.payload[5]);
            descriptor.visibility = static_cast<std::uint8_t>(r.payload[4]);
            const auto cch = static_cast<unsigned char>(r.payload[6]);
            const auto nameLen = cch & 0x7F;
            if (cch & 0x80) {
                // 16-bit name
                if (r.payload.size() < 7u + nameLen * 2u) throw std::runtime_error("XLS: truncated sheet name");
                std::u16string name;
                for (std::size_t i = 0; i < nameLen; ++i)
                    name.push_back(static_cast<char16_t>(recU16(r, 7 + i * 2)));
                descriptor.name = utf16ToUtf8(name);
            } else {
                if (r.payload.size() < 7u + nameLen) throw std::runtime_error("XLS: truncated sheet name");
                for (std::size_t i = 0; i < nameLen; ++i)
                    descriptor.name.push_back(static_cast<char>(static_cast<unsigned char>(r.payload[7 + i])));
            }
            sheets.push_back(std::move(descriptor));
        }
    }

    stage = "sst";
    // Shared string table: SST record followed by CONTINUE records.
    std::vector<std::u16string> sharedStrings;
    for (std::size_t i = 0; i < records.size(); ++i) {
        if (records[i].type != RecordSst) continue;
        std::vector<std::string_view> payloads;
        payloads.push_back(records[i].payload);
        std::size_t j = i + 1;
        while (j < records.size() && records[j].type == RecordContinue) {
            payloads.push_back(records[j].payload);
            ++j;
        }
        SstReader reader{payloads};
        reader.readU32(); // cstTotal
        const auto cstUnique = reader.readU32();
        // Hardening: each string consumes at least 3 payload bytes (cch + flags
        // + one char), so a malformed count must not drive a huge reserve().
        std::size_t payloadBytes = 0;
        for (const auto& payload : payloads) payloadBytes += payload.size();
        const auto maxPlausible = payloadBytes / 3 + 1;
        if (static_cast<std::size_t>(cstUnique) > maxPlausible)
            throw std::runtime_error("XLS: shared string count exceeds the table size");
        sharedStrings.reserve(cstUnique);
        for (std::uint32_t k = 0; k < cstUnique; ++k) sharedStrings.push_back(parseSstString(reader));
        break;
    }

    stage = "date1904";
    // Date system: BIFF8 stores a DATE1904 record in workbook globals.
    for (const auto& r : records) {
        if (r.type == RecordDate1904 && r.payload.size() >= 2) {
            workbook.setDate1904(recU16(r, 0) != 0);
            break;
        }
    }

    // Sequentially scan records; worksheet BOFs open a new sheet.
    xlpp::Worksheet* current = nullptr;
    std::size_t sheetIndex = 0;
    bool formulaStringPending = false;
    std::size_t lastFormulaRow = 0;
    std::size_t lastFormulaCol = 0;

    auto makeSheet = [&]() -> xlpp::Worksheet* {
        // Skip macro (1), chartsheet (2) and VB module (6) descriptors; they
        // never carry worksheet cells.
        while (sheetIndex < sheets.size() && sheets[sheetIndex].type != 0x00) ++sheetIndex;
        xlpp::Worksheet* sheet = nullptr;
        if (sheetIndex < sheets.size()) {
            auto& ws = workbook.addWorksheet(sheets[sheetIndex].name);
            sheet = &ws;
            ++sheetIndex;
        }
        return sheet;
    };

    for (const auto& r : records) {
        switch (r.type) {
            case RecordWorkbookBof: {
                if (r.payload.size() < 4) break;
                const auto documentType = recU16(r, 2);
                current = documentType == 0x0010 ? makeSheet() : nullptr;
                break;
            }
            case RecordEof:
                current = nullptr;
                break;
            case RecordRow: {
                if (current == nullptr || r.payload.size() < 6) break;
                const auto row = static_cast<std::size_t>(recU16(r, 0)) + 1;
                auto& dimension = current->rowDimension(row);
                const auto miyRw = recU16(r, 6);
                if (miyRw != 0x00FF) dimension.height = miyRw / 20.0; // twips -> points
                if (r.payload.size() >= 10) dimension.hidden = (recU16(r, 8) & 0x0001) != 0;
                break;
            }
            case RecordColInfo: {
                if (current == nullptr || r.payload.size() < 8) break;
                const auto colFirst = static_cast<std::size_t>(recU16(r, 0)) + 1;
                const auto colLast = static_cast<std::size_t>(recU16(r, 2)) + 1;
                const auto widthUnits = recU16(r, 4); // 1/256 character units
                const auto grbit = recU16(r, 8);
                for (std::size_t col = colFirst; col <= colLast; ++col) {
                    auto& dimension = current->columnDimension(col);
                    dimension.width = widthUnits / 256.0;
                    dimension.hidden = (grbit & 0x0001) != 0;
                }
                break;
            }
            case RecordMergedCells: {
                if (current == nullptr || r.payload.size() < 2) break;
                const auto count = recU16(r, 0);
                for (std::size_t i = 0; i < count; ++i) {
                    const auto offset = 2 + i * 8;
                    if (offset + 8 > r.payload.size()) break;
                    const auto rowFirst = static_cast<std::size_t>(recU16(r, offset)) + 1;
                    const auto rowLast = static_cast<std::size_t>(recU16(r, offset + 2)) + 1;
                    const auto colFirst = static_cast<std::size_t>(recU16(r, offset + 4)) + 1;
                    const auto colLast = static_cast<std::size_t>(recU16(r, offset + 6)) + 1;
                    const auto first = xlpp::CellReference{rowFirst, colFirst}.address();
                    const auto last = xlpp::CellReference{rowLast, colLast}.address();
                    current->mergeCells(first + ":" + last);
                }
                break;
            }
            case RecordNumber: {
                if (current == nullptr || r.payload.size() < 14) break;
                const auto row = static_cast<std::size_t>(recU16(r, 0)) + 1;
                const auto col = static_cast<std::size_t>(recU16(r, 2)) + 1;
                double value;
                std::memcpy(&value, r.payload.data() + 6, sizeof(value));
                current->cell(row, col).setValue(value);
                break;
            }
            case RecordRk: {
                if (current == nullptr || r.payload.size() < 10) break;
                const auto row = static_cast<std::size_t>(recU16(r, 0)) + 1;
                const auto col = static_cast<std::size_t>(recU16(r, 2)) + 1;
                current->cell(row, col).setValue(decodeRk(recU32(r, 6)));
                break;
            }
            case RecordMulRk: {
                if (current == nullptr || r.payload.size() < 6) break;
                const auto row = static_cast<std::size_t>(recU16(r, 0)) + 1;
                const auto colFirst = static_cast<std::size_t>(recU16(r, 2)) + 1;
                const auto colLast = static_cast<std::size_t>(recU16(r, r.payload.size() - 2)) + 1;
                if (colLast < colFirst) break;
                for (std::size_t col = colFirst; col <= colLast; ++col) {
                    const auto offset = 4 + (col - colFirst) * 6;
                    if (offset + 6 > r.payload.size()) break;
                    current->cell(row, col).setValue(decodeRk(recU32(r, offset)));
                }
                break;
            }
            case RecordLabelSst: {
                if (current == nullptr || r.payload.size() < 10) break;
                const auto row = static_cast<std::size_t>(recU16(r, 0)) + 1;
                const auto col = static_cast<std::size_t>(recU16(r, 2)) + 1;
                const auto index = recU32(r, 6);
                if (index < sharedStrings.size())
                    current->cell(row, col).setValue(utf16ToUtf8(sharedStrings[index]));
                break;
            }
            case RecordLabel: {
                if (current == nullptr || r.payload.size() < 8) break;
                const auto row = static_cast<std::size_t>(recU16(r, 0)) + 1;
                const auto col = static_cast<std::size_t>(recU16(r, 2)) + 1;
                std::string text;
                const auto cch = recU16(r, 6);
                if (cch & 0x8000) {
                    if (r.payload.size() < 9u) break;
                    const auto flags = static_cast<unsigned char>(r.payload[8]);
                    if (flags & 0x01) {
                        if (r.payload.size() < 9u + static_cast<std::size_t>(cch & 0x7FFF) * 2u) break;
                        std::u16string name;
                        for (std::size_t i = 0; i < (cch & 0x7FFF); ++i)
                            name.push_back(static_cast<char16_t>(recU16(r, 9 + i * 2)));
                        text = utf16ToUtf8(name);
                    } else {
                        if (r.payload.size() < 9u + static_cast<std::size_t>(cch & 0x7FFF)) break;
                        for (std::size_t i = 0; i < (cch & 0x7FFF); ++i)
                            text.push_back(static_cast<char>(r.payload[9 + i]));
                    }
                } else {
                    if (r.payload.size() < 8u + cch) break;
                    for (std::size_t i = 0; i < cch; ++i) text.push_back(static_cast<char>(r.payload[8 + i]));
                }
                current->cell(row, col).setValue(std::move(text));
                break;
            }
            case RecordBoolErr: {
                if (current == nullptr || r.payload.size() < 8) break;
                const auto row = static_cast<std::size_t>(recU16(r, 0)) + 1;
                const auto col = static_cast<std::size_t>(recU16(r, 2)) + 1;
                const auto value = static_cast<unsigned char>(r.payload[6]);
                const auto isError = static_cast<unsigned char>(r.payload[7]) != 0;
                if (isError) current->cell(row, col).setError(biffErrorCode(value));
                else current->cell(row, col).setValue(value != 0);
                break;
            }
            case RecordFormula: {
                if (current == nullptr || r.payload.size() < 20) break;
                const auto row = static_cast<std::size_t>(recU16(r, 0)) + 1;
                const auto col = static_cast<std::size_t>(recU16(r, 2)) + 1;
                const auto cachedType = static_cast<unsigned char>(r.payload[6]);
                formulaStringPending = false;
                if (cachedType == 0x01 && r.payload.size() >= 9) {
                    current->cell(row, col).setValue(r.payload[8] != 0);
                } else if (cachedType == 0x02 && r.payload.size() >= 9) {
                    current->cell(row, col).setError(biffErrorCode(static_cast<unsigned char>(r.payload[8])));
                } else if (cachedType == 0x03) {
                    formulaStringPending = true;
                    lastFormulaRow = row;
                    lastFormulaCol = col;
                } else {
                    double value;
                    std::memcpy(&value, r.payload.data() + 6, sizeof(value));
                    current->cell(row, col).setValue(value);
                }
                break;
            }
            case RecordString: {
                if (current == nullptr || !formulaStringPending || r.payload.size() < 4) break;
                // String result of the previous formula: XLUnicodeString.
                const auto cch = recU16(r, 0);
                std::string text;
                if (cch & 0x8000) {
                    if (r.payload.size() < 7u) break;
                    const auto flags = static_cast<unsigned char>(r.payload[4]);
                    if (flags & 0x01) {
                        if (r.payload.size() < 5u + static_cast<std::size_t>(cch & 0x7FFF) * 2u) break;
                        std::u16string name;
                        for (std::size_t i = 0; i < (cch & 0x7FFF); ++i)
                            name.push_back(static_cast<char16_t>(recU16(r, 5 + i * 2)));
                        text = utf16ToUtf8(name);
                    } else {
                        if (r.payload.size() < 5u + static_cast<std::size_t>(cch & 0x7FFF)) break;
                        for (std::size_t i = 0; i < (cch & 0x7FFF); ++i) text.push_back(static_cast<char>(r.payload[5 + i]));
                    }
                } else {
                    if (r.payload.size() < 4u + cch) break;
                    for (std::size_t i = 0; i < cch; ++i) text.push_back(static_cast<char>(r.payload[2 + i]));
                }
                if (lastFormulaRow != 0) {
                    current->cell(lastFormulaRow, lastFormulaCol).setValue(std::move(text));
                    lastFormulaRow = 0;
                }
                formulaStringPending = false;
                break;
            }
            default:
                break;
        }
    }
    } catch (const std::exception& e) {
        throw std::runtime_error(std::string("XLS stage ") + stage + ": " + e.what());
    }
}

} // namespace internal
} // namespace xlpp
