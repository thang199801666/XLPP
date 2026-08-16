#include "XlsBinaryWriter.h"
#include "XlsBinaryReader.h"
#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Cell/Cell.h>
#include <XLPP/Cell/CellReference.h>
#include <algorithm>
#include <cstring>
#include <stdexcept>
#include <string>
#include <unordered_map>
#include <vector>
#include <cstdint>

namespace xlpp {
namespace internal {

namespace {

void putU16(std::vector<unsigned char>& out, std::uint16_t value) {
    out.push_back(static_cast<unsigned char>(value & 0xFF));
    out.push_back(static_cast<unsigned char>((value >> 8) & 0xFF));
}

void putU32(std::vector<unsigned char>& out, std::uint32_t value) {
    putU16(out, static_cast<std::uint16_t>(value & 0xFFFF));
    putU16(out, static_cast<std::uint16_t>((value >> 16) & 0xFFFF));
}

void putRecord(std::vector<unsigned char>& out, std::uint16_t type, const std::vector<unsigned char>& payload) {
    putU16(out, type);
    putU16(out, static_cast<std::uint16_t>(payload.size()));
    out.insert(out.end(), payload.begin(), payload.end());
}

void putDouble(std::vector<unsigned char>& payload, double value) {
    const auto begin = payload.size();
    payload.resize(begin + 8);
    std::memcpy(payload.data() + begin, &value, sizeof(value));
}

// Appends a BIFF8 XLUnicodeString (compressed or 16-bit) to `payload`.
void putXlUnicodeString(std::vector<unsigned char>& payload, const std::string& text) {
    const auto begin = payload.size();
    // Fast path: if every character is in the Latin-1 range, use compressed.
    bool compressed = true;
    for (const unsigned char ch : text) {
        if (ch > 0x7F) { compressed = false; break; }
    }
    const auto charCount = text.size();
    if (charCount > 0x7FFF) throw std::runtime_error("XLS: string too long for BIFF8");
    putU16(payload, static_cast<std::uint16_t>(charCount));
    if (compressed) {
        payload.push_back(0x00);
        payload.insert(payload.end(), text.begin(), text.end());
    } else {
        payload.push_back(0x01);
        for (const unsigned char ch : text) {
            putU16(payload, ch);
        }
    }
    (void)begin;
}

// Container builders ----------------------------------------------------------

std::vector<unsigned char> buildCfb(const std::vector<unsigned char>& workbookStream) {
    constexpr std::size_t sectorSize = 512;
    const auto streamSectors = (workbookStream.size() + sectorSize - 1) / sectorSize;
    // Sector layout: 0 = FAT, 1 = directory, 2..2+streamSectors-1 = workbook.
    const auto totalSectors = 2 + streamSectors;

    std::vector<unsigned char> file(512, 0);
    const unsigned char magic[8] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
    std::memcpy(file.data(), magic, 8);
    const auto put16 = [&](std::size_t off, std::uint16_t v) { file[off] = static_cast<unsigned char>(v & 0xFF); file[off + 1] = static_cast<unsigned char>((v >> 8) & 0xFF); };
    const auto put32 = [&](std::size_t off, std::uint32_t v) { put16(off, static_cast<std::uint16_t>(v & 0xFFFF)); put16(off + 2, static_cast<std::uint16_t>((v >> 16) & 0xFFFF)); };
    put16(0x18, 0x003E); // minor version
    put16(0x1A, 0x0003); // major version
    put16(0x1C, 0xFFFE); // byte order
    file[0x1E] = 9;      // sector shift
    file[0x1F] = 6;      // mini sector shift
    put32(0x2C, 1);      // number of FAT sectors
    put32(0x30, 1);      // first directory sector
    put32(0x34, 0);      // transaction signature
    put32(0x38, 0);      // mini stream cutoff (0: all streams via FAT)
    put32(0x3C, 0xFFFFFFFE); // first mini FAT sector
    put32(0x40, 0);      // number of mini FAT sectors
    put32(0x44, 0xFFFFFFFE); // first DIFAT sector
    put32(0x48, 0);      // number of DIFAT sectors
    for (std::size_t i = 0; i < 109; ++i) put32(0x4C + i * 4, i == 0 ? 0 : 0xFFFFFFFF);

    file.resize(512 + sectorSize * totalSectors, 0);
    // FAT sector (sector 0).
    auto fat = reinterpret_cast<std::uint32_t*>(file.data() + 512);
    for (std::size_t i = 0; i < totalSectors; ++i) fat[i] = 0xFFFFFFFF; // FREE
    fat[0] = 0xFFFFFFFE; // FAT sector chain end
    fat[1] = 0xFFFFFFFE; // directory sector chain end
    for (std::size_t i = 2; i < 2 + streamSectors; ++i) fat[i] = (i + 1 == 2 + streamSectors) ? 0xFFFFFFFE : static_cast<std::uint32_t>(i + 1);
    if (streamSectors == 0) fat[2] = 0xFFFFFFFE;

    // Directory sector (sector 1): entry 0 = Root Entry, entry 1 = Workbook stream.
    const auto entry0 = 512 + sectorSize;
    const auto entry1 = entry0 + 128;
    const auto writeName = [&](std::size_t base, const char16_t* name, std::size_t cch) {
        for (std::size_t i = 0; i < cch; ++i) { put16(base + i * 2, static_cast<std::uint16_t>(name[i])); }
        put16(base + 0x40, static_cast<std::uint16_t>((cch + 1) * 2));
    };
    writeName(entry0, u"Root Entry", 11);
    file[entry0 + 0x42] = 5; // root storage
    file[entry0 + 0x43] = 1;
    put32(entry0 + 0x44, 0xFFFFFFFF);
    put32(entry0 + 0x48, 0xFFFFFFFF);
    put32(entry0 + 0x4C, 1); // child = Workbook entry
    put32(entry0 + 0x60, 0);
    put32(entry0 + 0x74, 0xFFFFFFFE);
    // size = 0 (no mini stream)

    writeName(entry1, u"Workbook", 8);
    file[entry1 + 0x42] = 2; // stream
    file[entry1 + 0x43] = 1;
    put32(entry1 + 0x44, 0xFFFFFFFF);
    put32(entry1 + 0x48, 0xFFFFFFFF);
    put32(entry1 + 0x4C, 0xFFFFFFFF);
    put32(entry1 + 0x60, 0);
    put32(entry1 + 0x74, 2); // starting sector
    for (int i = 0; i < 8; ++i)
        file[entry1 + 0x78 + static_cast<std::size_t>(i)] = static_cast<unsigned char>((workbookStream.size() >> (8 * i)) & 0xFF);

    // Workbook stream bytes.
    std::memcpy(file.data() + 512 + sectorSize * 2, workbookStream.data(), workbookStream.size());
    return file;
}

} // namespace

void writeLegacyXls(const xlpp::Workbook& workbook, std::vector<unsigned char>& out) {
    const auto sheetNames = workbook.sheetNames();
    // Collect shared strings.
    std::vector<std::string> strings;
    std::unordered_map<std::string, std::uint32_t> stringIndex;
    for (const auto& name : sheetNames) {
        const auto* sheet = workbook.worksheet(name);
        if (sheet == nullptr) continue;
        for (const auto& [key, cell] : sheet->cells()) {
            if (const auto* text = std::get_if<std::string>(&cell.value())) {
                const auto [it, inserted] = stringIndex.emplace(*text, static_cast<std::uint32_t>(strings.size()));
                if (inserted) strings.push_back(*text);
            }
        }
    }

    std::vector<unsigned char> stream;

    // Workbook globals.
    auto bofGlobals = std::vector<unsigned char>();
    putU16(bofGlobals, 0x0600); putU16(bofGlobals, 0x0005);
    putU16(bofGlobals, 0x0DBB); putU16(bofGlobals, 0x07CC);
    putU32(bofGlobals, 0); putU32(bofGlobals, 0);
    putRecord(stream, 0x0809, bofGlobals);

    if (workbook.date1904()) {
        std::vector<unsigned char> date1904; putU16(date1904, 1);
        putRecord(stream, 0x0022, date1904);
    }

    // BOUNDSHEET entries (worksheet names); positions patched below.
    struct SheetPos { std::size_t recordOffset; std::size_t posOffset; };
    std::vector<SheetPos> sheetPositions;
    for (const auto& name : sheetNames) {
        const auto* sheet = workbook.worksheet(name);
        if (sheet == nullptr) continue;
        const auto& name = sheet->name();
        if (name.size() > 255) throw std::runtime_error("XLS: sheet name too long");
        const auto recordOffset = stream.size();
        std::vector<unsigned char> boundsheet(4 + 3, 0);
        boundsheet[4] = 0; // hidden state: visible
        boundsheet[5] = 0; // sheet type: worksheet
        boundsheet[6] = static_cast<unsigned char>(name.size());
        boundsheet.insert(boundsheet.end(), name.begin(), name.end());
        putRecord(stream, 0x0085, boundsheet);
        // The 4-byte lbPlyPos starts at recordOffset + 4.
        sheetPositions.push_back({recordOffset, recordOffset + 4});
    }

    // Shared string table.
    if (!strings.empty()) {
        std::vector<unsigned char> sst;
        putU32(sst, static_cast<std::uint32_t>(strings.size())); // cstTotal
        putU32(sst, static_cast<std::uint32_t>(strings.size())); // cstUnique
        for (const auto& text : strings) putXlUnicodeString(sst, text);
        putRecord(stream, 0x00FC, sst);
    }

    putRecord(stream, 0x000A, {}); // EOF (globals)

    // Worksheet sections.
    std::size_t sheetOrdinal = 0;
    for (const auto& name : sheetNames) {
        const auto* sheet = workbook.worksheet(name);
        if (sheet == nullptr) continue;
        const auto sheetBofPos = stream.size();
        // Patch BOUNDSHEET lbPlyPos.
        const auto& pos = sheetPositions[sheetOrdinal++];
        stream[pos.posOffset] = static_cast<unsigned char>(sheetBofPos & 0xFF);
        stream[pos.posOffset + 1] = static_cast<unsigned char>((sheetBofPos >> 8) & 0xFF);
        stream[pos.posOffset + 2] = static_cast<unsigned char>((sheetBofPos >> 16) & 0xFF);
        stream[pos.posOffset + 3] = static_cast<unsigned char>((sheetBofPos >> 24) & 0xFF);
        (void)pos;

        auto bofWs = std::vector<unsigned char>();
        putU16(bofWs, 0x0600); putU16(bofWs, 0x0010);
        putU16(bofWs, 0x0DBB); putU16(bofWs, 0x07CC);
        putU32(bofWs, 0); putU32(bofWs, 0);
        putRecord(stream, 0x0809, bofWs);

        std::size_t minRow = 0, maxRow = 0, minCol = 0, maxCol = 0;
        for (const auto& [key, cell] : sheet->cells()) {
            if (cell.empty() && !cell.hasNonDefaultStyle()) continue;
            if (minRow == 0) { minRow = maxRow = cell.row(); minCol = maxCol = cell.column(); }
            else {
                minRow = std::min(minRow, cell.row()); maxRow = std::max(maxRow, cell.row());
                minCol = std::min(minCol, cell.column()); maxCol = std::max(maxCol, cell.column());
            }
        }
        if (minRow != 0) {
            std::vector<unsigned char> dim;
            putU32(dim, static_cast<std::uint32_t>(minRow - 1));
            putU32(dim, static_cast<std::uint32_t>(maxRow));
            putU16(dim, static_cast<std::uint16_t>(minCol - 1));
            putU16(dim, static_cast<std::uint16_t>(maxCol));
            putU16(dim, 0);
            putRecord(stream, 0x0200, dim);
        }

        // Row heights.
        for (const auto& [row, dim] : sheet->rowDimensions()) {
            if (!dim.height && !dim.hidden) continue;
            std::vector<unsigned char> rowRec;
            putU16(rowRec, static_cast<std::uint16_t>(row - 1));
            putU16(rowRec, 0);
            putU16(rowRec, static_cast<std::uint16_t>(sheet->cells().empty() ? 0 : 16383));
            putU16(rowRec, dim.height ? static_cast<std::uint16_t>(*dim.height * 20.0 + 0.5) : static_cast<std::uint16_t>(0x00FF));
            putU16(rowRec, 0);
            putU16(rowRec, dim.hidden ? 1 : 0);
            putU16(rowRec, 0);
            putRecord(stream, 0x0208, rowRec);
        }

        // Column widths.
        for (const auto& [col, dim] : sheet->columnDimensions()) {
            if (!dim.width && !dim.hidden) continue;
            std::vector<unsigned char> colInfo;
            putU16(colInfo, static_cast<std::uint16_t>(col - 1));
            putU16(colInfo, static_cast<std::uint16_t>(col - 1));
            putU16(colInfo, dim.width ? static_cast<std::uint16_t>(*dim.width * 256.0 + 0.5) : static_cast<std::uint16_t>(0));
            putU16(colInfo, 0);
            putU16(colInfo, dim.hidden ? 1 : 0);
            putU16(colInfo, 0);
            putRecord(stream, 0x007D, colInfo);
        }

        // Cells.
        for (const auto& [key, cell] : sheet->cells()) {
            if (cell.empty() && !cell.hasNonDefaultStyle()) continue;
            const auto row = static_cast<std::uint16_t>(cell.row() - 1);
            const auto col = static_cast<std::uint16_t>(cell.column() - 1);
            std::vector<unsigned char> payload;
            putU16(payload, row); putU16(payload, col); putU16(payload, 0);
            if (const auto* number = std::get_if<double>(&cell.value())) {
                putDouble(payload, *number);
                putRecord(stream, 0x0203, payload);
            } else if (const auto* text = std::get_if<std::string>(&cell.value())) {
                const auto it = stringIndex.find(*text);
                putU32(payload, it == stringIndex.end() ? 0 : it->second);
                putRecord(stream, 0x00FD, payload);
            } else if (const auto* boolean = std::get_if<bool>(&cell.value())) {
                payload.push_back(*boolean ? 1 : 0);
                payload.push_back(0);
                putRecord(stream, 0x0205, payload);
            } else if (const auto* error = std::get_if<xlpp::CellError>(&cell.value())) {
                std::uint8_t code = 0x0F; // default #VALUE!
                switch (*error) {
                    case xlpp::CellError::Null: code = 0x00; break;
                    case xlpp::CellError::DivisionByZero: code = 0x07; break;
                    case xlpp::CellError::Value: code = 0x0F; break;
                    case xlpp::CellError::Reference: code = 0x17; break;
                    case xlpp::CellError::Name: code = 0x1D; break;
                    case xlpp::CellError::Number: code = 0x24; break;
                    case xlpp::CellError::NotAvailable: code = 0x2A; break;
                    case xlpp::CellError::GettingData: code = 0x2B; break;
                }
                payload.push_back(code);
                payload.push_back(1);
                putRecord(stream, 0x0205, payload);
            } else if (const auto* date = std::get_if<xlpp::DateTime>(&cell.value())) {
                putDouble(payload, xlpp::toExcelSerial(*date, workbook.date1904()));
                putRecord(stream, 0x0203, payload);
            }
        }

        // Merged ranges.
        const auto& merged = sheet->mergedRanges();
        if (!merged.empty()) {
            std::vector<unsigned char> mergedRec;
            putU16(mergedRec, static_cast<std::uint16_t>(merged.size()));
            for (const auto& range : merged) {
                const auto colon = range.find(':');
                if (colon == std::string::npos) continue;
                const auto first = xlpp::CellReference::parse(range.substr(0, colon));
                const auto last = xlpp::CellReference::parse(range.substr(colon + 1));
                putU16(mergedRec, static_cast<std::uint16_t>(first.row - 1));
                putU16(mergedRec, static_cast<std::uint16_t>(last.row - 1));
                putU16(mergedRec, static_cast<std::uint16_t>(first.column - 1));
                putU16(mergedRec, static_cast<std::uint16_t>(last.column - 1));
            }
            putRecord(stream, 0x00E5, mergedRec);
        }

        putRecord(stream, 0x000A, {}); // EOF (worksheet)
    }

    out = buildCfb(stream);
}

} // namespace internal
} // namespace xlpp
