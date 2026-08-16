#include "XlsbBinaryReader.h"
#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Worksheet/Worksheet.h>
#include "../Packaging/ZipArchive.h"
#include "../XML/XmlUtilities.h"
#include <cstdint>
#include <cstring>
#include <stdexcept>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

namespace xlpp {
namespace internal {

bool isXlsbPackage(const xlpp::internal::ZipArchive& z) {
    return z.contains("xl/workbook.bin");
}

namespace {

// BIFF12 record ids used by the basic reader.
constexpr std::uint8_t BrtRowHdr = 0x00;
constexpr std::uint8_t BrtCellBlank = 0x01;
constexpr std::uint8_t BrtCellRk = 0x02;
constexpr std::uint8_t BrtCellError = 0x03;
constexpr std::uint8_t BrtCellBool = 0x04;
constexpr std::uint8_t BrtCellReal = 0x05;
constexpr std::uint8_t BrtCellSt = 0x06;
constexpr std::uint8_t BrtCellIsst = 0x07;
constexpr std::uint8_t BrtFmlaString = 0x08;
constexpr std::uint8_t BrtFmlaNum = 0x09;
constexpr std::uint8_t BrtFmlaBool = 0x0A;
constexpr std::uint8_t BrtFmlaError = 0x0B;
constexpr std::uint8_t BrtSSTItem = 0x13;
constexpr std::uint8_t BrtBeginSheetData = 0x91;
constexpr std::uint8_t BrtEndSheetData = 0x94;
constexpr std::uint8_t BrtBeginBundleShs = 0x92;
constexpr std::uint8_t BrtBundleSh = 0x93;
constexpr std::uint8_t BrtBeginSst = 0x9F;
constexpr std::uint8_t BrtEndSst = 0xA0;
constexpr std::uint8_t BrtBeginBook = 0x83;
constexpr std::uint8_t BrtEndBook = 0x84;

struct Record {
    std::uint8_t type = 0;
    std::string_view payload;
};

std::vector<Record> splitRecords(const std::string& bytes) {
    std::vector<Record> records;
    std::size_t offset = 0;
    while (offset + 5 <= bytes.size()) {
        Record record;
        record.type = static_cast<std::uint8_t>(bytes[offset]);
        const auto length = static_cast<std::size_t>(bytes[offset + 1])
            | (static_cast<std::size_t>(bytes[offset + 2]) << 8)
            | (static_cast<std::size_t>(bytes[offset + 3]) << 16)
            | (static_cast<std::size_t>(bytes[offset + 4]) << 24);
        if (offset + 5 + length > bytes.size()) throw std::runtime_error("XLSB: record extends past the stream");
        record.payload = std::string_view(bytes).substr(offset + 5, length);
        records.push_back(record);
        offset += 5 + length;
    }
    return records;
}

std::uint32_t recU32(const std::string_view& payload, std::size_t offset) {
    if (offset + 4 > payload.size()) throw std::runtime_error("XLSB: truncated u32");
    return static_cast<std::uint32_t>(static_cast<std::uint8_t>(payload[offset]))
        | (static_cast<std::uint32_t>(static_cast<std::uint8_t>(payload[offset + 1])) << 8)
        | (static_cast<std::uint32_t>(static_cast<std::uint8_t>(payload[offset + 2])) << 16)
        | (static_cast<std::uint32_t>(static_cast<std::uint8_t>(payload[offset + 3])) << 24);
}

std::uint16_t recU16(const std::string_view& payload, std::size_t offset) {
    if (offset + 2 > payload.size()) throw std::runtime_error("XLSB: truncated u16");
    return static_cast<std::uint16_t>(static_cast<std::uint8_t>(payload[offset]))
        | (static_cast<std::uint16_t>(static_cast<std::uint8_t>(payload[offset + 1])) << 8);
}

// XLWideString: cch (4) followed by UTF-16LE characters.
std::pair<std::string, std::size_t> readWideString(const std::string_view& payload, std::size_t offset) {
    const auto cch = recU32(payload, offset);
    offset += 4;
    if (offset + cch * 2 > payload.size()) throw std::runtime_error("XLSB: wide string extends past the record");
    std::string utf8;
    utf8.reserve(cch);
    for (std::uint32_t i = 0; i < cch; ++i) {
        const auto unit = static_cast<std::uint16_t>(static_cast<std::uint8_t>(payload[offset + i * 2]))
            | (static_cast<std::uint16_t>(static_cast<std::uint8_t>(payload[offset + i * 2 + 1])) << 8);
        if (unit < 0x80) utf8.push_back(static_cast<char>(unit));
        else if (unit < 0x800) {
            utf8.push_back(static_cast<char>(0xC0 | (unit >> 6)));
            utf8.push_back(static_cast<char>(0x80 | (unit & 0x3F)));
        } else {
            utf8.push_back(static_cast<char>(0xE0 | (unit >> 12)));
            utf8.push_back(static_cast<char>(0x80 | ((unit >> 6) & 0x3F)));
            utf8.push_back(static_cast<char>(0x80 | (unit & 0x3F)));
        }
    }
    return {std::move(utf8), offset + cch * 2};
}

// XLNullableWideString: flags (1) then, when not null, an XLWideString.
std::pair<std::string, std::size_t> readNullableWideString(const std::string_view& payload, std::size_t offset) {
    if (offset + 1 > payload.size()) throw std::runtime_error("XLSB: truncated nullable string flags");
    const auto flags = static_cast<std::uint8_t>(payload[offset]);
    ++offset;
    if (flags & 0x01) return {{}, offset};
    return readWideString(payload, offset);
}

double decodeRk(std::uint32_t rk) {
    const bool fNum = (rk & 0x01) != 0;
    const bool fInt = (rk & 0x02) != 0;
    double value;
    if (fInt) {
        std::int32_t v = static_cast<std::int32_t>(rk >> 2);
        if (rk & 0x80000000) v |= 0xC0000000;
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

} // namespace

void readXlsb(const xlpp::internal::ZipArchive& z, xlpp::Workbook& workbook) {
    if (!isXlsbPackage(z)) throw std::runtime_error("XLSB: package has no xl/workbook.bin");

    // Shared strings table.
    std::vector<std::string> sharedStrings;
    if (z.contains("xl/sharedStrings.bin")) {
        for (const auto& record : splitRecords(z.get("xl/sharedStrings.bin"))) {
            if (record.type != BrtSSTItem) continue;
            const auto [text, after] = readWideString(record.payload, 0);
            sharedStrings.push_back(text);
            (void)after;
        }
    }

    // Workbook: sheet names + relationship ids.
    std::vector<std::pair<std::string, std::string>> sheets; // (name, relId)
    for (const auto& record : splitRecords(z.get("xl/workbook.bin"))) {
        if (record.type != BrtBundleSh) continue;
        std::size_t offset = 0;
        ++offset; // hsState
        offset += 4; // iTabID
        const auto [relId, afterRel] = readNullableWideString(record.payload, offset);
        const auto [name, afterName] = readWideString(record.payload, afterRel);
        (void)afterName;
        sheets.push_back({name, relId});
    }

    // Map relationship id -> target part from workbook.bin.rels.
    std::unordered_map<std::string, std::string> sheetTargets;
    if (z.contains("xl/_rels/workbook.bin.rels")) {
        const auto rels = z.get("xl/_rels/workbook.bin.rels");
        for (const auto& node : internal::tags(rels, "Relationship")) {
            const auto id = internal::attribute(node, "Id");
            const auto target = internal::attribute(node, "Target");
            if (!id.empty()) sheetTargets[id] = target;
        }
    }

    for (const auto& [name, relId] : sheets) {
        if (name.empty()) continue;
        auto& ws = workbook.addWorksheet(name);
        const auto targetIt = sheetTargets.find(relId);
        if (targetIt == sheetTargets.end()) continue;
        std::string part = targetIt->second;
        if (part.rfind("/", 0) == 0) part.erase(part.begin());
        else if (part.rfind("xl/", 0) != 0) part = "xl/" + part;
        if (!z.contains(part)) continue;

        std::size_t currentRow = 0;
        for (const auto& record : splitRecords(z.get(part))) {
            switch (record.type) {
                case BrtRowHdr: {
                    if (record.payload.size() < 4) break;
                    currentRow = static_cast<std::size_t>(recU32(record.payload, 0) & 0xFFFFF) + 1;
                    break;
                }
                case BrtCellBlank: {
                    if (record.payload.size() < 4) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    if (currentRow == 0) currentRow = 1;
                    ws.cell(currentRow, column);
                    break;
                }
                case BrtCellRk: {
                    if (record.payload.size() < 8) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    ws.cell(currentRow, column).setValue(decodeRk(recU32(record.payload, 4)));
                    break;
                }
                case BrtCellReal: {
                    if (record.payload.size() < 12) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    double value;
                    std::memcpy(&value, record.payload.data() + 4, sizeof(value));
                    ws.cell(currentRow, column).setValue(value);
                    break;
                }
                case BrtCellBool: {
                    if (record.payload.size() < 5) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    ws.cell(currentRow, column).setValue(record.payload[4] != 0);
                    break;
                }
                case BrtCellError: {
                    if (record.payload.size() < 5) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    ws.cell(currentRow, column).setError(biffErrorCode(static_cast<std::uint8_t>(record.payload[4])));
                    break;
                }
                case BrtCellIsst: {
                    if (record.payload.size() < 8) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    const auto index = recU32(record.payload, 4);
                    if (index < sharedStrings.size())
                        ws.cell(currentRow, column).setValue(sharedStrings[index]);
                    break;
                }
                case BrtCellSt: {
                    if (record.payload.size() < 6) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    const auto [text, after] = readWideString(record.payload, 4);
                    ws.cell(currentRow, column).setValue(std::move(text));
                    (void)after;
                    break;
                }
                case BrtFmlaNum: {
                    if (record.payload.size() < 12) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    double value;
                    std::memcpy(&value, record.payload.data() + 4, sizeof(value));
                    ws.cell(currentRow, column).setValue(value);
                    break;
                }
                case BrtFmlaBool: {
                    if (record.payload.size() < 5) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    ws.cell(currentRow, column).setValue(record.payload[4] != 0);
                    break;
                }
                case BrtFmlaError: {
                    if (record.payload.size() < 5) break;
                    const auto column = static_cast<std::size_t>(recU16(record.payload, 0) & 0x3FFF) + 1;
                    ws.cell(currentRow, column).setError(biffErrorCode(static_cast<std::uint8_t>(record.payload[4])));
                    break;
                }
                default:
                    break;
            }
        }
    }
}

} // namespace internal
} // namespace xlpp
