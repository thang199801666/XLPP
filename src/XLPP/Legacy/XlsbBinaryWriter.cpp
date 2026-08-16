#include "XlsbBinaryWriter.h"
#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Cell/Cell.h>
#include <XLPP/Cell/CellReference.h>
#include <XLPP/Cell/DateTime.h>
#include "../Packaging/ZipArchive.h"
#include "../XML/XmlUtilities.h"
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <sstream>
#include <stdexcept>
#include <string>
#include <unordered_map>
#include <vector>

namespace xlpp {
namespace internal {

namespace {

// BIFF12 record ids (1-byte record type + 4-byte length).
constexpr std::uint8_t BrtRowHdr = 0x00;
constexpr std::uint8_t BrtCellBlank = 0x01;
constexpr std::uint8_t BrtCellRk = 0x02;
constexpr std::uint8_t BrtCellError = 0x03;
constexpr std::uint8_t BrtCellBool = 0x04;
constexpr std::uint8_t BrtCellReal = 0x05;
constexpr std::uint8_t BrtCellIsst = 0x07;
constexpr std::uint8_t BrtSSTItem = 0x13;
constexpr std::uint8_t BrtBeginSheetData = 0x91;
constexpr std::uint8_t BrtEndSheetData = 0x94;
constexpr std::uint8_t BrtBeginBundleShs = 0x92;
constexpr std::uint8_t BrtBundleSh = 0x93;
constexpr std::uint8_t BrtEndBundleShs = 0x95;
constexpr std::uint8_t BrtBeginSst = 0x9F;
constexpr std::uint8_t BrtEndSst = 0xA0;
constexpr std::uint8_t BrtBeginBook = 0x83;
constexpr std::uint8_t BrtEndBook = 0x84;

void putU16(std::vector<unsigned char>& out, std::uint16_t value) {
    out.push_back(static_cast<unsigned char>(value & 0xFF));
    out.push_back(static_cast<unsigned char>((value >> 8) & 0xFF));
}
void putU32(std::vector<unsigned char>& out, std::uint32_t value) {
    putU16(out, static_cast<std::uint16_t>(value & 0xFFFF));
    putU16(out, static_cast<std::uint16_t>((value >> 16) & 0xFFFF));
}
void putDouble(std::vector<unsigned char>& out, double value) {
    const auto begin = out.size();
    out.resize(begin + 8);
    std::memcpy(out.data() + begin, &value, sizeof(value));
}
void putRecord(std::vector<unsigned char>& out, std::uint8_t type, const std::vector<unsigned char>& payload) {
    out.push_back(type);
    putU32(out, static_cast<std::uint32_t>(payload.size()));
    out.insert(out.end(), payload.begin(), payload.end());
}
void putWideString(std::vector<unsigned char>& out, const std::string& text) {
    putU32(out, static_cast<std::uint32_t>(text.size()));
    for (const unsigned char ch : text) { out.push_back(ch); out.push_back(0); }
}
void putNullableWideString(std::vector<unsigned char>& out, const std::string& text) {
    out.push_back(0); // flags: present
    putWideString(out, text);
}

// workbook.bin
std::vector<unsigned char> xlsbWorkbookBin(const xlpp::Workbook& workbook,
                                           const std::vector<std::string>& sheetNames) {
    std::vector<unsigned char> out;
    putRecord(out, BrtBeginBook, {});
    putRecord(out, BrtBeginBundleShs, {});
    for (std::size_t i = 0; i < sheetNames.size(); ++i) {
        std::vector<unsigned char> bundle;
        bundle.push_back(0);                     // hsState: visible
        putU32(bundle, static_cast<std::uint32_t>(i + 1)); // iTabID
        putNullableWideString(bundle, "rId" + std::to_string(i + 1));
        putWideString(bundle, sheetNames[i]);
        putRecord(out, BrtBundleSh, bundle);
    }
    putRecord(out, BrtEndBundleShs, {});
    putRecord(out, BrtEndBook, {});
    return out;
}

// sharedStrings.bin
std::vector<unsigned char> xlsbSharedStringsBin(const std::vector<std::string>& strings) {
    std::vector<unsigned char> out;
    std::vector<unsigned char> begin;
    putU32(begin, static_cast<std::uint32_t>(strings.size()));
    putU32(begin, static_cast<std::uint32_t>(strings.size()));
    putRecord(out, BrtBeginSst, begin);
    for (const auto& text : strings) {
        std::vector<unsigned char> item;
        putWideString(item, text);
        putRecord(out, BrtSSTItem, item);
    }
    putRecord(out, BrtEndSst, {});
    return out;
}

// sheetN.bin
std::vector<unsigned char> xlsbSheetBin(const xlpp::Worksheet& sheet, bool date1904,
                                        const std::vector<std::string>& strings,
                                        const std::unordered_map<std::string, std::uint32_t>& stringIndex) {
    std::vector<unsigned char> out;
    putRecord(out, BrtBeginSheetData, {});
    std::size_t currentRow = 0;
    for (const auto& [key, cell] : sheet.cells()) {
        if (cell.empty() && !cell.hasNonDefaultStyle()) continue;
        if (cell.row() != currentRow) {
            std::vector<unsigned char> rowHdr;
            putU32(rowHdr, static_cast<std::uint32_t>(cell.row() - 1));
            putU16(rowHdr, 0); putU16(rowHdr, 0);
            putRecord(out, BrtRowHdr, rowHdr);
            currentRow = cell.row();
        }
        std::vector<unsigned char> payload;
        putU16(payload, static_cast<std::uint16_t>(cell.column() - 1)); // col
        putU16(payload, 0xFFFF);                                       // iStyleRef: none
        if (const auto* number = std::get_if<double>(&cell.value())) {
            putDouble(payload, *number);
            putRecord(out, BrtCellReal, payload);
        } else if (const auto* text = std::get_if<std::string>(&cell.value())) {
            const auto it = stringIndex.find(*text);
            putU32(payload, it == stringIndex.end() ? 0 : it->second);
            putRecord(out, BrtCellIsst, payload);
        } else if (const auto* boolean = std::get_if<bool>(&cell.value())) {
            payload.push_back(*boolean ? 1 : 0);
            putRecord(out, BrtCellBool, payload);
        } else if (const auto* error = std::get_if<xlpp::CellError>(&cell.value())) {
            std::uint8_t code = 0x0F;
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
            putRecord(out, BrtCellError, payload);
        } else if (const auto* date = std::get_if<xlpp::DateTime>(&cell.value())) {
            putDouble(payload, xlpp::toExcelSerial(*date, date1904));
            putRecord(out, BrtCellReal, payload);
        } else {
            putRecord(out, BrtCellBlank, payload);
        }
    }
    putRecord(out, BrtEndSheetData, {});
    return out;
}

} // namespace

void writeXlsbPackage(const xlpp::Workbook& workbook,
                      const std::filesystem::path& path,
                      int compressionLevel) {
    const auto sheetNames = workbook.sheetNames();
    if (sheetNames.empty()) throw std::runtime_error("XLSB: workbook has no worksheets");

    // Shared strings.
    std::vector<std::string> strings;
    std::unordered_map<std::string, std::uint32_t> stringIndex;
    for (const auto& name : sheetNames) {
        const auto* sheet = workbook.worksheet(name);
        if (!sheet) continue;
        for (const auto& [key, cell] : sheet->cells()) {
            if (const auto* text = std::get_if<std::string>(&cell.value())) {
                const auto [it, inserted] = stringIndex.emplace(*text, static_cast<std::uint32_t>(strings.size()));
                if (inserted) strings.push_back(*text);
            }
        }
    }

    internal::ZipArchive z;
    z.setCompressionLevel(compressionLevel);

    std::ostringstream contentTypes;
    contentTypes << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">)"
        << R"(<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>)"
        << R"(<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>)"
        << R"(<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>)";
    for (std::size_t i = 0; i < sheetNames.size(); ++i)
        contentTypes << "<Override PartName=\"/xl/worksheets/sheet" << (i + 1)
                     << ".bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>";
    if (!strings.empty())
        contentTypes << R"(<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>)";
    contentTypes << "</Types>";
    z.add("[Content_Types].xml", contentTypes.str());

    z.add("_rels/.rels",
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)"
        R"(<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>)");

    const auto workbookBin = xlsbWorkbookBin(workbook, sheetNames);
    z.add("xl/workbook.bin", std::string(workbookBin.begin(), workbookBin.end()));

    std::ostringstream wbRels;
    wbRels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)";
    if (!strings.empty())
        wbRels << R"(<Relationship Id="rIdSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>)";
    for (std::size_t i = 0; i < sheetNames.size(); ++i)
        wbRels << "<Relationship Id=\"rId" << (i + 1)
               << "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet"
               << (i + 1) << ".bin\"/>";
    wbRels << "</Relationships>";
    z.add("xl/_rels/workbook.bin.rels", wbRels.str());

    if (!strings.empty()) {
        const auto sst = xlsbSharedStringsBin(strings);
        z.add("xl/sharedStrings.bin", std::string(sst.begin(), sst.end()));
    }

    for (std::size_t i = 0; i < sheetNames.size(); ++i) {
        const auto* sheet = workbook.worksheet(sheetNames[i]);
        if (!sheet) continue;
        const auto sheetBin = xlsbSheetBin(*sheet, workbook.date1904(), strings, stringIndex);
        z.add("xl/worksheets/sheet" + std::to_string(i + 1) + ".bin",
              std::string(sheetBin.begin(), sheetBin.end()));
    }

    z.save(path, {});
}

} // namespace internal
} // namespace xlpp
