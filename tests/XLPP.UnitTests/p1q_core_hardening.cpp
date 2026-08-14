#include "XML/XmlPullReader.h"
#include "Packaging/ZipArchive.h"
#include "Vba/VbaProjectBinary.h"
#include <XLPP/Streaming/StreamingWorkbookReader.h>

#include <algorithm>
#include <chrono>
#include <filesystem>
#include <iostream>
#include <stdexcept>
#include <string>
#include <string_view>

namespace {

int failures = 0;

void check(bool value, const char* message) {
    if (!value) {
        ++failures;
        std::cerr << "FAIL: " << message << '\n';
    }
}

template <class F>
void checkThrows(F&& fn, const char* message) {
    try {
        fn();
        check(false, message);
    } catch (const std::exception&) {
        check(true, message);
    }
}

std::filesystem::path tempPath(std::string_view stem) {
    const auto tick = std::chrono::high_resolution_clock::now().time_since_epoch().count();
    return std::filesystem::temp_directory_path() /
           (std::string(stem) + "_" + std::to_string(tick) + ".xlsx");
}

void testPullReaderHardening() {
    const std::string xml = R"(<root><row note="1 > 0"><c>42</c></row><row id="2" /></root>)";
    std::size_t offset = 0;
    xlpp::internal::XmlPullReader reader([&](unsigned char* out, std::size_t capacity) {
        const auto count = std::min<std::size_t>({capacity, 7u, xml.size() - offset});
        if (!count) return std::size_t{0};
        std::copy_n(reinterpret_cast<const unsigned char*>(xml.data() + offset), count, out);
        offset += count;
        return count;
    }, 1024);

    const auto first = reader.nextElement("row");
    check(first.find("1 > 0") != std::string_view::npos, "quoted '>' does not terminate opening tag");
    const auto second = reader.nextElement("row");
    check(second.find("id=\"2\"") != std::string_view::npos, "self-closing element with whitespace is returned");
    check(reader.nextElement("row").empty(), "clean EOF returns empty view");

    const std::string truncated = "<root><row><c>42</c>";
    offset = 0;
    xlpp::internal::XmlPullReader bad([&](unsigned char* out, std::size_t capacity) {
        const auto count = std::min(capacity, truncated.size() - offset);
        if (!count) return std::size_t{0};
        std::copy_n(reinterpret_cast<const unsigned char*>(truncated.data() + offset), count, out);
        offset += count;
        return count;
    }, 1024);
    checkThrows([&] { (void)bad.nextElement("row"); }, "truncated matching XML element is rejected");

    const std::string oversized = "<row>" + std::string(256, 'x') + "</row>";
    offset = 0;
    xlpp::internal::XmlPullReader limited([&](unsigned char* out, std::size_t capacity) {
        const auto count = std::min<std::size_t>({capacity, 32u, oversized.size() - offset});
        if (!count) return std::size_t{0};
        std::copy_n(reinterpret_cast<const unsigned char*>(oversized.data() + offset), count, out);
        offset += count;
        return count;
    }, 128);
    checkThrows([&] { (void)limited.nextElement("row"); }, "streaming XML element buffer limit is enforced");
}

std::filesystem::path writeStreamingPackage(std::string target,
                                            std::string extraRelationships = {},
                                            std::string sheetRid = "rId1") {
    xlpp::internal::ZipArchive zip;
    zip.add("xl/workbook.xml",
            "<?xml version=\"1.0\"?><workbook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Sheet1\" r:id=\"" +
                sheetRid + "\"/></sheets></workbook>", false);
    zip.add("xl/_rels/workbook.xml.rels",
            "<?xml version=\"1.0\"?><Relationships><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"" +
                target + "\"/>" + extraRelationships + "</Relationships>", false);
    zip.add("xl/worksheets/sheet1.xml",
            "<?xml version=\"1.0\"?><worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>7</v></c></row></sheetData></worksheet>", false);
    zip.add("evil.xml", "<worksheet/>", false);
    auto path = tempPath("xlpp_p1q_streaming");
    zip.save(path);
    return path;
}

void testRelationshipHardening() {
    {
        const auto path = writeStreamingPackage("worksheets/./sheet1.xml");
        try {
            xlpp::StreamingWorkbookReader reader(path);
            const auto names = reader.worksheetNames();
            check(names.size() == 1 && names[0] == "Sheet1", "normalized internal worksheet target loads");
            auto sheet = reader.worksheet("Sheet1");
            auto it = sheet.begin();
            check(it != sheet.end() && it.rowNumber() == 1, "normalized target streams worksheet rows");
        } catch (const std::exception& ex) {
            std::cerr << "Unexpected normalized-target failure: " << ex.what() << '\n';
            ++failures;
        }
        std::filesystem::remove(path);
    }
    {
        const auto path = writeStreamingPackage("../../evil.xml");
        checkThrows([&] { xlpp::StreamingWorkbookReader reader(path); },
                    "relationship target cannot escape OPC package root");
        std::filesystem::remove(path);
    }
    {
        const auto path = writeStreamingPackage(
            "worksheets/sheet1.xml",
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>");
        checkThrows([&] { xlpp::StreamingWorkbookReader reader(path); },
                    "duplicate workbook relationship Id is rejected");
        std::filesystem::remove(path);
    }
    {
        const auto path = writeStreamingPackage("worksheets/sheet1.xml", {}, "rIdMissing");
        checkThrows([&] { xlpp::StreamingWorkbookReader reader(path); },
                    "dangling worksheet relationship is rejected when rels part exists");
        std::filesystem::remove(path);
    }
}

void testStreamingResourceLimits() {
    const auto path = writeStreamingPackage("worksheets/sheet1.xml");
    {
        xlpp::StreamingReadOptions options;
        options.maxEntries = 3;
        checkThrows([&] { xlpp::StreamingWorkbookReader reader(path, options); },
                    "streaming reader enforces ZIP entry-count limit");
    }
    {
        xlpp::StreamingReadOptions options;
        options.maxEntryBytes = 32;
        checkThrows([&] { xlpp::StreamingWorkbookReader reader(path, options); },
                    "streaming reader enforces ZIP per-entry size limit");
    }
    {
        xlpp::StreamingReadOptions options;
        options.maxTotalBytes = 64;
        checkThrows([&] { xlpp::StreamingWorkbookReader reader(path, options); },
                    "streaming reader enforces ZIP total-uncompressed limit");
    }
    {
        xlpp::StreamingReadOptions options;
        options.maxXmlElementBytes = 64;
        xlpp::StreamingWorkbookReader reader(path, options);
        auto sheet = reader.worksheet("Sheet1");
        checkThrows([&] { (void)sheet.begin(); },
                    "streaming reader enforces per-element XML buffer limit");
    }
    std::filesystem::remove(path);
}

std::uint32_t readLe32(const std::vector<unsigned char>& bytes, std::size_t offset) {
    return static_cast<std::uint32_t>(bytes[offset]) |
           (static_cast<std::uint32_t>(bytes[offset + 1]) << 8u) |
           (static_cast<std::uint32_t>(bytes[offset + 2]) << 16u) |
           (static_cast<std::uint32_t>(bytes[offset + 3]) << 24u);
}

void writeLe64(std::vector<unsigned char>& bytes, std::size_t offset, std::uint64_t value) {
    for (unsigned i = 0; i < 8; ++i) bytes[offset + i] = static_cast<unsigned char>((value >> (8u * i)) & 0xffu);
}

void testCompoundFileTruncationHardening() {
    auto cfb = xlpp::internal::buildRootCompoundFile({{"Data", std::vector<unsigned char>(16, 0x5a)}});
    check(xlpp::internal::readCompoundFileStream(cfb, "Data").size() == 16,
          "valid compact CFB stream round-trips before corruption");

    const auto directorySector = readLe32(cfb, 48);
    const std::size_t directoryOffset = 512u + static_cast<std::size_t>(directorySector) * 512u;
    bool patched = false;
    for (std::size_t entry = 0; entry < 4; ++entry) {
        const auto off = directoryOffset + entry * 128u;
        if (off + 128u > cfb.size()) break;
        if (cfb[off] == 'D' && cfb[off + 1] == 0 && cfb[off + 2] == 'a' && cfb[off + 3] == 0) {
            writeLe64(cfb, off + 120u, 200u); // larger than the one mini-sector chain
            patched = true;
            break;
        }
    }
    check(patched, "test located CFB Data directory entry");
    checkThrows([&] { (void)xlpp::internal::readCompoundFileStream(cfb, "Data"); },
                "CFB stream whose declared size exceeds its FAT chain is rejected");
}

} // namespace

int main() {
    testPullReaderHardening();
    testRelationshipHardening();
    testStreamingResourceLimits();
    testCompoundFileTruncationHardening();
    if (failures) {
        std::cerr << failures << " P1Q hardening check(s) failed\n";
        return 1;
    }
    std::cout << "P1Q core hardening checks PASS\n";
    return 0;
}
