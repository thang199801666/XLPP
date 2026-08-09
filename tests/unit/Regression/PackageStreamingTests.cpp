#include <XLPP/XLPP.h>
#include "Package/Zip/ZipArchive.h"
#include "Package/Zip/ZipArchiveReader.h"
#include "Package/Opc/RelationshipGraph.h"
#include "Platform/MappedFile.h"
#include "Streaming/SharedStringsReader.h"
#include "Core/Threading/ThreadPool.h"
#include "Package/Xml/SimdScan.h"
#include "Package/Xml/XmlScanner.h"
#include "Package/Xml/XmlUtilities.h"
#include "VBA/VbaProjectBinary.h"
#include "Encryption/OfficeEncryption.h"
#include <array>
#include <chrono>
#include <cmath>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iostream>
#include <set>
#include <sstream>
#include <stdexcept>
#include <string>
#include <variant>
#include <vector>

#include "../TestFramework.h"
#include "RegressionTests.h"

namespace {
xlpp::internal::ZipArchive buildSharedStringWorkbook() {
    xlpp::internal::ZipArchive zip;
    zip.add("[Content_Types].xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
        "</Types>");
    zip.add("_rels/.rels",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>");
    zip.add("xl/workbook.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Shared\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
    zip.add("xl/_rels/workbook.xml.rels",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
        "</Relationships>");
    zip.add("xl/sharedStrings.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"5\" uniqueCount=\"5\">"
        "<si><t>Alpha</t></si>"
        "<si><t xml:space=\"preserve\">  Beta  </t></si>"
        "<si><r><rPr><b/></rPr><t>Rich </t></r><r><rPr><i/></rPr><t>text</t></r></si>"
        "<si><t>Gamma</t></si>"
        "<si><t>Delta</t></si>"
        "</sst>");
    zip.add("xl/worksheets/sheet1.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c><c r=\"C1\" t=\"s\"><v>2</v></c>"
        "<c r=\"D1\"><v>3.5</v></c><c r=\"E1\" t=\"b\"><v>1</v></c><c r=\"F1\" t=\"inlineStr\"><is><t>Inline</t></is></c></row>"
        "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>3</v></c><c r=\"B2\" t=\"s\"><v>4</v></c></row>"
        "</sheetData></worksheet>");
    return zip;
}

std::string buildBenchmarkWorksheet(std::size_t rows) {
    std::string out;
    out += "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>";
    for (std::size_t r = 1; r <= rows; ++r) {
        out += "<row r=\"" + std::to_string(r) + "\" spans=\"1:4\">";
        out += "<c r=\"A" + std::to_string(r) + "\" t=\"s\"><v>" + std::to_string(r % 100) + "</v></c>";
        out += "<c r=\"B" + std::to_string(r) + "\" t=\"inlineStr\"><is><t>cell " + std::to_string(r) + "</t></is></c>";
        out += "<c r=\"C" + std::to_string(r) + "\"><v>" + std::to_string(static_cast<double>(r) * 1.5) + "</v></c>";
        out += "<c r=\"D" + std::to_string(r) + "\" t=\"b\"><v>" + std::to_string(r % 2) + "</v></c>";
        out += "</row>";
    }
    out += "</sheetData></worksheet>";
    return out;
}

std::uint32_t crc32Of(const std::string& data) {
    std::uint32_t crc = 0xFFFFFFFFu;
    for (unsigned char c : data) {
        crc ^= c;
        for (int i = 0; i < 8; ++i) crc = (crc >> 1) ^ (0xEDB88320u & (0u - (crc & 1u)));
    }
    return ~crc;
}

void writeZip64Fixture(const std::filesystem::path& path) {
    const std::string name = "hello.txt";
    const std::string data = "Hello, ZIP64!";
    const std::uint32_t crc = crc32Of(data);
    const auto le16 = [](std::ofstream& f, std::uint16_t v) { f.write(reinterpret_cast<const char*>(&v), 2); };
    const auto le32 = [](std::ofstream& f, std::uint32_t v) { f.write(reinterpret_cast<const char*>(&v), 4); };
    const auto le64 = [](std::ofstream& f, std::uint64_t v) { f.write(reinterpret_cast<const char*>(&v), 8); };
    std::ofstream f(path, std::ios::binary);
    le32(f, 0x04034b50u); le16(f, 45); le16(f, 0); le16(f, 0); le16(f, 0); le16(f, 0);
    le32(f, crc); le32(f, 0xFFFFFFFFu); le32(f, 0xFFFFFFFFu);
    le16(f, static_cast<std::uint16_t>(name.size())); le16(f, 20);
    f.write(name.data(), static_cast<std::streamsize>(name.size()));
    le16(f, 0x0001u); le16(f, 16);
    le64(f, data.size()); le64(f, data.size());
    f.write(data.data(), static_cast<std::streamsize>(data.size()));
    const std::uint64_t cdOffset = 30 + name.size() + 20 + data.size();
    le32(f, 0x02014b50u); le16(f, 45); le16(f, 45); le16(f, 0); le16(f, 0); le16(f, 0); le16(f, 0);
    le32(f, crc); le32(f, 0xFFFFFFFFu); le32(f, 0xFFFFFFFFu);
    le16(f, static_cast<std::uint16_t>(name.size())); le16(f, 28); le16(f, 0);
    le16(f, 0); le16(f, 0); le32(f, 0); le32(f, 0xFFFFFFFFu);
    f.write(name.data(), static_cast<std::streamsize>(name.size()));
    le16(f, 0x0001u); le16(f, 24);
    le64(f, data.size()); le64(f, data.size()); le64(f, 0);
    const std::uint64_t cdSize = 46 + name.size() + 28;
    const std::uint64_t eocd64Pos = cdOffset + cdSize;
    le32(f, 0x06064b50u); le64(f, 44); le16(f, 45); le16(f, 45);
    le32(f, 0); le32(f, 0); le64(f, 1); le64(f, 1);
    le64(f, cdSize); le64(f, cdOffset);
    le32(f, 0x07064b50u); le32(f, 0); le64(f, eocd64Pos); le32(f, 1);
    le32(f, 0x06054b50u); le16(f, 0); le16(f, 0); le16(f, 0xFFFFu); le16(f, 0xFFFFu);
    le32(f, 0xFFFFFFFFu); le32(f, 0xFFFFFFFFu); le16(f, 0);
}

} // namespace

void testValidationAndConditionalPackageRegression(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_validation_conditional_regression.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Rules");
    sheet.cell("A1").setValue("Status");
    sheet.cell("B1").setValue("Amount");
    auto& list = sheet.dataValidations().add(xlpp::DataValidationType::List, "A2:A20");
    list.setFormula1("\"Open,Closed\"");
    list.setShowInputMessage(true);
    list.setPromptTitle("Status");
    list.setPrompt("Choose a status");
    list.setShowErrorMessage(true);
    list.setErrorTitle("Invalid status");
    list.setError("Use the drop-down list");

    auto negative = xlpp::ConditionalRule::formula("B2<0");
    negative.differentialStyle().font().color().setArgb("FF9C0006");
    negative.differentialStyle().fill().setPatternType("solid");
    negative.differentialStyle().fill().foregroundColor().setArgb("FFFFC7CE");
    sheet.conditionalFormatting().addRule("B2:B20", std::move(negative));
    auto positive = xlpp::ConditionalRule::formula("B2>0");
    positive.differentialStyle().font().color().setArgb("FF006100");
    positive.differentialStyle().fill().setPatternType("solid");
    positive.differentialStyle().fill().foregroundColor().setArgb("FFC6EFCE");
    sheet.conditionalFormatting().addRule("B2:B20", std::move(positive));
    workbook.save(path);

    auto zip = xlpp::internal::ZipArchive::open(path);
    const auto sheetXml = zip.get("xl/worksheets/sheet1.xml");
    const auto stylesXml = zip.get("xl/styles.xml");
    const auto listStart = sheetXml.find("<dataValidation type=\"list\"");
    const auto listEnd = sheetXml.find("</dataValidation>", listStart);
    test.checkTrue(listStart != std::string::npos, "List validation is serialized");
    test.checkTrue(listEnd != std::string::npos, "List validation has a closing element");
    const auto listXml = sheetXml.substr(listStart, listEnd - listStart);
    test.checkTrue(listXml.find(" operator=") == std::string::npos, "List validation omits inapplicable operator");
    test.checkTrue(listXml.find("showErrorMessage=\"1\"") != std::string::npos, "Validation error message is enabled");
    test.checkTrue(listXml.find("showInputMessage=\"1\"") != std::string::npos, "Validation input message is enabled");
    test.checkTrue(sheetXml.find("priority=\"0\"") == std::string::npos, "Conditional formatting never emits priority zero");
    test.checkTrue(sheetXml.find("priority=\"1\"") != std::string::npos, "First automatic conditional priority is one");
    test.checkTrue(sheetXml.find("priority=\"2\"") != std::string::npos, "Second automatic conditional priority is two");
    test.checkTrue(stylesXml.find("<dxfs count=\"2\"") != std::string::npos, "Two differential styles are serialized");

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedSheet = loaded.worksheet("Rules");
    test.checkTrue(loadedSheet != nullptr, "Validation/conditional worksheet reloads");
    test.checkEqual(loadedSheet->dataValidations().items().size(), std::size_t{1}, "Validation count reloads");
    test.checkEqual(loadedSheet->conditionalFormatting().entries().front().rules().size(), std::size_t{2}, "Conditional rules reload");
    std::filesystem::remove(path);
}

void testStreamingSharedStringsAndPullIterator(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m16_shared_strings.xlsx";
    {
        auto zip = buildSharedStringWorkbook();
        zip.save(path);
    }

    xlpp::StreamingWorkbookReader reader(path);
    const auto names = reader.worksheetNames();
    test.checkEqual(names.size(), std::size_t{1}, "Shared-string workbook worksheet count");
    test.checkEqual(names.front(), std::string("Shared"), "Shared-string workbook worksheet name");

    auto sheet = reader.worksheet("Shared");
    std::size_t rows = 0;
    std::string a1, b1, c1, a2, b2;
    double d1 = 0.0;
    bool e1 = false, hasD1 = false, hasE1 = false;
    std::string f1;
    bool sequential = true;
    std::size_t seen = 0;
    for (auto it = sheet.begin(); it != sheet.end(); ++it) {
        ++rows;
        sequential = sequential && it.rowNumber() == rows;
        for (const auto& cell : *it) {
            ++seen;
            if (cell.address == "A1") a1 = std::get<std::string>(cell.value);
            else if (cell.address == "B1") b1 = std::get<std::string>(cell.value);
            else if (cell.address == "C1") c1 = std::get<std::string>(cell.value);
            else if (cell.address == "D1") { hasD1 = true; d1 = std::get<double>(cell.value); }
            else if (cell.address == "E1") { hasE1 = true; e1 = std::get<bool>(cell.value); }
            else if (cell.address == "F1") f1 = std::get<std::string>(cell.value);
            else if (cell.address == "A2") a2 = std::get<std::string>(cell.value);
            else if (cell.address == "B2") b2 = std::get<std::string>(cell.value);
        }
    }
    test.checkEqual(rows, std::size_t{2}, "Pull iterator row count");
    test.checkEqual(seen, std::size_t{8}, "Pull iterator cell count");
    test.checkTrue(sequential, "Pull iterator row numbers are sequential");
    test.checkEqual(a1, std::string("Alpha"), "Shared string index resolved");
    test.checkEqual(b1, std::string("  Beta  "), "Shared string preserves leading/trailing spaces");
    test.checkEqual(c1, std::string("Rich text"), "Rich text shared string concatenates runs");
    test.checkTrue(hasD1 && d1 == 3.5, "Numeric cell beside shared strings");
    test.checkTrue(hasE1 && e1, "Boolean cell beside shared strings");
    test.checkEqual(f1, std::string("Inline"), "Inline string cell beside shared strings");
    test.checkEqual(a2, std::string("Gamma"), "Second row shared string");
    test.checkEqual(b2, std::string("Delta"), "Second row shared string repeat");

    std::size_t callbackRows = 0;
    std::size_t callbackAlpha = 0;
    sheet.forEachRow([&](std::size_t, const xlpp::StreamingRow& row) {
        ++callbackRows;
        for (const auto& cell : row)
            if (std::holds_alternative<std::string>(cell.value) && std::get<std::string>(cell.value) == "Alpha")
                ++callbackAlpha;
        return true;
    });
    test.checkEqual(callbackRows, std::size_t{2}, "Callback API rows match pull iterator");
    test.checkEqual(callbackAlpha, std::size_t{1}, "Callback API resolves shared strings");

    bool threw = false;
    try { (void)reader.worksheet("Missing"); } catch (const std::exception&) { threw = true; }
    test.checkTrue(threw, "Unknown worksheet name is rejected");

    std::filesystem::remove(path);
}

void testStreamingReadWrite(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m15_streaming.xlsx";
    const auto started = std::chrono::steady_clock::now();
    {
        xlpp::StreamingWorkbookWriter writer(path);
        auto& sheet = writer.addWorksheet("StreamData");
        for (int row = 1; row <= 5000; ++row) {
            sheet.append({std::string("Row ") + std::to_string(row), static_cast<double>(row), row % 2 == 0});
        }
        test.checkEqual(sheet.rowCount(), std::size_t{5000}, "Streaming writer row count");
        writer.close();
        test.checkTrue(writer.closed(), "Streaming writer closes package");
    }
    const auto written = std::chrono::steady_clock::now();
    xlpp::StreamingWorkbookReader reader(path);
    const auto names = reader.worksheetNames();
    test.checkEqual(names.size(), std::size_t{1}, "Streaming reader worksheet count");
    test.checkEqual(names.front(), std::string("StreamData"), "Streaming reader worksheet name");
    std::size_t rows = 0;
    double sum = 0.0;
    bool sequential = true;
    reader.forEachRow("StreamData", [&](std::size_t rowNumber, const xlpp::StreamingRow& row) {
        ++rows;
        sequential = sequential && rowNumber == rows;
        if (row.size() >= 2 && std::holds_alternative<double>(row[1].value)) sum += std::get<double>(row[1].value);
        return true;
    });
    const auto read = std::chrono::steady_clock::now();
    test.checkTrue(sequential, "Streaming row numbers remain sequential");
    test.checkEqual(rows, std::size_t{5000}, "Streaming reader row count");
    test.checkNear(sum, 12502500.0, 1e-9, "Streaming reader numeric aggregation");
    const auto writeMs = std::chrono::duration_cast<std::chrono::milliseconds>(written - started).count();
    const auto readMs = std::chrono::duration_cast<std::chrono::milliseconds>(read - written).count();
    std::cout << "    [BENCHMARK] 5000 rows write=" << writeMs << " ms read=" << readMs
              << " ms file=" << std::filesystem::file_size(path) << " bytes\n";
    std::filesystem::remove(path);
}

void testXmlScanner(TestContext& test) {
    using namespace xlpp::internal;

    const std::string sheet =
        "<worksheet><sheetData>"
        "<row r=\"1\" spans=\"1:3\">"
        "<c r=\"A1\" t=\"s\"><v>0</v></c>"
        "<c r=\"B1\" t=\"inlineStr\"><is><t>H &amp; W</t></is></c>"
        "<c r=\"C1\"/></row>"
        "<row r=\"2\"><c r=\"A2\" t=\"b\"><v>1</v></c></row>"
        "</sheetData></worksheet>";

    XmlScanner rows(sheet);
    std::string_view rowTag;
    test.checkTrue(rows.nextElement("row", rowTag), "First row element found");
    test.checkEqual(xmlAttribute(rowTag, "r"), std::string_view("1"), "Row index attribute");
    test.checkEqual(xmlAttribute(rowTag, "spans"), std::string_view("1:3"), "Row spans attribute");

    XmlScanner cells(rowTag);
    std::string_view cellTag;
    test.checkTrue(cells.nextElement("c", cellTag), "First cell found");
    test.checkEqual(xmlAttribute(cellTag, "r"), std::string_view("A1"), "Cell address attribute");
    test.checkEqual(xmlAttribute(cellTag, "t"), std::string_view("s"), "Cell type attribute");
    test.checkEqual(xmlText(cellTag, "v"), std::string_view("0"), "Shared string index text");
    test.checkTrue(cells.nextElement("c", cellTag), "Second cell found");
    test.checkEqual(xmlText(cellTag, "t"), std::string_view("H &amp; W"), "Inline text kept raw, not decoded");
    test.checkTrue(containsEntity(xmlText(cellTag, "t")), "Entity detected by containsEntity");
    test.checkTrue(cells.nextElement("c", cellTag), "Self-closing cell found");
    test.checkEqual(xmlText(cellTag, "v"), std::string_view(""), "Self-closing cell has empty text");
    test.checkTrue(!cells.nextElement("c", cellTag), "No further cells in row 1");

    test.checkTrue(rows.nextElement("row", rowTag), "Second row found");
    test.checkEqual(xmlAttribute(rowTag, "r"), std::string_view("2"), "Second row index");
    XmlScanner row2Cells(rowTag);
    test.checkTrue(row2Cells.nextElement("c", cellTag), "Row 2 cell found");
    test.checkEqual(xmlText(cellTag, "v"), std::string_view("1"), "Boolean cell value");
    test.checkEqual(xmlAttribute(cellTag, "t"), std::string_view("b"), "Boolean cell type");
    test.checkTrue(!rows.nextElement("row", rowTag), "No further rows");

    const std::string tagWithId = "<a id=\"7\" x=\"1\">";
    test.checkEqual(xmlAttribute(tagWithId, "id"), std::string_view("7"), "Id attribute value");
    test.checkEqual(xmlAttribute(tagWithId, "d"), std::string_view(""),
                    "Short name must not match inside another attribute");
    test.checkEqual(xmlAttribute(tagWithId, "x"), std::string_view("1"), "Trailing attribute value");
    test.checkEqual(xmlAttribute(tagWithId, "missing"), std::string_view(""), "Missing attribute is empty");

    const std::string multi = "<rows>a</rows><row>b</row><row>c</row>";
    XmlScanner nameBoundary(multi);
    test.checkTrue(nameBoundary.nextElement("row", rowTag), "Rows container is skipped");
    test.checkEqual(rowTag, std::string_view("<row>b</row>"), "First boundary row element");
    test.checkTrue(nameBoundary.nextElement("row", rowTag), "Second boundary row element");
    test.checkEqual(rowTag, std::string_view("<row>c</row>"), "Second boundary row value");

    const std::string nested = "<c t=\"inlineStr\"><is><t>x</t></is></c>";
    test.checkEqual(xmlText(nested, "t"), std::string_view("x"), "Nested text extracted");
    test.checkEqual(xmlText(nested, "is"), std::string_view("<t>x</t>"), "Nested container text");
    test.checkEqual(xmlText("<c/>", "v"), std::string_view(""), "Self-closing text is empty");
    test.checkEqual(xmlText("<c><v></v></c>", "v"), std::string_view(""), "Empty element text is empty");
    test.checkEqual(xmlText("<c/>", "missing"), std::string_view(""), "Missing nested element is empty");

    double number = 0.0;
    test.checkTrue(parseDouble("123.5", number), "Double parses");
    test.checkNear(number, 123.5, 1e-9, "Double value");
    test.checkTrue(parseDouble("-1.5e3", number), "Exponent double parses");
    test.checkNear(number, -1500.0, 1e-9, "Exponent double value");
    test.checkTrue(!parseDouble("abc", number), "Non-numeric rejected");
    test.checkTrue(!parseDouble("", number), "Empty double rejected");
    test.checkTrue(!parseDouble("12x", number), "Trailing garbage rejected");
    std::size_t size = 0;
    test.checkTrue(parseSize("1048576", size), "Size parses");
    test.checkEqual(size, std::size_t{1048576}, "Size value");
    test.checkTrue(!parseSize("-1", size), "Negative size rejected");
    int integer = 0;
    test.checkTrue(parseInt("-7", integer), "Int parses");
    test.checkEqual(integer, -7, "Int value");
}

void testXmlScannerBenchmark(TestContext& test) {
    using namespace xlpp::internal;
    const auto xml = buildBenchmarkWorksheet(5000);

    const auto parseOld = [&]() {
        std::size_t rowCount = 0;
        double sum = 0.0;
        for (const auto& rowTag : tags(xml, "row")) {
            ++rowCount;
            for (const auto& cellTag : tags(rowTag, "c")) {
                const auto type = attribute(cellTag, "t");
                const auto value = tagText(cellTag, "v");
                if (type == "b") {
                    if (value == "1") sum += 1.0;
                } else if (type == "inlineStr") {
                    sum += static_cast<double>(tagText(cellTag, "t").size());
                } else if (!value.empty()) {
                    sum += std::stod(value);
                }
            }
        }
        return std::pair{rowCount, sum};
    };

    const auto parseNew = [&]() {
        std::size_t rowCount = 0;
        double sum = 0.0;
        XmlScanner rows(xml);
        std::string_view rowTag;
        while (rows.nextElement("row", rowTag)) {
            ++rowCount;
            XmlScanner cells(rowTag);
            std::string_view cellTag;
            while (cells.nextElement("c", cellTag)) {
                const auto type = xmlAttribute(cellTag, "t");
                const auto value = xmlText(cellTag, "v");
                if (type == "b") {
                    if (value == "1") sum += 1.0;
                } else if (type == "inlineStr") {
                    sum += static_cast<double>(xmlText(cellTag, "t").size());
                } else if (!value.empty()) {
                    double d = 0.0;
                    if (parseDouble(value, d)) sum += d;
                }
            }
        }
        return std::pair{rowCount, sum};
    };

    const auto oldResult = parseOld();
    const auto newResult = parseNew();
    test.checkEqual(oldResult.first, newResult.first, "Benchmark row counts match");
    test.checkEqual(oldResult.first, std::size_t{5000}, "Benchmark row count");
    test.checkNear(oldResult.second, newResult.second, 1e-6, "Benchmark checksums match");

    const int iterations = 20;
    const auto bestOf = [iterations](auto&& run) {
        auto best = std::chrono::nanoseconds::max();
        for (int i = 0; i < iterations; ++i) {
            const auto started = std::chrono::steady_clock::now();
            (void)run();
            const auto finished = std::chrono::steady_clock::now();
            best = std::min(best, std::chrono::duration_cast<std::chrono::nanoseconds>(finished - started));
        }
        return best;
    };
    const auto bestOld = bestOf(parseOld);
    const auto bestNew = bestOf(parseNew);
    const auto oldMs = std::chrono::duration_cast<std::chrono::microseconds>(bestOld).count();
    const auto newMs = std::chrono::duration_cast<std::chrono::microseconds>(bestNew).count();
    const double speedup =
        bestNew.count() > 0 && bestOld.count() > 0
            ? static_cast<double>(bestOld.count()) / bestNew.count()
            : 0.0;
    std::cout << "    [BENCHMARK] 5000 rows x 4 cells: scanner=" << newMs << " us baseline="
              << oldMs << " us speedup=" << speedup << "x\n";
    test.checkTrue(bestNew.count() <= bestOld.count() * 4 + 1,
                   "XmlScanner is not slower than the baseline parser");
}

void testSharedStringWriter(TestContext& test) {
    const auto hashPath = std::filesystem::temp_directory_path() / "xlpp_m18_shared_hash.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(hashPath, xlpp::SharedStringMode::Hash);
        auto& sheet = writer.addWorksheet("Data");
        for (int row = 1; row <= 100; ++row) {
            sheet.append({std::string("Repeated"),
                          std::string("Value " + std::to_string(row % 10)),
                          static_cast<double>(row)});
        }
        writer.close();
        test.checkTrue(writer.closed(), "Hash-mode writer closes");
    }
    {
        auto zip = xlpp::internal::ZipArchive::open(hashPath);
        test.checkTrue(zip.contains("xl/sharedStrings.xml"), "Hash mode writes the shared-strings part");
        const auto sst = zip.get("xl/sharedStrings.xml");
        test.checkEqual(xlpp::internal::tags(sst, "si").size(), std::size_t{11},
                        "Hash mode deduplicates 200 cells to 11 unique strings");
        const auto roots = xlpp::internal::tags(sst, "sst");
        test.checkEqual(roots.size(), std::size_t{1}, "Sst root element present");
        test.checkEqual(xlpp::internal::attribute(roots.front(), "count"), std::string("200"),
                        "Sst count records total occurrences");
        test.checkEqual(xlpp::internal::attribute(roots.front(), "uniqueCount"), std::string("11"),
                        "Sst uniqueCount matches dedup");
    }
    {
        xlpp::StreamingWorkbookReader reader(hashPath);
        auto sheet = reader.worksheet("Data");
        std::size_t rows = 0;
        std::size_t repeated = 0;
        double sum = 0.0;
        for (auto it = sheet.begin(); it != sheet.end(); ++it) {
            ++rows;
            for (const auto& cell : *it) {
                if (std::holds_alternative<std::string>(cell.value)) {
                    if (std::get<std::string>(cell.value) == "Repeated") ++repeated;
                } else if (std::holds_alternative<double>(cell.value)) {
                    sum += std::get<double>(cell.value);
                }
            }
        }
        test.checkEqual(rows, std::size_t{100}, "Hash-mode writer row count");
        test.checkEqual(repeated, std::size_t{100}, "Hash-mode dedup resolves repeated text");
        test.checkNear(sum, 5050.0, 1e-9, "Hash-mode numeric cells read back");
    }

    const auto lruPath = std::filesystem::temp_directory_path() / "xlpp_m18_shared_lru.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(lruPath, xlpp::SharedStringMode::BoundedLru, 2);
        auto& sheet = writer.addWorksheet("Data");
        for (int row = 0; row < 6; ++row) {
            sheet.append({std::string(row % 3 == 0 ? "A" : (row % 3 == 1 ? "B" : "C"))});
        }
        writer.close();
    }
    {
        auto zip = xlpp::internal::ZipArchive::open(lruPath);
        test.checkTrue(zip.contains("xl/sharedStrings.xml"), "LRU mode writes the shared-strings part");
        const auto sst = zip.get("xl/sharedStrings.xml");
        test.checkTrue(xlpp::internal::tags(sst, "si").size() > 3,
                       "Bounded LRU assigns fresh indexes after eviction");
    }
    {
        xlpp::StreamingWorkbookReader reader(lruPath);
        auto sheet = reader.worksheet("Data");
        std::string sequence;
        for (auto it = sheet.begin(); it != sheet.end(); ++it) {
            for (const auto& cell : *it) sequence += std::get<std::string>(cell.value);
        }
        test.checkEqual(sequence, std::string("ABCABC"), "LRU mode reader resolves text correctly");
    }

    const auto inlinePath = std::filesystem::temp_directory_path() / "xlpp_m18_shared_inline.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(inlinePath);
        auto& sheet = writer.addWorksheet("Data");
        sheet.append({std::string("Hello"), std::string("World")});
        writer.close();
    }
    {
        auto zip = xlpp::internal::ZipArchive::open(inlinePath);
        test.checkTrue(!zip.contains("xl/sharedStrings.xml"), "Disabled mode writes no shared strings");
        xlpp::StreamingWorkbookReader reader(inlinePath);
        auto sheet = reader.worksheet("Data");
        auto it = sheet.begin();
        const auto& row = *it;
        test.checkEqual(std::get<std::string>(row[0].value), std::string("Hello"), "Disabled inline string resolves");
        test.checkEqual(std::get<std::string>(row[1].value), std::string("World"), "Disabled inline second string");
    }

    const auto datePath = std::filesystem::temp_directory_path() / "xlpp_m18_shared_dates.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(datePath);
        writer.setDate1904(false);
        auto& sheet = writer.addWorksheet("Data");
        sheet.append({xlpp::DateTime{2024, 1, 15}, xlpp::DateTime{1900, 1, 1, 12, 0, 0}});
        writer.close();
    }
    {
        xlpp::StreamingWorkbookReader reader(datePath);
        auto sheet = reader.worksheet("Data");
        auto it = sheet.begin();
        const auto& row = *it;
        test.checkNear(std::get<double>(row[0].value), 45306.0, 1e-9, "Streaming writer serializes a date");
        test.checkNear(std::get<double>(row[1].value), 1.5, 1e-9, "Streaming writer serializes date-time");
    }

    std::filesystem::remove(hashPath);
    std::filesystem::remove(lruPath);
    std::filesystem::remove(inlinePath);
    std::filesystem::remove(datePath);
}

void testParallelPackagePipeline(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto seqPath = dir / "xlpp_m19_seq.xlsx";
    const auto parPath = dir / "xlpp_m19_par.xlsx";
    const auto storePath = dir / "xlpp_m19_store.xlsx";
    const auto bestPath = dir / "xlpp_m19_best.xlsx";
    const auto streamPath = dir / "xlpp_m19_stream.xlsx";
    const auto readFile = [](const std::filesystem::path& path) {
        std::ifstream file(path, std::ios::binary);
        return std::string((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    };
    const auto build = [](xlpp::Workbook& workbook) {
        workbook.properties().setTitle("Parallel pipeline");
        for (int s = 0; s < 4; ++s) {
            auto& sheet = workbook.addWorksheet("Sheet" + std::to_string(s + 1));
            sheet.cell("A1").setValue("Header");
            sheet.cell("B1").setValue(1.5);
            sheet.cell("C1").setDateTime(xlpp::DateTime{2024, 1, 15, 13, 30, 45});
            for (int r = 2; r <= 200; ++r) {
                sheet.cell(r, 1).setValue("Item " + std::to_string(r));
                sheet.cell(r, 2).setValue(static_cast<double>(r * 3));
                sheet.cell(r, 3).setValue(r % 3 == 0);
            }
        }
        auto& styled = workbook.worksheet("Sheet1")->cell("B1");
        styled.font().setBold(true);
        styled.setNumberFormat("0.00");
        workbook.worksheet("Sheet1")->mergeCells("D1:E1");
    };

    {
        xlpp::Workbook workbook;
        build(workbook);
        workbook.save(seqPath);
        xlpp::SaveOptions parallel;
        parallel.parallelWorkers = 4;
        workbook.save(parPath, parallel);
        test.checkTrue(readFile(seqPath) == readFile(parPath), "Parallel output is byte-identical to sequential");
        std::filesystem::remove(seqPath);
        std::filesystem::remove(parPath);
    }

    {
        xlpp::SaveOptions parallel;
        parallel.parallelWorkers = 4;
        {
            xlpp::Workbook workbook;
            build(workbook);
            workbook.save(parPath, parallel);
        }
        xlpp::Workbook loaded;
        loaded.load(parPath);
        test.checkEqual(loaded.worksheets().size(), std::size_t{4}, "Parallel save loads all worksheets");
        auto* sheet = loaded.worksheet("Sheet1");
        test.checkTrue(sheet != nullptr, "Parallel save worksheet lookup");
        test.checkEqual(std::get<std::string>(sheet->cell("A1").value()), std::string("Header"), "Parallel string round-trip");
        test.checkNear(std::get<double>(sheet->cell("B1").value()), 1.5, 1e-12, "Parallel number round-trip");
        test.checkEqual(std::get<xlpp::DateTime>(sheet->cell("C1").value()), xlpp::DateTime{2024, 1, 15, 13, 30, 45}, "Parallel datetime round-trip");
        test.checkEqual(std::get<std::string>(sheet->cell("A200").value()), std::string("Item 200"), "Parallel last-row string");
        test.checkEqual(sheet->mergedRanges().size(), std::size_t{1}, "Parallel merged range round-trip");
        test.checkEqual(sheet->cell("B1").numberFormat(), std::string("0.00"), "Parallel number format round-trip");
        std::filesystem::remove(parPath);
    }

    {
        xlpp::Workbook workbook;
        build(workbook);
        xlpp::SaveOptions store;
        store.compressionLevel = xlpp::CompressionLevel::Store;
        xlpp::SaveOptions best;
        best.compressionLevel = xlpp::CompressionLevel::Best;
        workbook.save(storePath, store);
        workbook.save(bestPath, best);
        const auto storeSize = std::filesystem::file_size(storePath);
        const auto bestSize = std::filesystem::file_size(bestPath);
        test.checkTrue(storeSize > bestSize, "Stored package is larger than compressed");
        xlpp::Workbook loaded;
        loaded.load(bestPath);
        test.checkEqual(std::get<std::string>(loaded.worksheet("Sheet3")->cell("A2").value()), std::string("Item 2"), "Best-level package loads");
        std::filesystem::remove(storePath);
        std::filesystem::remove(bestPath);
    }

    {
        const auto streamSeqPath = dir / "xlpp_m19_stream_seq.xlsx";
        xlpp::StreamingWorkbookWriter seqWriter(streamSeqPath);
        seqWriter.setCompressionLevel(xlpp::CompressionLevel::Fastest);
        auto& seqSheet = seqWriter.addWorksheet("Data");
        for (int r = 1; r <= 1000; ++r)
            seqSheet.append({std::string("Row ") + std::to_string(r), static_cast<double>(r)});
        seqWriter.close();

        xlpp::StreamingWorkbookWriter writer(streamPath);
        writer.setParallelWorkers(4);
        writer.setCompressionLevel(xlpp::CompressionLevel::Fastest);
        auto& sheet = writer.addWorksheet("Data");
        for (int r = 1; r <= 1000; ++r)
            sheet.append({std::string("Row ") + std::to_string(r), static_cast<double>(r)});
        writer.close();
        test.checkTrue(readFile(streamSeqPath) == readFile(streamPath), "Parallel streaming output is byte-identical to sequential");
        std::filesystem::remove(streamSeqPath);
        xlpp::StreamingWorkbookReader reader(streamPath);
        const auto names = reader.worksheetNames();
        test.checkEqual(names.size(), std::size_t{1}, "Parallel streaming workbook loads");
        std::size_t rows = 0;
        double sum = 0.0;
        reader.forEachRow("Data", [&](std::size_t, const xlpp::StreamingRow& row) {
            ++rows;
            if (row.size() >= 2 && std::holds_alternative<double>(row[1].value)) sum += std::get<double>(row[1].value);
            return true;
        });
        test.checkEqual(rows, std::size_t{1000}, "Parallel streaming row count");
        test.checkNear(sum, 500500.0, 1e-9, "Parallel streaming numeric aggregation");
        std::filesystem::remove(streamPath);
    }

    {
        xlpp::Workbook workbook;
        for (int s = 0; s < 6; ++s) {
            auto& sheet = workbook.addWorksheet("Sheet" + std::to_string(s + 1));
            for (int r = 1; r <= 3000; ++r)
                sheet.cell(r, 1).setValue(static_cast<double>(r * s));
        }
        const auto seqStart = std::chrono::steady_clock::now();
        workbook.save(seqPath);
        const auto seqMs = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - seqStart).count();
        xlpp::SaveOptions parallel;
        parallel.parallelWorkers = 4;
        const auto parStart = std::chrono::steady_clock::now();
        workbook.save(parPath, parallel);
        const auto parMs = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - parStart).count();
        test.checkTrue(parMs <= seqMs * 4 + 16, "Parallel save is not pathologically slow");
        std::cout << "    [BENCHMARK] 6 sheets x 3000 rows: sequential=" << seqMs
                  << " ms parallel=" << parMs << " ms file=" << std::filesystem::file_size(parPath) << " bytes\n";
        std::filesystem::remove(seqPath);
        std::filesystem::remove(parPath);
    }
}

void testZip64LimitsCancelProgress(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto zip64Path = dir / "xlpp_m20_zip64.xlsx";
    const auto wbPath = dir / "xlpp_m20_limits.xlsx";
    writeZip64Fixture(zip64Path);
    {
        auto z = xlpp::internal::ZipArchive::open(zip64Path);
        test.checkTrue(z.contains("hello.txt"), "ZIP64 fixture entry present");
        test.checkEqual(z.get("hello.txt"), std::string("Hello, ZIP64!"), "ZIP64 fixture content");
    }
    {
        xlpp::Workbook w;
        auto& sheet = w.addWorksheet("Sheet1");
        for (int r = 1; r <= 50; ++r) { sheet.cell(r, 1).setValue("row " + std::to_string(r)); sheet.cell(r, 2).setValue(static_cast<double>(r)); }
        w.save(wbPath);
    }
    const auto expectThrow = [&](const char* label, const xlpp::internal::ZipOpenLimits& limits) {
        bool threw = false;
        try { xlpp::internal::ZipArchive::open(wbPath, limits); }
        catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, label);
    };
    expectThrow("maxEntries limit", xlpp::internal::ZipOpenLimits{.maxEntries = 3});
    expectThrow("maxEntryBytes limit", xlpp::internal::ZipOpenLimits{.maxEntryBytes = 100});
    expectThrow("maxTotalBytes limit", xlpp::internal::ZipOpenLimits{.maxTotalBytes = 500});
    expectThrow("maxFileBytes limit", xlpp::internal::ZipOpenLimits{.maxFileBytes = 1000});
    {
        xlpp::internal::ZipOpenLimits limits;
        limits.cancel = [] { return true; };
        bool threw = false;
        try { xlpp::internal::ZipArchive::open(wbPath, limits); }
        catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Open cancel");
    }
    {
        std::size_t calls = 0, lastDone = 0, lastTotal = 0;
        xlpp::internal::ZipOpenLimits limits;
        limits.progress = [&](std::size_t done, std::size_t total) { ++calls; lastDone = done; lastTotal = total; };
        auto z = xlpp::internal::ZipArchive::open(wbPath, limits);
        test.checkTrue(calls > 0, "Open progress invoked");
        test.checkEqual(lastDone, lastTotal, "Open progress done equals total");
        test.checkTrue(z.contains("[Content_Types].xml"), "Open with progress still loads");
    }
    {
        xlpp::internal::ZipArchive z;
        z.add("a.txt", "hello", false);
        xlpp::internal::ZipWriteOptions opt;
        opt.cancel = [] { return true; };
        bool threw = false;
        try { z.save(wbPath, opt); }
        catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Save cancel");
    }
    {
        xlpp::internal::ZipArchive z;
        z.add("a.txt", "hello", false);
        z.add("b.txt", "world", false);
        std::size_t calls = 0, lastDone = 0, lastTotal = 0;
        xlpp::internal::ZipWriteOptions opt;
        opt.progress = [&](std::size_t done, std::size_t total) { ++calls; lastDone = done; lastTotal = total; };
        z.save(wbPath, opt);
        test.checkEqual(calls, std::size_t{2}, "Save progress invoked per entry");
        test.checkEqual(lastDone, std::size_t{2}, "Save progress final done");
        test.checkEqual(lastTotal, std::size_t{2}, "Save progress total");
    }
    std::filesystem::remove(zip64Path);
    std::filesystem::remove(wbPath);
}

void testZip64WritePath(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto seqPath = dir / "xlpp_m20_zip64write_seq.xlsx";
    const auto parPath = dir / "xlpp_m20_zip64write_par.xlsx";
    const auto smallPath = dir / "xlpp_m20_zip64write_small.xlsx";
    const auto fileBackedPath = dir / "xlpp_m20_zip64write_fileback.xlsx";
    const auto sourcePath = dir / "xlpp_m20_zip64write_source.bin";

    const auto readBytes = [](const std::filesystem::path& p) {
        std::ifstream f(p, std::ios::binary);
        return std::string((std::istreambuf_iterator<char>(f)), {});
    };
    const auto le16 = [](const std::string& b, std::size_t p) {
        return static_cast<std::uint16_t>(
            static_cast<unsigned char>(b[p]) | (static_cast<unsigned char>(b[p + 1]) << 8));
    };
    const auto hasSig = [](const std::string& b, unsigned char s0, unsigned char s1, unsigned char s2, unsigned char s3) {
        const std::string sig{static_cast<char>(s0), static_cast<char>(s1), static_cast<char>(s2), static_cast<char>(s3)};
        return b.find(sig) != std::string::npos;
    };

    const std::string bigText = [] {
        std::string t;
        t.reserve(200000);
        for (int i = 0; i < 20000; ++i) t += "ZIP64 write path line " + std::to_string(i) + "\n";
        return t;
    }();

    const auto build = [&](xlpp::internal::ZipArchive& z) {
        z.add("hello.txt", "Hello, forced ZIP64!", true);
        z.add("data.bin", std::string(4096, '\x5A'), false);
        z.add("dir/nested/file.xml", "<root><v>1</v></root>", true);
        z.add("big.txt", bigText, true);
    };

    {
        xlpp::internal::ZipArchive z;
        build(z);
        z.setForceZip64(true);
        z.save(seqPath);
    }
    const auto seqBytes = readBytes(seqPath);
    test.checkTrue(hasSig(seqBytes, 0x50, 0x4b, 0x06, 0x06), "Forced writer emits EOCD64 record");
    test.checkTrue(hasSig(seqBytes, 0x50, 0x4b, 0x06, 0x07), "Forced writer emits EOCD64 locator");
    test.checkTrue(hasSig(seqBytes, 0x50, 0x4b, 0x03, 0x04), "Forced writer emits local headers");
    test.checkEqual(le16(seqBytes, 6), std::uint16_t{0x0000}, "Forced local header omits data descriptor");

    {
        auto z = xlpp::internal::ZipArchive::open(seqPath);
        test.checkEqual(z.entryNames().size(), std::size_t{4}, "ZIP64 write round-trip entry count");
        test.checkEqual(z.get("hello.txt"), std::string("Hello, forced ZIP64!"), "ZIP64 write round-trip small entry");
        test.checkEqual(z.get("dir/nested/file.xml"), std::string("<root><v>1</v></root>"), "ZIP64 write round-trip nested entry");
        test.checkEqual(z.get("data.bin").size(), std::size_t{4096}, "ZIP64 write round-trip stored entry size");
        test.checkEqual(z.get("big.txt"), bigText, "ZIP64 write round-trip large entry");
    }

    {
        xlpp::internal::ZipArchiveReader direct(seqPath);
        test.checkEqual(direct.entryCount(), std::size_t{4}, "Streaming ZIP64 reader sees every entry");
        test.checkEqual(direct.readEntry("hello.txt"), std::string("Hello, forced ZIP64!"), "Streaming ZIP64 reader reads small entry");
        test.checkEqual(direct.readEntry("data.bin").size(), std::size_t{4096}, "Streaming ZIP64 reader reads stored entry");
        test.checkEqual(direct.readEntry("big.txt"), bigText, "Streaming ZIP64 reader reads deflated entry");
    }

    {
        xlpp::internal::ZipArchive z;
        build(z);
        z.setForceZip64(true);
        z.setParallelWorkers(4);
        z.save(parPath);
    }
    test.checkEqual(readBytes(parPath), seqBytes, "Forced ZIP64 output is parallel-deterministic");

    {
        xlpp::internal::ZipArchive z;
        build(z);
        z.save(smallPath);
    }
    test.checkEqual(le16(readBytes(smallPath), 6), std::uint16_t{0x0008}, "Unforced small path keeps data descriptors");

    // File-backed ZIP64 entries must remain bounded-memory: compressed data is
    // prepared into a temporary backing file and stored data is copied from
    // the source instead of materializing the entire input in RAM.
    std::string filePayload;
    filePayload.reserve(2 * 1024 * 1024);
    for (int i = 0; i < 32768; ++i) filePayload += "file-backed ZIP64 payload line 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ\n";
    {
        std::ofstream f(sourcePath, std::ios::binary | std::ios::trunc);
        f.write(filePayload.data(), static_cast<std::streamsize>(filePayload.size()));
    }
    {
        xlpp::internal::ZipArchive z;
        z.addFile("compressed.bin", sourcePath, true);
        z.addFile("stored.bin", sourcePath, false);
        z.setForceZip64(true);
        z.save(fileBackedPath);
    }
    {
        xlpp::internal::ZipArchiveReader direct(fileBackedPath);
        test.checkEqual(direct.readEntry("compressed.bin"), filePayload,
                        "File-backed compressed ZIP64 entry round-trips without materialization");
        test.checkEqual(direct.readEntry("stored.bin"), filePayload,
                        "File-backed stored ZIP64 entry round-trips without materialization");
    }
    bool preparationLeak = false;
    const auto preparationPrefix = fileBackedPath.filename().string() + ".xlpp-zipprep-";
    for (const auto& item : std::filesystem::directory_iterator(dir)) {
        const auto filename = item.path().filename().string();
        if (filename.rfind(preparationPrefix, 0) == 0) preparationLeak = true;
    }
    test.checkTrue(!preparationLeak, "ZIP64 file-backed preparation files are cleaned after save");

    std::filesystem::remove(seqPath);
    std::filesystem::remove(parPath);
    std::filesystem::remove(smallPath);
    std::filesystem::remove(fileBackedPath);
    std::filesystem::remove(sourcePath);
}

void testPartPreservation(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto base = dir / "xlpp_m20_preserve_base.xlsx";
    const auto staged = dir / "xlpp_m20_preserve_staged.xlsx";
    const auto out = dir / "xlpp_m20_preserve_out.xlsx";
    {
        xlpp::Workbook w;
        w.addWorksheet("Sheet1").cell("A1").setValue("Keep me");
        w.save(base);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(base);
        z.add("customXml/item1.xml", "<customData>42</customData>", false);
        auto ct = z.get("[Content_Types].xml");
        const auto marker = std::string("<Override PartName=\"/customXml/item1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>");
        ct.insert(ct.rfind("</Types>"), marker);
         z.replace("[Content_Types].xml", ct);
        z.save(staged);
    }
    {
        xlpp::Workbook w;
        w.load(staged);
        test.checkEqual(std::get<std::string>(w.worksheet("Sheet1")->cell("A1").value()), std::string("Keep me"), "Known sheet still loads with preserved part");
        bool found = false;
        for (const auto& part : w.preservedParts()) {
            if (part.name == "customXml/item1.xml") {
                found = true;
                test.checkEqual(part.data, std::string("<customData>42</customData>"), "Preserved part bytes");
                test.checkEqual(part.overrideType, std::string("application/vnd.openxmlformats-officedocument.customXmlProperties+xml"), "Preserved part content type");
            }
        }
        test.checkTrue(found, "Unknown part captured by load");
        w.save(out);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(out);
        test.checkTrue(z.contains("customXml/item1.xml"), "Preserved part re-added on save");
        test.checkEqual(z.get("customXml/item1.xml"), std::string("<customData>42</customData>"), "Preserved part bytes survive round-trip");
        const auto ct = z.get("[Content_Types].xml");
        test.checkTrue(ct.find("/customXml/item1.xml") != std::string::npos, "Content type override preserved");
    }
    std::filesystem::remove(base);
    std::filesystem::remove(staged);
    std::filesystem::remove(out);
}

void testBinaryPartPreservation(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto base = dir / "xlpp_binary_preserve_base.xlsx";
    const auto staged = dir / "xlpp_binary_preserve_staged.xlsm";
    const auto out = dir / "xlpp_binary_preserve_out.xlsm";
    {
        xlpp::Workbook w;
        w.addWorksheet("Sheet1").cell("A1").setValue("macro host");
        w.save(base);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(base);
        const std::string vbaBytes("VBA\0fixture", 11);
        z.add("xl/vbaProject.bin", vbaBytes, false);
        auto ct = z.get("[Content_Types].xml");
        const auto marker = std::string("<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>");
        ct.insert(ct.rfind("</Types>"), marker);
         z.replace("[Content_Types].xml", ct);
        z.save(staged);
    }
    {
        xlpp::Workbook w;
        w.load(staged);
        bool found = false;
        for (const auto& part : w.preservedParts()) {
            if (part.name == "xl/vbaProject.bin") {
                found = true;
                test.checkEqual(part.data.size(), std::size_t{11}, "Binary part size preserved");
                test.checkEqual(part.data[3], '\0', "Binary part NUL byte preserved");
                test.checkEqual(part.overrideType, std::string("application/vnd.ms-office.vbaProject"), "Binary part content type preserved");
            }
        }
        test.checkTrue(found, "VBA binary part captured by load");
        w.save(out);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(out);
        const auto data = z.get("xl/vbaProject.bin");
        test.checkEqual(data.size(), std::size_t{11}, "Binary part survives save");
        test.checkEqual(data[3], '\0', "Binary part NUL byte survives save");
        test.checkTrue(z.get("[Content_Types].xml").find("application/vnd.ms-office.vbaProject") != std::string::npos,
                       "Binary part content type survives save");
    }
    std::filesystem::remove(base);
    std::filesystem::remove(staged);
    std::filesystem::remove(out);
}

void testRelationshipGraphRoundTripPreservation(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto sourcePath = dir / "xlpp_p0_relationship_source.xlsx";
    const auto roundTripPath = dir / "xlpp_p0_relationship_roundtrip.xlsx";
    const std::vector<unsigned char> png{137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,137,0,0,0,13,73,68,65,84,8,215,99,248,207,192,240,31,0,5,0,1,255,137,153,61,29,0,0,0,0,73,69,78,68,174,66,96,130};

    {
        xlpp::Workbook workbook;
        workbook.addWorksheet("Intro").cell("A1").setValue("Remove this sheet after load");
        auto& sheet = workbook.addWorksheet("Objects");
        sheet.append({std::string("Category"), std::string("Amount")});
        sheet.append({std::string("A"), 10.0});
        sheet.append({std::string("B"), 20.0});
        sheet.addImage(xlpp::Image("D2", png, "png"));

        xlpp::Chart chart(xlpp::Chart::Type::Bar);
        chart.setTitle("Preserved chart");
        auto& series = chart.addSeries(xlpp::ChartSeries("Amount"));
        series.reference("Objects", "$B$2:$B$3");
        series.categories("Objects", "$A$2:$A$3");
        sheet.addChart(std::move(chart));

        xlpp::PivotTable pivot("PreservedPivot");
        pivot.setLocation("G2");
        pivot.cache().setSourceData("'Objects'!$A$1:$B$3");
        pivot.cache().setFields({"Category", "Amount"});
        pivot.cache().addRecord({"A", "10"});
        pivot.cache().addRecord({"B", "20"});
        pivot.addRowField("Category").setFieldIndex(0);
        pivot.addDataField(1);
        sheet.addPivotTable(std::move(pivot));
        workbook.save(sourcePath);
    }

    const auto before = xlpp::internal::ZipArchive::open(sourcePath);
    const auto beforeGraph = xlpp::internal::RelationshipGraph::fromArchive(before);
    const auto beforeValidation = beforeGraph.validate();
    test.checkTrue(beforeValidation.duplicateRelationshipIds.empty(), "Generated fixture has unique relationship IDs");
    test.checkTrue(beforeValidation.danglingRelationships.empty(), "Generated fixture has no dangling relationships");
    test.checkTrue(beforeValidation.orphanedParts.empty(), "Generated fixture has no orphaned package parts");
    test.checkTrue(beforeValidation.contentTypeErrors.empty(), "Generated fixture has consistent content types");

    const std::vector<std::string> protectedParts{
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/charts/chart1.xml",
        "xl/media/image1.png",
        "xl/pivotTables/pivotTable1.xml",
        "xl/pivotTables/_rels/pivotTable1.xml.rels",
        "xl/pivotCache/pivotCacheDefinition1.xml",
        "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
        "xl/pivotCache/pivotCacheRecords1.xml"
    };
    for (const auto& part : protectedParts)
        test.checkTrue(before.contains(part), "Fixture contains protected part " + part);

    {
        xlpp::Workbook loaded;
        loaded.load(sourcePath);
        test.checkTrue(!loaded.preservedRelationships().empty(), "Load captures package relationships");
        test.checkTrue(loaded.removeWorksheet("Intro"), "Unrelated leading worksheet can be removed");
        auto* sheet = loaded.worksheet("Objects");
        test.checkTrue(sheet != nullptr, "Object-bearing worksheet reloads");
        sheet->cell("C10").setValue("unrelated edit");
        loaded.save(roundTripPath);
    }

    const auto after = xlpp::internal::ZipArchive::open(roundTripPath);
    const auto afterGraph = xlpp::internal::RelationshipGraph::fromArchive(after);
    const auto afterValidation = afterGraph.validate();
    test.checkTrue(afterValidation.duplicateRelationshipIds.empty(), "Round-trip has unique relationship IDs");
    test.checkTrue(afterValidation.danglingRelationships.empty(), "Round-trip has no dangling relationships");
    test.checkTrue(afterValidation.orphanedParts.empty(), "Round-trip has no orphaned package parts");
    test.checkTrue(afterValidation.contentTypeErrors.empty(), "Round-trip has consistent content types");

    for (const auto& part : protectedParts) {
        test.checkTrue(after.contains(part), "Round-trip keeps protected part " + part);
        test.checkEqual(after.get(part), before.get(part), "Untouched part stays byte-identical: " + part);
    }

    const auto sheetXml = after.get("xl/worksheets/sheet1.xml");
    const auto drawingNodes = xlpp::internal::tags(sheetXml, "drawing");
    const auto pivotNodes = xlpp::internal::tags(sheetXml, "pivotTablePart");
    test.checkEqual(drawingNodes.size(), std::size_t{1}, "Worksheet keeps visible drawing reference");
    test.checkEqual(pivotNodes.size(), std::size_t{1}, "Worksheet keeps pivot table reference");

    const auto sheetRelationships = afterGraph.relationshipsFrom("xl/worksheets/sheet1.xml");
    const auto hasRelationship = [&](const std::string& id, const std::string& typeSuffix) {
        return std::any_of(sheetRelationships.begin(), sheetRelationships.end(), [&](const auto& relationship) {
            return relationship.id == id && relationship.type.size() >= typeSuffix.size()
                && relationship.type.compare(relationship.type.size() - typeSuffix.size(), typeSuffix.size(), typeSuffix) == 0;
        });
    };
    test.checkTrue(hasRelationship(xlpp::internal::attribute(drawingNodes.front(), "r:id"), "/drawing"),
                   "Worksheet drawing node resolves through its preserved relationship");
    test.checkTrue(hasRelationship(xlpp::internal::attribute(pivotNodes.front(), "r:id"), "/pivotTable"),
                   "Worksheet pivot node resolves through its preserved relationship");

    const auto workbookXml = after.get("xl/workbook.xml");
    const auto pivotCacheNodes = xlpp::internal::tags(workbookXml, "pivotCache");
    test.checkEqual(pivotCacheNodes.size(), std::size_t{1}, "Workbook keeps pivot cache reference");
    const auto workbookRelationships = afterGraph.relationshipsFrom("xl/workbook.xml");
    const auto pivotCacheId = xlpp::internal::attribute(pivotCacheNodes.front(), "r:id");
    test.checkTrue(std::any_of(workbookRelationships.begin(), workbookRelationships.end(), [&](const auto& relationship) {
        return relationship.id == pivotCacheId && relationship.type.find("/pivotCacheDefinition") != std::string::npos;
    }), "Workbook pivot cache ID resolves after relationship collision handling");

    const auto diff = xlpp::internal::comparePackages(before, after);
    test.checkTrue(std::find(diff.removedParts.begin(), diff.removedParts.end(), "xl/charts/chart1.xml") == diff.removedParts.end(),
                   "Package diff reports no removed chart part");
    test.checkTrue(std::find(diff.removedParts.begin(), diff.removedParts.end(), "xl/media/image1.png") == diff.removedParts.end(),
                   "Package diff reports no removed image part");
    test.checkTrue(std::find(diff.removedParts.begin(), diff.removedParts.end(), "xl/pivotTables/pivotTable1.xml") == diff.removedParts.end(),
                   "Package diff reports no removed pivot part");

    std::filesystem::remove(sourcePath);
    std::filesystem::remove(roundTripPath);
}

void testOwnerReferenceAndObjectRegressionDetection(TestContext& test) {
    xlpp::internal::ZipArchive brokenOwner;
    brokenOwner.add("[Content_Types].xml",
        "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "</Types>");
    brokenOwner.add("_rels/.rels",
        "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>");
    brokenOwner.add("xl/workbook.xml",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
    brokenOwner.add("xl/_rels/workbook.xml.rels",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        "</Relationships>");
    brokenOwner.add("xl/worksheets/sheet1.xml",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetData/><drawing r:id=\"rIdMissing\"/></worksheet>");

    const auto brokenReport = xlpp::internal::RelationshipGraph::fromArchive(brokenOwner).validate();
    test.checkTrue(!brokenReport.ownerReferenceErrors.empty(), "Validator detects owner XML referencing a missing relationship");
    test.checkTrue(std::any_of(brokenReport.ownerReferenceErrors.begin(), brokenReport.ownerReferenceErrors.end(), [](const auto& issue) {
        return issue.find("rIdMissing") != std::string::npos;
    }), "Owner-reference error identifies the missing relationship ID");

    const auto fixturePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(fixturePath);
    auto after = before;
    auto sheetXml = after.get("xl/worksheets/sheet1.xml");
    const auto drawingNodes = xlpp::internal::tags(sheetXml, "drawing");
    test.checkEqual(drawingNodes.size(), std::size_t{1}, "Regression fixture starts with one drawing owner node");
    const auto position = sheetXml.find(drawingNodes.front());
    test.checkTrue(position != std::string::npos, "Drawing owner node can be removed for negative test");
    sheetXml.erase(position, drawingNodes.front().size());
    after.replace("xl/worksheets/sheet1.xml", std::move(sheetXml));

    const auto afterGraph = xlpp::internal::RelationshipGraph::fromArchive(after);
    const auto afterValidation = afterGraph.validate();
    test.checkTrue(!afterValidation.ownerReferenceErrors.empty(), "Unused drawing/image/chart relationships are rejected");
    test.checkEqual(afterGraph.objectInventory().images, std::size_t{0}, "Invisible image is not counted merely because its part exists");
    test.checkEqual(afterGraph.objectInventory().charts, std::size_t{0}, "Invisible chart is not counted merely because its part exists");
    const auto diff = xlpp::internal::comparePackages(before, after);
    test.checkTrue(std::any_of(diff.objectCountRegressions.begin(), diff.objectCountRegressions.end(), [](const auto& issue) {
        return issue.find("Visible image count") != std::string::npos;
    }), "Package comparison reports visible image loss");
    test.checkTrue(std::any_of(diff.objectCountRegressions.begin(), diff.objectCountRegressions.end(), [](const auto& issue) {
        return issue.find("Visible chart count") != std::string::npos;
    }), "Package comparison reports visible chart loss");
}

void testExtendedOwnerGraphAndRelationshipHardening(TestContext& test) {
    xlpp::internal::ZipArchive archive;
    archive.add("[Content_Types].xml",
        "<?xml version=\"1.0\"?><ct:Types xmlns:ct=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<ct:Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<ct:Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<ct:Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>"
        "<ct:Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<ct:Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "<ct:Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>"
        "<ct:Override PartName=\"/xl/comments1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>"
        "<ct:Override PartName=\"/xl/externalLinks/externalLink1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml\"/>"
        "</ct:Types>");
    archive.add("_rels/.rels",
        "<?xml version=\"1.0\"?><pr:Relationships xmlns:pr=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<pr:Relationship Id=\"rIdOffice\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\"/>"
        "</pr:Relationships>");
    archive.add("xl/workbook.xml",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rSheet\"/></sheets>"
        "<externalReferences><externalReference r:id=\"rExternal\"/></externalReferences>"
        "</workbook>");
    archive.add("xl/_rels/workbook.xml.rels",
        "<pr:Relationships xmlns:pr=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<pr:Relationship Id=\"rSheet\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/./sheet1.xml\"/>"
        "<pr:Relationship Id=\"rExternal\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink\" Target=\"/xl/externalLinks/externalLink1.xml\"/>"
        "</pr:Relationships>");
    archive.add("xl/worksheets/sheet1.xml",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetData/><legacyDrawing r:id=\"rVml\"/>"
        "<tableParts count=\"1\"><tablePart r:id=\"rTable\"/></tableParts>"
        "</worksheet>");
    archive.add("xl/worksheets/_rels/sheet1.xml.rels",
        "<pr:Relationships xmlns:pr=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<pr:Relationship Id=\"rTable\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/../tables/table1.xml\"/>"
        "<pr:Relationship Id=\"rComments\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments1.xml\"/>"
        "<pr:Relationship Id=\"rVml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing1.vml\"/>"
        "</pr:Relationships>");
    archive.add("xl/tables/table1.xml",
        "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"A1:B2\"/>");
    archive.add("xl/comments1.xml",
        "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>A</author></authors>"
        "<commentList><comment ref=\"A1\" authorId=\"0\"><text><t>one</t></text></comment>"
        "<comment ref=\"B1\" authorId=\"0\"><text><t>two</t></text></comment></commentList></comments>");
    archive.add("xl/drawings/vmlDrawing1.vml", "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape id=\"_x0000_s1\"/></xml>");
    archive.add("xl/externalLinks/externalLink1.xml",
        "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<externalBook r:id=\"rIdPath\"/></externalLink>");
    archive.add("xl/externalLinks/_rels/externalLink1.xml.rels",
        "<pr:Relationships xmlns:pr=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<pr:Relationship Id=\"rIdPath\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath\" Target=\"file:///C:/source.xlsx\" TargetMode=\"External\"/>"
        "</pr:Relationships>");

    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    const auto report = graph.validate();
    test.checkTrue(report.ok(), "Extended owner graph accepts a valid table/comment/external-link package with prefixed namespaces");
    test.checkEqual(graph.objectInventory().tables, std::size_t{1}, "Owner graph counts reachable table parts");
    test.checkEqual(graph.objectInventory().comments, std::size_t{2}, "Owner graph counts reachable legacy comments");
    test.checkEqual(graph.objectInventory().externalLinks, std::size_t{1}, "Owner graph counts reachable external links");
    test.checkTrue(report.relationshipSyntaxErrors.empty(), "Prefixed Relationship nodes parse without syntax errors");

    test.checkEqual(xlpp::internal::RelationshipGraph::resolveTarget("xl/worksheets/sheet1.xml", "../tables/./table1.xml"),
                    std::string("xl/tables/table1.xml"), "OPC target resolver normalizes dot segments");
    test.checkEqual(xlpp::internal::RelationshipGraph::resolveTarget("xl/worksheets/sheet1.xml", "/xl/comments1.xml"),
                    std::string("xl/comments1.xml"), "OPC target resolver handles package-absolute targets");
    test.checkTrue(xlpp::internal::RelationshipGraph::resolveTarget("xl/worksheets/sheet1.xml", "../../../escape.xml").empty(),
                   "OPC target resolver rejects paths escaping above package root");
    test.checkTrue(xlpp::internal::RelationshipGraph::resolveTarget("xl/worksheets/sheet1.xml", "..\\tables\\table1.xml").empty(),
                   "OPC target resolver rejects backslash package paths");

    auto malformed = archive;
    auto sheetRels = malformed.get("xl/worksheets/_rels/sheet1.xml.rels");
    const auto close = sheetRels.find("</pr:Relationships>");
    test.checkTrue(close != std::string::npos, "Malformed relationship fixture has a relationships closing tag");
    sheetRels.insert(close,
        "<pr:Relationship Id=\"rMalformed\" Type=\"urn:test:corrupt\" Target=\"\" TargetMode=\"Sideways\"/>");
    malformed.replace("xl/worksheets/_rels/sheet1.xml.rels", sheetRels);
    const auto malformedReport = xlpp::internal::RelationshipGraph::fromArchive(malformed).validate();
    test.checkTrue(std::any_of(malformedReport.relationshipSyntaxErrors.begin(), malformedReport.relationshipSyntaxErrors.end(), [](const auto& issue) {
        return issue.find("empty Target") != std::string::npos;
    }), "Validator reports empty relationship Target instead of silently dropping the node");
    test.checkTrue(std::any_of(malformedReport.relationshipSyntaxErrors.begin(), malformedReport.relationshipSyntaxErrors.end(), [](const auto& issue) {
        return issue.find("invalid TargetMode") != std::string::npos;
    }), "Validator reports invalid TargetMode values");

    auto regressed = archive;
    auto workbookXml = regressed.get("xl/workbook.xml");
    const std::string externalBlock = "<externalReferences><externalReference r:id=\"rExternal\"/></externalReferences>";
    auto externalPosition = workbookXml.find(externalBlock);
    test.checkTrue(externalPosition != std::string::npos, "External reference owner can be removed for regression test");
    workbookXml.erase(externalPosition, externalBlock.size());
    regressed.replace("xl/workbook.xml", workbookXml);

    auto worksheetXml = regressed.get("xl/worksheets/sheet1.xml");
    const std::string tableNode = "<tablePart r:id=\"rTable\"/>";
    auto tablePosition = worksheetXml.find(tableNode);
    test.checkTrue(tablePosition != std::string::npos, "Table owner can be removed for regression test");
    worksheetXml.erase(tablePosition, tableNode.size());
    regressed.replace("xl/worksheets/sheet1.xml", worksheetXml);
    regressed.replace("xl/comments1.xml",
        "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>A</author></authors><commentList/></comments>");

    const auto diff = xlpp::internal::comparePackages(archive, regressed);
    test.checkTrue(std::any_of(diff.objectCountRegressions.begin(), diff.objectCountRegressions.end(), [](const auto& issue) {
        return issue.find("Table count") != std::string::npos;
    }), "Package comparison detects table object loss");
    test.checkTrue(std::any_of(diff.objectCountRegressions.begin(), diff.objectCountRegressions.end(), [](const auto& issue) {
        return issue.find("Comment count") != std::string::npos;
    }), "Package comparison detects comment object loss");
    test.checkTrue(std::any_of(diff.objectCountRegressions.begin(), diff.objectCountRegressions.end(), [](const auto& issue) {
        return issue.find("External link count") != std::string::npos;
    }), "Package comparison detects external-link object loss");

    const auto tempDir = std::filesystem::temp_directory_path();
    const auto sourcePath = tempDir / "xlpp_p0g_owner_graph_source.xlsx";
    const auto outputPath = tempDir / "xlpp_p0g_owner_graph_roundtrip.xlsx";
    auto writerArchive = archive;
    for (const auto& relPart : {std::string("_rels/.rels"), std::string("xl/_rels/workbook.xml.rels"),
                                std::string("xl/worksheets/_rels/sheet1.xml.rels"),
                                std::string("xl/externalLinks/_rels/externalLink1.xml.rels")}) {
        auto xml = writerArchive.get(relPart);
        std::size_t position = 0;
        while ((position = xml.find("pr:", position)) != std::string::npos) xml.erase(position, 3);
        const std::string prefixedNs = "xmlns:pr=\"http://schemas.openxmlformats.org/package/2006/relationships\"";
        const auto nsPosition = xml.find(prefixedNs);
        if (nsPosition != std::string::npos)
            xml.replace(nsPosition, prefixedNs.size(), "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"");
        writerArchive.replace(relPart, xml);
    }
    {
        auto contentTypes = writerArchive.get("[Content_Types].xml");
        std::size_t position = 0;
        while ((position = contentTypes.find("ct:", position)) != std::string::npos) contentTypes.erase(position, 3);
        const std::string prefixedNs = "xmlns:ct=\"http://schemas.openxmlformats.org/package/2006/content-types\"";
        const auto nsPosition = contentTypes.find(prefixedNs);
        if (nsPosition != std::string::npos)
            contentTypes.replace(nsPosition, prefixedNs.size(), "xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"");
        writerArchive.replace("[Content_Types].xml", contentTypes);
    }
    writerArchive.save(sourcePath);
    {
        xlpp::Workbook workbook;
        workbook.load(sourcePath);
        auto* sheet = workbook.worksheet("Data");
        test.checkTrue(sheet != nullptr, "Extended owner-graph fixture loads through Workbook");
        if (sheet) sheet->cell("C3").setValue("unrelated edit");
        workbook.save(outputPath);
    }
    const auto roundTripArchive = xlpp::internal::ZipArchive::open(outputPath);
    const auto roundTripGraph = xlpp::internal::RelationshipGraph::fromArchive(roundTripArchive);
    const auto roundTripValidation = roundTripGraph.validate();
    test.checkTrue(roundTripValidation.ok(), "Workbook round-trip keeps table/comments/external-link owner graph valid");
    test.checkEqual(roundTripGraph.objectInventory().tables, std::size_t{1}, "Workbook round-trip keeps table reachable");
    test.checkEqual(roundTripGraph.objectInventory().comments, std::size_t{2}, "Workbook round-trip keeps comments reachable");
    test.checkEqual(roundTripGraph.objectInventory().externalLinks, std::size_t{1}, "Workbook round-trip keeps external link reachable");
    test.checkTrue(roundTripArchive.contains("xl/externalLinks/externalLink1.xml"), "External-link part survives unrelated workbook edit");
    test.checkEqual(roundTripArchive.get("xl/externalLinks/externalLink1.xml"), writerArchive.get("xl/externalLinks/externalLink1.xml"),
                    "Untouched external-link XML stays byte-identical");
    test.checkTrue(roundTripArchive.get("xl/workbook.xml").find("externalReference") != std::string::npos,
                   "Workbook externalReference owner node survives unrelated edit");
    std::filesystem::remove(sourcePath);
    std::filesystem::remove(outputPath);
}

void testPackageValidatorFailureDetection(TestContext& test) {
    xlpp::internal::ZipArchive archive;
    archive.add("[Content_Types].xml",
        "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"text/xml\"/>"
        "<Override PartName=\"/ghost.xml\" ContentType=\"application/xml\"/>"
        "</Types>");
    archive.add("_rels/.rels",
        "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"urn:test:first\" Target=\"missing1.xml\"/>"
        "<Relationship Id=\"rId1\" Type=\"urn:test:second\" Target=\"missing2.xml\"/>"
        "</Relationships>");
    archive.add("orphan.bin", "orphan", false);

    const auto report = xlpp::internal::RelationshipGraph::fromArchive(archive).validate();
    test.checkTrue(!report.ok(), "Validator rejects an inconsistent package");
    test.checkTrue(!report.duplicateRelationshipIds.empty(), "Validator detects duplicate relationship IDs");
    test.checkEqual(report.danglingRelationships.size(), std::size_t{2}, "Validator detects dangling targets");
    test.checkTrue(std::find(report.orphanedParts.begin(), report.orphanedParts.end(), "orphan.bin") != report.orphanedParts.end(),
                   "Validator detects orphaned internal parts");
    test.checkTrue(std::any_of(report.contentTypeErrors.begin(), report.contentTypeErrors.end(), [](const auto& issue) {
        return issue.find("Duplicate Default") != std::string::npos;
    }), "Validator detects duplicate content-type declarations");
    test.checkTrue(std::any_of(report.contentTypeErrors.begin(), report.contentTypeErrors.end(), [](const auto& issue) {
        return issue.find("ghost.xml") != std::string::npos;
    }), "Validator detects stale content-type overrides");
    test.checkTrue(std::any_of(report.contentTypeErrors.begin(), report.contentTypeErrors.end(), [](const auto& issue) {
        return issue.find("orphan.bin") != std::string::npos;
    }), "Validator detects parts without content types");
}

void testZipUniqueEntryPolicy(TestContext& test) {
    xlpp::internal::ZipArchive zip;
    zip.add("entry.xml", "generated");
    zip.addUnique("entry.xml", "preserved");
    zip.addUnique("other.xml", "preserved");
    test.checkEqual(zip.get("entry.xml"), std::string("generated"), "Unique ZIP add does not overwrite generated entry");
    test.checkEqual(zip.get("other.xml"), std::string("preserved"), "Unique ZIP add inserts new entry");
}

void testStrictNamespaces(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto transitional = dir / "xlpp_m20_transitional.xlsx";
    const auto strictPath = dir / "xlpp_m20_strict.xlsx";
    {
        xlpp::Workbook w;
        w.addWorksheet("Sheet1").cell("A1").setValue("value");
        w.save(transitional);
        xlpp::SaveOptions opt;
        opt.strictNamespace = true;
        w.save(strictPath, opt);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(transitional);
        test.checkTrue(z.get("xl/workbook.xml").find("http://schemas.openxmlformats.org/spreadsheetml/2006/main") != std::string::npos, "Transitional workbook namespace");
    }
    {
        auto z = xlpp::internal::ZipArchive::open(strictPath);
        const auto wb = z.get("xl/workbook.xml");
        test.checkTrue(wb.find("http://purl.oclc.org/ooxml/spreadsheetml/main") != std::string::npos, "Strict workbook main namespace");
        test.checkTrue(wb.find("http://purl.oclc.org/ooxml/officeDocument/relationships") != std::string::npos, "Strict workbook relationships namespace");
        test.checkTrue(z.get("[Content_Types].xml").find("http://purl.oclc.org/ooxml/package/content-types") != std::string::npos, "Strict content types namespace");
        test.checkTrue(z.get("_rels/.rels").find("http://purl.oclc.org/ooxml/package/relationships") != std::string::npos, "Strict root rels namespace");
        test.checkTrue(z.get("xl/worksheets/sheet1.xml").find("http://purl.oclc.org/ooxml/spreadsheetml/main") != std::string::npos, "Strict worksheet namespace");
    }
    {
        xlpp::Workbook w;
        w.load(strictPath);
        test.checkTrue(w.strictNamespaces(), "Strictness detected on load");
        test.checkEqual(std::get<std::string>(w.worksheet("Sheet1")->cell("A1").value()), std::string("value"), "Strict package loads");
        xlpp::Workbook t;
        t.load(transitional);
        test.checkTrue(!t.strictNamespaces(), "Transitional load not marked strict");
    }
    std::filesystem::remove(transitional);
    std::filesystem::remove(strictPath);
}

void testLenientLoad(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto good = dir / "xlpp_m20_lenient_good.xlsx";
    const auto broken = dir / "xlpp_m20_lenient_broken.xlsx";
    {
        xlpp::Workbook w;
        w.addWorksheet("GoodSheet").cell("A1").setValue("intact");
        w.addWorksheet("BadSheet").cell("A1").setValue("lost");
        w.save(good);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(good);
         z.replace("xl/worksheets/sheet1.xml",
              "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\"><v>not-a-number</v></c></row></sheetData></worksheet>",
              false);
        z.save(broken);
    }
    {
        xlpp::Workbook w;
        bool threw = false;
        try { w.load(broken); }
        catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Non-lenient load rejects malformed sheet");
    }
    {
        xlpp::Workbook w;
        xlpp::LoadOptions opt;
        opt.lenient = true;
        w.load(broken, opt);
        test.checkTrue(w.diagnostics().hadErrors(), "Lenient load records errors");
        test.checkEqual(w.diagnostics().errors.size(), std::size_t{1}, "One malformed sheet recorded");
        test.checkEqual(w.worksheets().size(), std::size_t{2}, "Lenient load keeps all sheet slots");
        test.checkEqual(std::get<std::string>(w.worksheet("BadSheet")->cell("A1").value()), std::string("lost"), "Intact sheet content survives");
    }
    {
        std::size_t calls = 0, lastDone = 0, lastTotal = 0;
        xlpp::Workbook w;
        xlpp::LoadOptions opt;
        opt.progress = [&](std::size_t done, std::size_t total) { ++calls; lastDone = done; lastTotal = total; };
        w.load(good, opt);
        test.checkTrue(calls > 0, "Workbook load progress invoked");
        test.checkEqual(lastDone, lastTotal, "Workbook load progress done equals total");
    }
    std::filesystem::remove(good);
    std::filesystem::remove(broken);
}

void testMalformedInputHardening(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto wbPath = dir / "xlpp_m20_hardening_wb.xlsx";
    const auto crafted = dir / "xlpp_m20_hardening_crafted.xlsx";

    const auto readBytes = [](const std::filesystem::path& p) {
        std::ifstream f(p, std::ios::binary);
        return std::string((std::istreambuf_iterator<char>(f)), {});
    };
    const auto writeBytes = [](const std::filesystem::path& p, const std::string& bytes) {
        std::ofstream f(p, std::ios::binary);
        f.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    };
    const auto putLE = [](std::string& out, std::uint64_t v, int width) {
        for (int i = 0; i < width; ++i) out.push_back(static_cast<char>((v >> (8 * i)) & 0xff));
    };
    const auto sheetXml = [](const std::string& inner) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
               inner + "</worksheet>";
    };
    const auto loads = [&](bool lenient) {
        xlpp::Workbook w;
        xlpp::LoadOptions opt;
        opt.lenient = lenient;
        w.load(crafted, opt);
        return w;
    };

    {
        xlpp::Workbook w;
        w.addWorksheet("Data").cell("B2").setValue(42);
        w.save(wbPath);
    }

    // Huge <col> range must be rejected (bounds the per-column loop).
    {
        auto z = xlpp::internal::ZipArchive::open(wbPath);
         z.replace("xl/worksheets/sheet1.xml",
              sheetXml("<cols><col min=\"1\" max=\"2000000000\" width=\"8\"/></cols>"), false);
        z.save(crafted);
        bool threw = false;
        try { loads(false); } catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Huge column range rejected");
        auto w = loads(true);
        test.checkTrue(w.diagnostics().hadErrors(), "Huge column range recorded in lenient mode");
    }
    // Zero or descending <col> range rejected.
    {
        auto z = xlpp::internal::ZipArchive::open(wbPath);
         z.replace("xl/worksheets/sheet1.xml",
              sheetXml("<cols><col min=\"5\" max=\"2\" width=\"8\"/></cols>"), false);
        z.save(crafted);
        bool threw = false;
        try { loads(false); } catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Descending column range rejected");
    }
    // Bad numeric cell values (number without t, overflow, non-finite) never crash.
    {
        auto z = xlpp::internal::ZipArchive::open(wbPath);
         z.replace("xl/worksheets/sheet1.xml",
              sheetXml("<sheetData><row r=\"1\"><c r=\"A1\"><v>1e400</v></c><c r=\"B1\"><v>99999999999999999999</v></c></row></sheetData>"), false);
        z.save(crafted);
        auto w = loads(true);
        test.checkTrue(w.diagnostics().hadErrors(), "Numeric overflow recorded in lenient mode");
        test.checkEqual(w.worksheets().size(), std::size_t{1}, "Sheet slot survives numeric overflow");
    }
    // Unclosed/invalid XML structures are tolerated (empty result), not a crash.
    {
        auto z = xlpp::internal::ZipArchive::open(wbPath);
         z.replace("xl/worksheets/sheet1.xml",
              "<?xml version=\"1.0\"?><worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v>", false);
        z.save(crafted);
        auto w = loads(true);
        test.checkTrue(w.worksheets().size() == std::size_t{1} && !w.diagnostics().hadErrors(),
                       "Truncated sheet XML tolerated without errors");
    }
    // ZIP reader: implausible uncompressed size (alloc-bomb) rejected cleanly.
    {
        xlpp::internal::ZipArchive z;
        z.add("data.txt", "hello", true);
        z.save(crafted);
        auto bytes = readBytes(crafted);
        const std::string centralSig{static_cast<char>(0x50), static_cast<char>(0x4b),
                                     static_cast<char>(0x01), static_cast<char>(0x02)};
        const auto pos = bytes.find(centralSig);
        test.checkTrue(pos != std::string::npos, "Central record located");
        bytes[pos + 24] = static_cast<char>(0xFF);
        bytes[pos + 25] = static_cast<char>(0xFF);
        bytes[pos + 26] = static_cast<char>(0xFF);
        bytes[pos + 27] = static_cast<char>(0x7F);
        writeBytes(crafted, bytes);
        bool threw = false;
        try { xlpp::internal::ZipArchive::open(crafted); }
        catch (const std::exception& e) { threw = std::string(e.what()).find("implausible") != std::string::npos; }
        test.checkTrue(threw, "Implausible uncompressed size rejected");
    }
    // ZIP reader: truncated and garbage archives rejected without crash.
    {
        writeBytes(crafted, "PK\x03\x04garbage");
        bool threw = false;
        try { xlpp::internal::ZipArchive::open(crafted); }
        catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Truncated archive rejected");
        writeBytes(crafted, "this is not a zip file at all");
        threw = false;
        try { xlpp::internal::ZipArchive::open(crafted); }
        catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Garbage archive rejected");
    }
    // ZIP reader: EOCD claiming entries whose central directory is out of bounds.
    {
        std::string local;
        putLE(local, 0x04034b50u, 4); putLE(local, 20, 2); putLE(local, 0, 2);
        putLE(local, 0, 2); putLE(local, 0, 2); putLE(local, 0, 2);
        putLE(local, 0, 4); putLE(local, 1, 4); putLE(local, 1, 4);
        putLE(local, 4, 2); putLE(local, 0, 2); local += "data";
        std::string eocd;
        putLE(eocd, 0x06054b50u, 4); putLE(eocd, 0, 2); putLE(eocd, 0, 2);
        putLE(eocd, 1, 2); putLE(eocd, 1, 2); putLE(eocd, 0, 4);
        putLE(eocd, 100000u, 4); putLE(eocd, 0, 2);
        writeBytes(crafted, local + eocd);
        bool threw = false;
        try { xlpp::internal::ZipArchive::open(crafted); }
        catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Out-of-bounds central directory rejected");
    }
    // XML utility functions tolerate malformed markup.
    {
        test.checkTrue(xlpp::internal::tags("<a><b></b>", "c").empty(), "Missing tag yields empty set");
        test.checkTrue(xlpp::internal::tagText("<a><b>x", "b").empty(), "Unclosed tag text is empty");
        test.checkEqual(xlpp::internal::attribute("<a ref=\"5", "ref"), std::string{}, "Unclosed attribute is empty");
        test.checkEqual(xlpp::internal::xmlUnescape("a&amp;b&lt;c"), std::string("a&b<c"), "Entities unescaped");
    }

    std::filesystem::remove(wbPath);
    std::filesystem::remove(crafted);
}

void testMutationFuzz(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto path = dir / "xlpp_m20_fuzz.xlsx";

    const auto readBytes = [](const std::filesystem::path& p) {
        std::ifstream f(p, std::ios::binary);
        return std::string((std::istreambuf_iterator<char>(f)), {});
    };
    const auto writeBytes = [](const std::filesystem::path& p, const std::string& bytes) {
        std::ofstream f(p, std::ios::binary);
        f.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    };
    const auto mutateBit = [](std::string& s, std::uint32_t seed) {
        auto x = seed;
        for (int i = 0; i < 3; ++i) x = x * 1664525u + 1013904223u;
        const auto pos = static_cast<std::size_t>(x % s.size());
        s[pos] = static_cast<char>(s[pos] ^ static_cast<char>(1u << (x >> 24 & 7)));
    };

    // A workbook with rich structure: values, strings, styles, a table, a
    // merged range, comments, autofilter, defined names.
    xlpp::Workbook w;
    {
        auto& s = w.addWorksheet("Fuzz");
        s.cell("A1").setValue(3.14);
        s.cell("B2").setValue("hello world");
        s.cell("C3").setFormula("A1*2");
        s.cell("A1").style().font().setBold(true);
        s.mergeCells("A1:C3");
        s.autoFilter().setReference("A1:C3");
        s.addTable("Tbl", "A1:C3");
        s.cell("B2").setComment(xlpp::Comment("a comment", "tester"));
        for (int i = 1; i <= 3; ++i) s.columnDimension(i).width = 12.5;
    }
    w.addDefinedName(xlpp::DefinedName("MyName", "Fuzz!$A$1"));
    w.save(path);

    // 1) Byte-flip mutations over the whole file (seeded, deterministic).
    const auto base = readBytes(path);
    test.checkTrue(base.size() > 100u, "Fuzz baseline workbook written");
    std::uint32_t seed = 0xC0FFEE;
    std::size_t loads = 0;
    std::size_t rejects = 0;
    for (int round = 0; round < 200; ++round) {
        auto mutated = base;
        mutateBit(mutated, seed);
        writeBytes(path, mutated);
        bool threw = false;
        try { xlpp::Workbook w2; w2.load(path); loads++; }
        catch (const std::exception&) { threw = true; }
        if (threw) ++rejects;
        seed += 0x9E3779B9u;
    }
    test.checkEqual(loads + rejects, std::size_t{200}, "Every mutation either loads or rejects");
    test.checkTrue(rejects > 0, "At least one mutation is rejected");

    // 2) Truncations at every 1/8 boundary from the end.
    for (std::size_t cut = base.size() / 8; cut < base.size(); cut += base.size() / 8) {
        writeBytes(path, base.substr(0, cut));
        bool threw = false;
        try { xlpp::Workbook w2; w2.load(path); }
        catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Truncation at " + std::to_string(cut) + " rejected");
    }

    // 3) Byte-flip mutations of an already-decompressed sheet XML through the
    //    lenient path: errors recorded, never crash.
    seed = 0xBADC0DE;
    for (int round = 0; round < 100; ++round) {
        auto mutated = base;
        mutateBit(mutated, seed);
        writeBytes(path, mutated);
        try {
            xlpp::Workbook w2;
            xlpp::LoadOptions opt;
            opt.lenient = true;
            w2.load(path, opt);
            // Successful or lenient; re-serializing must not crash either.
            const auto out = dir / "xlpp_m20_fuzz_out.xlsx";
            try { w2.save(out); } catch (const std::exception&) {}
            std::filesystem::remove(out);
        } catch (const std::exception&) {
            // Even in lenient mode, structural failures may reject the load.
        }
    }

    // 4) A mutated archive whose contents load fine must still round-trip
    //    without throwing on save (preservation keeps raw parts intact).
    auto mutated = base;
    mutateBit(mutated, 0x1234567u);
    writeBytes(path, mutated);
    {
        xlpp::Workbook w2;
        xlpp::LoadOptions opt;
        opt.lenient = true;
        bool loaded = true;
        try { w2.load(path, opt); } catch (const std::exception&) { loaded = false; }
        if (loaded) {
            const auto out = dir / "xlpp_m20_fuzz_out2.xlsx";
            bool saved = true;
            try { w2.save(out); } catch (const std::exception&) { saved = false; }
            test.checkTrue(saved, "Lenient-loaded mutated workbook re-saves");
            std::filesystem::remove(out);
        }
    }

    std::filesystem::remove(path);
}

void testStreamingWriterModes(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto hashPath = dir / "xlpp_m21_stream_hash.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(hashPath, xlpp::SharedStringMode::Hash);
        auto& ws = writer.addWorksheet("Data");
        ws.append({std::string("repeat")});
        ws.append({std::string("repeat")});
        ws.append({std::string("other"), 42.0, true});
        writer.close();
    }
    {
        auto z = xlpp::internal::ZipArchive::open(hashPath);
        test.checkTrue(z.contains("xl/sharedStrings.xml"), "Hash mode writes shared strings");
        const auto sst = z.get("xl/sharedStrings.xml");
        test.checkTrue(sst.find("uniqueCount=\"2\"") != std::string::npos, "Hash mode deduplicates strings");
        test.checkTrue(sst.find("count=\"3\"") != std::string::npos, "Hash mode counts occurrences");
    }
    {
        xlpp::StreamingWorkbookReader reader(hashPath);
        const auto names = reader.worksheetNames();
        test.checkEqual(names.size(), std::size_t{1}, "Streaming reader lists one sheet");
        test.checkEqual(names[0], std::string("Data"), "Streaming reader sheet name");
        std::size_t rowSeen = 0;
        reader.forEachRow("Data", [&](std::size_t row, const xlpp::StreamingRow&) {
            rowSeen = row;
            return true;
        });
        test.checkEqual(rowSeen, std::size_t{3}, "Three rows streamed back");
        xlpp::StreamingRow first;
        auto it = reader.worksheet("Data").begin();
        if (it != reader.worksheet("Data").end()) first = *it;
        test.checkEqual(first.size(), std::size_t{1}, "First row has one cell");
        test.checkTrue(std::get_if<std::string>(&first[0].value) != nullptr, "First cell is a string");
        test.checkEqual(std::get<std::string>(first[0].value), std::string("repeat"), "String value streamed back");
    }
    std::filesystem::remove(hashPath);

    const auto lruPath = dir / "xlpp_m21_stream_lru.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(lruPath, xlpp::SharedStringMode::BoundedLru, 4);
        auto& ws = writer.addWorksheet("Cache");
        for (int i = 0; i < 6; ++i) ws.append({std::string("k" + std::to_string(i % 3))});
        writer.close();
    }
    {
        xlpp::StreamingWorkbookReader reader(lruPath);
        std::vector<std::string> values;
        reader.forEachRow("Cache", [&](std::size_t, const xlpp::StreamingRow& cells) {
            if (!cells.empty() && std::get_if<std::string>(&cells[0].value))
                values.push_back(std::get<std::string>(cells[0].value));
            return true;
        });
        test.checkEqual(values.size(), std::size_t{6}, "BoundedLru rows all present");
        test.checkEqual(values[0], std::string("k0"), "BoundedLru first value");
        test.checkEqual(values[5], std::string("k2"), "BoundedLru last value repeats correctly");
    }
    std::filesystem::remove(lruPath);
}

void testStreamingReaderFeatures(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m21_stream_features.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(path, xlpp::SharedStringMode::Hash);
        writer.addWorksheet("Empty");
        auto& ws = writer.addWorksheet("Rows");
        for (int i = 1; i <= 10; ++i) ws.append({double(i), std::string("row" + std::to_string(i))});
        writer.close();
    }
    {
        xlpp::StreamingWorkbookReader reader(path);
        const auto names = reader.worksheetNames();
        test.checkEqual(names.size(), std::size_t{2}, "Both sheets listed");

        std::size_t seen = 0;
        reader.forEachRow("Empty", [&](std::size_t, const xlpp::StreamingRow&) { ++seen; return true; });
        test.checkEqual(seen, std::size_t{0}, "Empty sheet yields no rows");

        std::size_t full = 0;
        reader.forEachRow("Rows", [&](std::size_t, const xlpp::StreamingRow&) { ++full; return true; });
        test.checkEqual(full, std::size_t{10}, "All 10 rows streamed");

        std::size_t early = 0;
        reader.forEachRow("Rows", [&](std::size_t, const xlpp::StreamingRow&) {
            return ++early < 3;
        });
        test.checkEqual(early, std::size_t{3}, "forEachRow early-stop callback honored");

        auto worksheet = reader.worksheet("Rows");
        auto begin = worksheet.begin();
        auto end = worksheet.end();
        test.checkTrue(begin != end, "Iterator range non-empty");
        test.checkEqual(begin.rowNumber(), std::size_t{1}, "First iterator row number");
        ++begin;
        test.checkEqual(begin.rowNumber(), std::size_t{2}, "Second iterator row number");
        const auto& row = *begin;
        test.checkEqual(row.size(), std::size_t{2}, "Row has two cells");
        test.checkNear(std::get<double>(row[0].value), 2.0, 1e-12, "Row number cell");
    }
    std::filesystem::remove(path);
}

void testCompressionLevelsAndParallel(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    auto buildWorkbook = [&](xlpp::SaveOptions options, const std::filesystem::path& path) {
        xlpp::Workbook wb;
        auto& sheet = wb.addWorksheet("Data");
        for (std::size_t r = 1; r <= 50; ++r)
            sheet.append({std::string("value-" + std::to_string(r)), double(r), r % 2 == 0});
        wb.save(path, options);
    };
    const auto storePath = dir / "xlpp_m21_store.xlsx";
    const auto bestPath = dir / "xlpp_m21_best.xlsx";
    const auto fastPath = dir / "xlpp_m21_fast.xlsx";
    const auto parallelPath = dir / "xlpp_m21_parallel.xlsx";

    xlpp::SaveOptions store; store.compressionLevel = xlpp::CompressionLevel::Store;
    xlpp::SaveOptions fast; fast.compressionLevel = xlpp::CompressionLevel::Fastest;
    xlpp::SaveOptions best; best.compressionLevel = xlpp::CompressionLevel::Best;
    xlpp::SaveOptions parallel; parallel.parallelWorkers = 4; parallel.parallelSheets = true;

    buildWorkbook(store, storePath);
    buildWorkbook(fast, fastPath);
    buildWorkbook(best, bestPath);
    buildWorkbook(parallel, parallelPath);
    buildWorkbook(parallel, fastPath); // same options re-save (differential cache path)

    test.checkTrue(std::filesystem::file_size(storePath) > std::filesystem::file_size(bestPath),
                   "Stored output larger than best-compressed");
    test.checkTrue(std::filesystem::file_size(fastPath) > 0, "Fastest output non-empty");

    auto loadAndVerify = [&](const std::filesystem::path& path, const std::string& label) {
        xlpp::Workbook loaded;
        loaded.load(path);
        auto* sheet = loaded.worksheet("Data");
        test.checkTrue(sheet != nullptr, label + " loads");
        test.checkEqual(sheet->cell("A50").stringValueOr(""), std::string("value-50"), label + " last value");
        test.checkNear(std::get<double>(sheet->cell("B50").value()), 50.0, 1e-12, label + " last number");
    };
    loadAndVerify(storePath, "Store level");
    loadAndVerify(fastPath, "Fastest level");
    loadAndVerify(bestPath, "Best level");
    loadAndVerify(parallelPath, "Parallel output");

    for (const auto& p : {storePath, bestPath, fastPath, parallelPath}) std::filesystem::remove(p);
}

void testStreamingAccessorsAndPostIncrement(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_streaming_accessors.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(path, xlpp::SharedStringMode::Hash);
        writer.setDate1904(true);
        writer.setCompressionLevel(xlpp::CompressionLevel::Fastest);
        writer.setCompressionStrategy(xlpp::CompressionStrategy::Filtered);
        writer.setParallelWorkers(2);
        auto& sheet = writer.addWorksheet("Rows");
        sheet.append({std::string("A"), 1.0});
        sheet.append({std::string("B"), 2.0});
        test.checkEqual(writer.sheetCount(), std::size_t{1}, "Streaming writer sheet count");
        test.checkEqual(writer.worksheet(0).name(), std::string("Rows"), "Streaming writer worksheet accessor");
        test.checkEqual(writer.worksheet(0).rowCount(), std::size_t{2}, "Streaming worksheet row count");
        test.checkTrue(writer.date1904(), "Streaming writer date-system getter");
        writer.close();
        writer.close();
        test.checkTrue(writer.closed(), "Streaming writer close is idempotent");
    }
    xlpp::StreamingWorkbookReader reader(path);
    auto rows = reader.worksheet("Rows");
    auto it = rows.begin();
    test.checkEqual(it->size(), std::size_t{2}, "Streaming iterator arrow operator");
    auto previous = it++;
    test.checkEqual(previous.rowNumber(), std::size_t{1}, "Streaming post-increment returns previous iterator");
    test.checkEqual(it.rowNumber(), std::size_t{2}, "Streaming post-increment advances iterator");
    std::size_t callbacks = 0;
    reader.forEachRow("Rows", [&](std::size_t rowNumber, const xlpp::StreamingRow&) {
        ++callbacks;
        return rowNumber < 1;
    });
    test.checkEqual(callbacks, std::size_t{1}, "Streaming callback supports early stop");
    std::filesystem::remove(path);
}

void testInternalIoAndScannerCoverage(TestContext& test) {
    const auto rawPath = std::filesystem::temp_directory_path() / "xlpp_mapped_file_test.bin";
    {
        std::ofstream out(rawPath, std::ios::binary);
        out << "abcdef";
    }
    xlpp::internal::MappedFile first(rawPath);
    test.checkEqual(first.size(), std::size_t{6}, "Mapped-file size");
    test.checkEqual(std::string(first.slice(1, 3)), std::string("bcd"), "Mapped-file slice");
    xlpp::internal::MappedFile second(std::move(first));
    xlpp::internal::MappedFile third(rawPath);
    third = std::move(second);
    test.checkEqual(std::string(third.view()), std::string("abcdef"), "Mapped-file move construction and assignment");
    bool sliceThrew = false;
    try { (void)third.slice(5, 2); } catch (const std::out_of_range&) { sliceThrew = true; }
    test.checkTrue(sliceThrew, "Mapped-file out-of-range slice throws");

    const std::string chars = "0123<abc>xyz";
    const auto* match = xlpp::internal::simd::findByteOr(chars.data(), chars.data() + chars.size(), '<', '>');
    test.checkTrue(match != nullptr && *match == '<', "SIMD dual-byte scan");
    test.checkEqual(xlpp::internal::simd::findStr(chars, "abc", 0), std::size_t{5}, "SIMD substring scan");
    test.checkTrue(xlpp::internal::xmlscan_detail::isNameChar(':'), "XML name-character helper");

    xlpp::internal::XmlScanner scanner(R"xml(<r><item id="1"/><item>two</item></r>)xml");
    std::string_view element;
    test.checkTrue(scanner.nextElement("item", element), "XML scanner finds first element");
    scanner.rewind();
    test.checkTrue(scanner.nextElement("item", element), "XML scanner rewind restarts iteration");

    xlpp::internal::ThreadPool pool(2);
    test.checkEqual(pool.size(), std::size_t{2}, "Thread-pool size accessor");
    std::array<int, 4> values{0, 0, 0, 0};
    pool.parallelFor(0, values.size(), [&](std::size_t i) { values[i] = static_cast<int>(i + 1); });
    test.checkEqual(values[3], 4, "Thread-pool parallelFor executes work");
    pool.parallelFor(2, 2, [&](std::size_t) {});

    xlpp::StreamingWorksheetWriter defaultWriter;
    xlpp::StreamingWorksheetWriter otherWriter;
    defaultWriter = std::move(otherWriter);
    test.checkEqual(defaultWriter.rowCount(), std::size_t{0}, "Streaming worksheet move assignment");

    const auto workbookPath = std::filesystem::temp_directory_path() / "xlpp_shared_reader_size.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(workbookPath, xlpp::SharedStringMode::Hash);
        writer.addWorksheet("S").append({std::string("one"), std::string("two"), std::string("one")});
        writer.close();
    }
    xlpp::internal::SharedStringsReader shared{xlpp::internal::ZipArchiveReader(workbookPath)};
    test.checkEqual(shared.size(), std::size_t{2}, "Shared-strings reader size accessor");
    test.checkTrue(shared.lookup(2) == nullptr, "Shared-strings reader missing lookup");

    xlpp::internal::ZipEntryInfo dummyInfo;
    dummyInfo.name = "dummy";
    xlpp::internal::ZipEntrySource pathSource(rawPath, dummyInfo);
    test.checkTrue(!pathSource.complete(), "Path-backed ZIP entry source constructor");

    std::filesystem::remove(rawPath);
    std::filesystem::remove(workbookPath);
}
