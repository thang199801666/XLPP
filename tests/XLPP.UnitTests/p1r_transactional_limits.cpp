#include "XML/XmlUtilities.h"
#include "Packaging/ZipArchive.h"
#include <XLPP/XLPP.h>

#include <chrono>
#include <filesystem>
#include <iostream>
#include <sstream>
#include <streambuf>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

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

std::size_t countTemporaryFiles(std::string_view prefix) {
    std::size_t count = 0;
    std::error_code ec;
    for (std::filesystem::directory_iterator it(std::filesystem::temp_directory_path(), ec), end;
         !ec && it != end; it.increment(ec)) {
        const auto name = it->path().filename().string();
        if (name.rfind(prefix, 0) == 0) ++count;
    }
    return count;
}

class RejectingStreamBuf final : public std::streambuf {
protected:
    std::streamsize xsputn(const char*, std::streamsize) override { return 0; }
    int_type overflow(int_type) override { return traits_type::eof(); }
};

void seedWorkbook(xlpp::Workbook& workbook, std::string sheetName = "Keep") {
    auto& sheet = workbook.addWorksheet(std::move(sheetName));
    sheet.cell("A1").setValue("sentinel");
}

std::filesystem::path writeSheetPackage(std::string sheetXml,
                                            std::string sharedStringsXml = {},
                                            std::string stylesXml = {}) {
    xlpp::internal::ZipArchive zip;
    zip.add("xl/workbook.xml",
            R"xml(<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>)xml", false);
    zip.add("xl/_rels/workbook.xml.rels",
            R"xml(<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>)xml", false);
    zip.add("xl/worksheets/sheet1.xml", std::move(sheetXml), false);
    if (!sharedStringsXml.empty()) zip.add("xl/sharedStrings.xml", std::move(sharedStringsXml), false);
    if (!stylesXml.empty()) zip.add("xl/styles.xml", std::move(stylesXml), false);
    const auto path = tempPath("xlpp_p1r_sheet");
    zip.save(path);
    return path;
}

std::filesystem::path writeRelationshipPackage(std::string extraRelationship,
                                               std::string primaryType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") {
    xlpp::internal::ZipArchive zip;
    zip.add("xl/workbook.xml",
            "<?xml version=\"1.0\"?><workbook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Sheet1\" r:id=\"rId1\"/></sheets></workbook>", false);
    zip.add("xl/_rels/workbook.xml.rels",
            "<?xml version=\"1.0\"?><Relationships><Relationship Id=\"rId1\" Type=\"" + primaryType +
                "\" Target=\"worksheets/sheet1.xml\"/>" + extraRelationship + "</Relationships>", false);
    zip.add("xl/worksheets/sheet1.xml",
            "<?xml version=\"1.0\"?><worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>7</v></c></row></sheetData></worksheet>", false);
    const auto path = tempPath("xlpp_p1r_relationship");
    zip.save(path);
    return path;
}

std::filesystem::path writeOwnedRelationshipPackage(std::string sheetXml,
                                                    std::string sheetRelationshipsXml,
                                                    std::string contentTypesXml = {},
                                                    std::vector<std::pair<std::string, std::string>> extraParts = {}) {
    xlpp::internal::ZipArchive zip;
    zip.add("xl/workbook.xml",
            R"xml(<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>)xml", false);
    zip.add("xl/_rels/workbook.xml.rels",
            R"xml(<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>)xml", false);
    zip.add("xl/worksheets/sheet1.xml", std::move(sheetXml), false);
    if (!sheetRelationshipsXml.empty())
        zip.add("xl/worksheets/_rels/sheet1.xml.rels", std::move(sheetRelationshipsXml), false);
    if (!contentTypesXml.empty()) zip.add("[Content_Types].xml", std::move(contentTypesXml), false);
    for (auto& [name, data] : extraParts) zip.add(std::move(name), std::move(data), false);
    const auto path = tempPath("xlpp_p1r_owned_relationship");
    zip.save(path);
    return path;
}


void testMaterializedXmlQuotedDelimiter() {
    const std::string xml = R"(<root><node note="1 > 0">payload</node><node note='x > y'/></root>)";
    const auto nodes = xlpp::internal::tags(xml, "node");
    check(nodes.size() == 2, "materialized tag scanner respects quoted > delimiters");
    check(nodes.size() >= 1 && xlpp::internal::attribute(nodes[0], "note") == "1 > 0",
          "materialized attribute remains intact after quoted > scan");
    check(xlpp::internal::tagTextView(xml, "node") == "payload",
          "zero-copy tagTextView respects quoted > delimiters");
    std::size_t visited = 0;
    xlpp::internal::tagsForEach(xml, "node", [&](std::string_view) { ++visited; });
    check(visited == 2, "zero-copy tagsForEach respects quoted > delimiters");

    const std::string lexical = R"xml(<root><!-- <node>fake</node> --><node>outer<node>inner</node><!-- </node> --><![CDATA[</node>]]>tail</node><?pi value="<node>"?><![CDATA[<node>fake</node>]]></root>)xml";
    const auto lexicalNodes = xlpp::internal::tags(lexical, "node");
    check(lexicalNodes.size() == 1, "materialized scanner ignores comment/CDATA/PI pseudo-elements and keeps nested same-name element balanced");
    check(lexicalNodes.size() == 1 && lexicalNodes.front().find("tail</node>") != std::string::npos,
          "fake closing tags inside comment/CDATA do not truncate the owning element");
    std::size_t lexicalVisited = 0;
    xlpp::internal::tagsForEach(lexical, "node", [&](std::string_view) { ++lexicalVisited; });
    check(lexicalVisited == 1, "zero-copy scanner shares lexical markup hardening");
    check(xlpp::internal::tags("<root><!-- <node>fake</node>", "node").empty(),
          "unterminated XML comment is not scanned as element payload");
}

void testTransactionalPathLoadAndRelationshipHardening() {
    {
        const auto path = writeRelationshipPackage(
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>");
        xlpp::Workbook workbook;
        seedWorkbook(workbook);
        checkThrows([&] { workbook.load(path); }, "materialized load rejects duplicate workbook relationship Id");
        check(workbook.sheetCount() == 1 && workbook.worksheet("Keep") != nullptr,
              "failed path load preserves the prior workbook state");
        std::filesystem::remove(path);
    }
    {
        const auto path = writeRelationshipPackage({},
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
        xlpp::Workbook workbook;
        seedWorkbook(workbook);
        checkThrows([&] { workbook.load(path); }, "sheet r:id cannot bind a non-worksheet relationship");
        check(workbook.sheetCount() == 1 && workbook.worksheet("Keep") != nullptr,
              "relationship-type failure is transactional");
        std::filesystem::remove(path);
    }
}


void testOwnedRelationshipAndContentTypeHardening() {
    const auto rejectTransactional = [](const std::filesystem::path& path, const char* message) {
        xlpp::Workbook workbook;
        seedWorkbook(workbook);
        checkThrows([&] { workbook.load(path); }, message);
        check(workbook.sheetCount() == 1 && workbook.worksheet("Keep") != nullptr,
              "owned relationship/content-type failure remains transactional");
        std::filesystem::remove(path);
    };

    {
        const auto path = writeOwnedRelationshipPackage(
            R"xml(<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><hyperlinks><hyperlink ref="A1" r:id="rId2"/></hyperlinks><sheetData/></worksheet>)xml",
            R"xml(<Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org" TargetMode="External"/></Relationships>)xml");
        rejectTransactional(path, "duplicate worksheet relationship Id is rejected");
    }
    {
        const auto path = writeOwnedRelationshipPackage(
            R"xml(<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><tableParts count="1"><tablePart r:id="rId2"/></tableParts><sheetData/></worksheet>)xml",
            R"xml(<Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>)xml");
        rejectTransactional(path, "tablePart cannot bind a hyperlink relationship");
    }
    {
        const auto path = writeOwnedRelationshipPackage(
            R"xml(<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><tableParts count="1"><tablePart r:id="rId2"/></tableParts><sheetData/></worksheet>)xml",
            R"xml(<Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../../../escape.xml"/></Relationships>)xml");
        rejectTransactional(path, "worksheet relationship cannot escape the package root");
    }
    {
        const auto path = writeOwnedRelationshipPackage(
            R"xml(<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><tableParts count="1"><tablePart r:id="rId2"/></tableParts><sheetData/></worksheet>)xml",
            R"xml(<Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="file:tables/table1.xml"/></Relationships>)xml");
        rejectTransactional(path, "internal relationship URI schemes are rejected");
    }
    {
        const auto path = writeOwnedRelationshipPackage(
            R"xml(<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>)xml", {},
            R"xml(<Types><Override PartName="/xl/workbook.xml" ContentType="application/a"/><Override PartName="/xl/workbook.xml" ContentType="application/b"/></Types>)xml");
        rejectTransactional(path, "duplicate content-type Override PartName is rejected");
    }
    {
        const auto path = writeOwnedRelationshipPackage(
            R"xml(<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><hyperlinks><hyperlink ref="A1" r:id="rId2"/></hyperlinks><sheetData/></worksheet>)xml",
            R"xml(<Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a?b=1" TargetMode="External"/></Relationships>)xml");
        xlpp::Workbook workbook;
        workbook.load(path);
        const auto* sheet = workbook.worksheet("Sheet1");
        const auto* cell = sheet ? sheet->tryCell("A1") : nullptr;
        check(cell && cell->hasHyperlink() && cell->hyperlinkValue()->target() == "https://example.com/a?b=1",
              "valid External hyperlink relationship remains supported");
        std::filesystem::remove(path);
    }
}

void testStrictMaterializedNumericParsing() {
    const auto checkSheetRejected = [&](std::string sheetXml, std::string shared = {}, std::string styles = {}, const char* message = "malformed numeric XML is rejected") {
        const auto path = writeSheetPackage(std::move(sheetXml), std::move(shared), std::move(styles));
        xlpp::Workbook workbook;
        seedWorkbook(workbook);
        checkThrows([&] { workbook.load(path); }, message);
        check(workbook.sheetCount() == 1 && workbook.worksheet("Keep") != nullptr,
              "strict numeric parse failure remains transactional");
        std::filesystem::remove(path);
    };

    checkSheetRejected(
        R"xml(<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1"><v>12junk</v></c></row></sheetData></worksheet>)xml",
        {}, {}, "cell numeric value with trailing garbage is rejected");
    checkSheetRejected(
        R"xml(<?xml version="1.0"?><worksheet><sheetData><row r="1junk"><c r="A1"><v>1</v></c></row></sheetData></worksheet>)xml",
        {}, {}, "row index with trailing garbage is rejected");
    checkSheetRejected(
        R"xml(<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>4</v></c></row></sheetData></worksheet>)xml",
        R"xml(<?xml version="1.0"?><sst><si><t>only</t></si></sst>)xml", {},
        "out-of-range shared string index is rejected");
    checkSheetRejected(
        R"xml(<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="2"><v>1</v></c></row></sheetData></worksheet>)xml",
        {}, R"xml(<?xml version="1.0"?><styleSheet><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>)xml",
        "out-of-range cell style index is rejected");
    checkSheetRejected(
        R"xml(<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1"><v>nan</v></c></row></sheetData></worksheet>)xml",
        {}, {}, "non-finite cell numeric value is rejected");
}

std::string serializedLimitFixture() {
    xlpp::Workbook source;
    auto& a = source.addWorksheet("Data");
    a.cell("A1").setValue("alpha");
    a.cell("A2").setValue("beta");
    auto& b = source.addWorksheet("More");
    b.cell("B2").setValue("gamma");
    source.addDefinedName(xlpp::DefinedName("NameOne", "Data!$A$1"));
    source.addDefinedName(xlpp::DefinedName("NameTwo", "Data!$A$2"));
    std::ostringstream output(std::ios::binary);
    source.save(output);
    return output.str();
}

void checkStreamLimit(const std::string& payload, xlpp::LoadOptions options, const char* message) {
    xlpp::Workbook target;
    seedWorkbook(target);
    std::istringstream input(payload, std::ios::binary);
    checkThrows([&] { target.load(input, options); }, message);
    check(target.sheetCount() == 1 && target.worksheet("Keep") != nullptr,
          "model-limit failure preserves existing workbook");
}

void testModelMaterializationLimits() {
    const auto payload = serializedLimitFixture();
    check(!payload.empty(), "limit fixture serialized successfully");

    {
        const auto before = countTemporaryFiles("xlpp_stream_load_");
        xlpp::LoadOptions options;
        options.maxFileBytes = payload.size() - 1;
        checkStreamLimit(payload, options, "istream maxFileBytes is enforced before package materialization");
        const auto after = countTemporaryFiles("xlpp_stream_load_");
        check(before == after, "failed stream load removes its temporary input file");
    }
    {
        const auto before = countTemporaryFiles("xlpp_stream_load_");
        xlpp::LoadOptions options;
        options.cancel = [] { return true; };
        checkStreamLimit(payload, options, "istream load cancellation is honored before materialization");
        const auto after = countTemporaryFiles("xlpp_stream_load_");
        check(before == after, "cancelled stream load removes its temporary input file");
    }
    {
        xlpp::LoadOptions options;
        options.maxWorksheets = 1;
        checkStreamLimit(payload, options, "maxWorksheets rejects oversized workbook model");
    }
    {
        xlpp::LoadOptions options;
        options.maxCells = 2;
        checkStreamLimit(payload, options, "maxCells rejects excessive materialized cells");
    }
    {
        xlpp::LoadOptions options;
        options.maxSharedStrings = 1;
        checkStreamLimit(payload, options, "maxSharedStrings rejects excessive SST cardinality");
    }
    {
        xlpp::LoadOptions options;
        options.maxDefinedNames = 1;
        checkStreamLimit(payload, options, "maxDefinedNames rejects excessive defined-name count");
    }
}

void testOutputStreamFailureAndTemporaryCleanup() {
    xlpp::Workbook workbook;
    seedWorkbook(workbook, "Output");

    const auto before = countTemporaryFiles("xlpp_stream_save_");
    RejectingStreamBuf buffer;
    std::ostream output(&buffer);
    checkThrows([&] { workbook.save(output); }, "save(ostream) reports destination write failures");
    const auto after = countTemporaryFiles("xlpp_stream_save_");
    check(before == after, "save(ostream) removes temporary file after destination failure");
}

} // namespace

int main() {
    testMaterializedXmlQuotedDelimiter();
    testTransactionalPathLoadAndRelationshipHardening();
    testOwnedRelationshipAndContentTypeHardening();
    testStrictMaterializedNumericParsing();
    testModelMaterializationLimits();
    testOutputStreamFailureAndTemporaryCleanup();

    if (failures) {
        std::cerr << failures << " P1R transactional/limit check(s) failed\n";
        return 1;
    }
    std::cout << "P1R transactional and model-limit checks PASS\n";
    return 0;
}
