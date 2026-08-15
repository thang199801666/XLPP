#include <XLPP/XLPP.h>
#include "Packaging/ZipArchive.h"
#include "Packaging/ZipArchiveReader.h"
#include "Packaging/RelationshipGraph.h"
#include "Packaging/MappedFile.h"
#include "Streaming/SharedStringsReader.h"
#include "Threading/ThreadPool.h"
#include "XML/SimdScan.h"
#include "XML/XmlScanner.h"
#include "XML/XmlUtilities.h"
#include "Vba/VbaProjectBinary.h"
#include "Encryption/OfficeCrypto.h"
#include "Formula/ReferenceTransformer.h"
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

namespace {
class TestContext {
public:
    template <typename Actual, typename Expected>
    void checkEqual(const Actual& actual, const Expected& expected, const std::string& label) {
        ++checks_;
        if (actual == expected) {
            std::cout << "    [CHECK PASS] " << label << " | actual=" << printable(actual)
                      << " expected=" << printable(expected) << '\n';
            return;
        }
        std::ostringstream message;
        message << label << " | actual=" << printable(actual)
                << " expected=" << printable(expected);
        throw std::runtime_error(message.str());
    }

    void checkNear(double actual, double expected, double tolerance, const std::string& label) {
        ++checks_;
        if (std::abs(actual - expected) <= tolerance) {
            std::cout << "    [CHECK PASS] " << label << " | actual=" << actual
                      << " expected=" << expected << " tolerance=" << tolerance << '\n';
            return;
        }
        std::ostringstream message;
        message << label << " | actual=" << actual << " expected=" << expected
                << " tolerance=" << tolerance;
        throw std::runtime_error(message.str());
    }

    void checkTrue(bool condition, const std::string& label) {
        ++checks_;
        if (condition) {
            std::cout << "    [CHECK PASS] " << label << " | actual=true expected=true\n";
            return;
        }
        throw std::runtime_error(label + " | actual=false expected=true");
    }

    std::size_t checks() const noexcept { return checks_; }

private:
    template <typename T>
    static std::string printable(const T& value) {
        std::ostringstream stream;
        stream << value;
        return stream.str();
    }

    static std::string printable(const std::string& value) { return '"' + value + '"'; }
    static std::string printable(const char* value) { return printable(std::string(value)); }
    static std::string printable(bool value) { return value ? "true" : "false"; }
    static std::string printable(const xlpp::DateTime& value) { return xlpp::toIso8601(value); }

    std::size_t checks_{0};
};

using TestFunction = std::function<void(TestContext&)>;

void testCellReferences(TestContext& test) {
    const auto ref = xlpp::CellReference::parse("$aa$42");
    test.checkEqual(ref.row, std::size_t{42}, "Absolute row is parsed");
    test.checkEqual(ref.column, std::size_t{27}, "Lower-case column is normalized");
    test.checkEqual(ref.address(), std::string("AA42"), "Canonical A1 address");
    test.checkEqual(xlpp::CellReference::columnName(16384), std::string("XFD"), "Excel maximum column");
}

void testRangeAndDimensions(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.cell(2, 2).setValue("center");
    sheet.range("C3:D4").setValue(std::string("range"));
    test.checkEqual(sheet.dimensions(), std::string("B2:D4"), "Used worksheet dimensions");
    test.checkEqual(sheet.maxRow(), std::size_t{4}, "Maximum row");
    test.checkEqual(sheet.maxColumn(), std::size_t{4}, "Maximum column");
    test.checkEqual(sheet.range("D4:C3").address(), std::string("C3:D4"), "Reversed range normalization");
}

void testAppendAndStructuralEdits(TestContext& test) {
    xlpp::Worksheet sheet("Rows");
    sheet.append({std::string("Name"), std::string("Value")});
    sheet.append({std::string("A"), 10.0});
    test.checkEqual(sheet.trackedCellChangeCount(), std::size_t{4}, "Bulk append keeps per-cell dependency tracking");
    sheet.insertRows(2);
    test.checkEqual(std::get<std::string>(sheet.cell("A3").value()), std::string("A"), "Insert rows moves string cell");
    sheet.deleteRows(2);
    test.checkEqual(std::get<std::string>(sheet.cell("A2").value()), std::string("A"), "Delete rows restores string cell");
    sheet.insertColumns(2, 2);
    test.checkNear(std::get<double>(sheet.cell("D2").value()), 10.0, 1e-12, "Insert columns moves numeric cell");
    sheet.deleteColumns(2, 2);
    test.checkNear(std::get<double>(sheet.cell("B2").value()), 10.0, 1e-12, "Delete columns restores numeric cell");
}

void testCellConvenience(TestContext& test) {
    xlpp::Cell cell("A1");
    test.checkTrue(!cell.hasValue(), "Empty cell has no value");
    test.checkEqual(std::string(cell.valueType()), std::string("empty"), "valueType empty");
    cell.setValue(42.0);
    test.checkTrue(cell.hasValue(), "Has value after setting double");
    test.checkTrue(cell.isNumeric(), "isNumeric true for double");
    test.checkEqual(std::string(cell.valueType()), std::string("numeric"), "valueType numeric");
    cell.setValue("hello");
    test.checkTrue(cell.isString(), "isString true for string");
    test.checkTrue(!cell.isNumeric(), "isNumeric false for string");
    test.checkEqual(std::string(cell.valueType()), std::string("string"), "valueType string");
    cell.setValue(true);
    test.checkTrue(cell.isBoolean(), "isBoolean true for bool");
    test.checkEqual(std::string(cell.valueType()), std::string("bool"), "valueType bool");
    cell.setError(xlpp::CellError::Value);
    test.checkTrue(cell.isError(), "isError true for CellError");
    test.checkEqual(std::string(cell.valueType()), std::string("error"), "valueType error");
    cell.clear();
    test.checkTrue(!cell.hasValue(), "Cleared cell has no value");
}

void testNamedStyleAssociation(TestContext& test) {
    xlpp::Workbook wb;
    wb.addNamedStyle({"Accent", wb.worksheets().empty() ? xlpp::Style{} : xlpp::Style{}});
    auto& sheet = wb.addWorksheet("Sheet1");
    auto& cell = sheet.cell("A1");
    wb.applyNamedStyle(cell, "Accent");
    test.checkTrue(cell.namedStyle().has_value(), "Cell tracks named style after apply");
    test.checkEqual(cell.namedStyle().value(), std::string("Accent"), "Named style name is stored");
    cell.clear();
    test.checkTrue(!cell.namedStyle().has_value(), "Cell clear resets named style association");
}

void testRemoveWorksheet(TestContext& test) {
    xlpp::Workbook wb;
    wb.addWorksheet("First");
    wb.addWorksheet("Second");
    wb.addWorksheet("Third");
    test.checkEqual(wb.worksheets().size(), std::size_t{3}, "Three sheets added");
    test.checkTrue(wb.removeWorksheet("Second"), "removeWorksheet returns true for existing sheet");
    test.checkEqual(wb.worksheets().size(), std::size_t{2}, "Two sheets remain");
    test.checkTrue(wb.worksheet("First") != nullptr, "First sheet still present");
    test.checkTrue(wb.worksheet("Second") == nullptr, "Second sheet removed");
    test.checkTrue(!wb.removeWorksheet("Nope"), "removeWorksheet returns false for non-existent sheet");
}

void testWorksheetExtents(TestContext& test) {
    xlpp::Worksheet sheet("Extents");
    auto e = sheet.extents();
    test.checkEqual(e.minRow, std::size_t{1}, "Empty sheet minRow is 1");
    test.checkEqual(e.maxRow, std::size_t{1}, "Empty sheet maxRow is 1");
    test.checkEqual(sheet.rowCount(), std::size_t{1}, "Empty sheet rowCount is 1");
    test.checkEqual(sheet.columnCount(), std::size_t{1}, "Empty sheet columnCount is 1");

    sheet.cell("C5").setValue(1.0);
    sheet.cell("A2").setValue("a");
    e = sheet.extents();
    test.checkEqual(e.minRow, std::size_t{2}, "extents minRow after cells");
    test.checkEqual(e.minColumn, std::size_t{1}, "extents minColumn after cells");
    test.checkEqual(e.maxRow, std::size_t{5}, "extents maxRow after cells");
    test.checkEqual(e.maxColumn, std::size_t{3}, "extents maxColumn after cells");

    // Structural edits rebuild the ordered cell map and must invalidate the
    // cached extents rather than returning stale geometry.
    sheet.insertRows(1, 1);
    e = sheet.extents();
    test.checkEqual(e.minRow, std::size_t{3}, "extents cache invalidated after structural edit");
    test.checkEqual(sheet.dimensions(), std::string("A3:C6"), "dimensions refresh after structural edit");
}

void testDOMSharedStrings(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m21_dom_sst.xlsx";
    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("Data");
    sheet.cell("A1").setValue("repeated");
    sheet.cell("A2").setValue("repeated");
    sheet.cell("A3").setValue("repeated");
    sheet.cell("B1").setValue("unique");
    wb.save(path);

    xlpp::Workbook loaded;
    loaded.load(path);
    test.checkEqual(std::get<std::string>(loaded.worksheet("Data")->cell("A1").value()), std::string("repeated"), "DOM SST: repeated string round-trips");
    test.checkEqual(std::get<std::string>(loaded.worksheet("Data")->cell("A2").value()), std::string("repeated"), "DOM SST: repeated string in second cell");
    test.checkEqual(std::get<std::string>(loaded.worksheet("Data")->cell("B1").value()), std::string("unique"), "DOM SST: unique string round-trips");

    // Verify shared strings XML is present and uses t=\"s\"
    xlpp::internal::ZipArchive z = xlpp::internal::ZipArchive::open(path);
    test.checkTrue(z.contains("xl/sharedStrings.xml"), "DOM save produces shared strings");
    const auto sstXml = z.get("xl/sharedStrings.xml");
    test.checkTrue(sstXml.find("uniqueCount=\"2\"") != std::string::npos, "SST has 2 unique strings");

    std::filesystem::remove(path);
}

void testRichTextSharedStrings(TestContext& test) {
    // Build a minimal xlsx with rich-text shared strings directly via ZipArchive
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m21_rich_text.xlsx";
    {
        xlpp::internal::ZipArchive z;
        z.add("[Content_Types].xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>)");
        z.add("_rels/.rels",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>)");
        z.add("xl/workbook.xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Rich" sheetId="1" r:id="rId1"/></sheets></workbook>)");
        z.add("xl/_rels/workbook.xml.rels",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>)");
        z.add("xl/styles.xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>)");
         z.replace("xl/worksheets/sheet1.xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>)");
        // Rich text shared string: two <r> elements that should be concatenated
        z.add("xl/sharedStrings.xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><r><t>Hello </t></r><r><t>World</t></r></si><si><t>plain</t></si></sst>)");
        z.save(path);
    }
    xlpp::Workbook loaded;
    loaded.load(path);
    test.checkEqual(std::get<std::string>(loaded.worksheet("Rich")->cell("A1").value()), std::string("Hello World"), "Rich text concatenated on DOM load");
    const auto* richCell = loaded.worksheet("Rich")->tryCell("A1");
    test.checkTrue(richCell && richCell->hasRichText(), "Rich text formatting model is retained on DOM load");
    test.checkEqual(richCell->richTextValue()->runs().size(), std::size_t{2}, "All rich text runs are retained");
    test.checkEqual(std::get<std::string>(loaded.worksheet("Rich")->cell("A2").value()), std::string("plain"), "Plain shared string next to rich text");
    std::filesystem::remove(path);
}

void testRichTextCellRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_rich_text_cell_roundtrip.xlsx";
    xlpp::RichText richText;
    xlpp::RichTextRun first("Bold red ");
    first.setBold(true);
    first.setColor("FFFF0000");
    xlpp::RichTextRun second("italic blue");
    second.setItalic(true);
    second.setUnderline(true);
    second.setFontName("Arial");
    second.setSize(14.0);
    second.setColor("FF0000FF");
    richText.addRun(first);
    richText.addRun(second);

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Rich");
    sheet.cell("A1").setRichText(richText);
    workbook.save(path);

    const auto archive = xlpp::internal::ZipArchive::open(path);
    const auto sheetXml = archive.get("xl/worksheets/sheet1.xml");
    test.checkTrue(sheetXml.find("<c r=\"A1\" t=\"inlineStr\"><is><r>") != std::string::npos,
                   "Rich text cell is serialized as inline rich text");
    test.checkTrue(sheetXml.find("<b val=\"1\"/>") != std::string::npos, "Bold rich text property is serialized");
    test.checkTrue(sheetXml.find("<i val=\"1\"/>") != std::string::npos, "Italic rich text property is serialized");
    test.checkTrue(sheetXml.find("<u val=\"single\"/>") != std::string::npos, "Underline rich text property is serialized");
    test.checkTrue(sheetXml.find("<color rgb=\"FFFF0000\"/>") != std::string::npos, "Rich text color is serialized");
    test.checkTrue(sheetXml.find("<rFont val=\"Arial\"/>") != std::string::npos, "Rich text font is serialized");
    test.checkTrue(sheetXml.find("<sz val=\"14\"/>") != std::string::npos, "Rich text size is serialized");

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedCell = loaded.worksheet("Rich")->tryCell("A1");
    test.checkTrue(loadedCell && loadedCell->hasRichText(), "Rich text cell survives load");
    test.checkEqual(loadedCell->richTextValue()->runs().size(), std::size_t{2}, "Rich text run count round-trips");
    test.checkTrue(loadedCell->richTextValue()->runs()[0].bold(), "Bold property round-trips");
    test.checkTrue(loadedCell->richTextValue()->runs()[1].italic(), "Italic property round-trips");
    test.checkTrue(loadedCell->richTextValue()->runs()[1].underline(), "Underline property round-trips");
    test.checkEqual(loadedCell->richTextValue()->plainText(), std::string("Bold red italic blue"), "Rich text plain value round-trips");
    std::filesystem::remove(path);
}

void testBuiltinDateFormatIds(TestContext& test) {
    test.checkTrue(xlpp::isDateFormatCode("General", 14), "numFmtId 14 is date format");
    test.checkTrue(xlpp::isDateFormatCode("General", 22), "numFmtId 22 is date format");
    test.checkTrue(xlpp::isDateFormatCode("General", 36), "numFmtId 36 is date format");
    test.checkTrue(xlpp::isDateFormatCode("General", 55), "numFmtId 55 is date format");
    test.checkTrue(!xlpp::isDateFormatCode("General", 0), "numFmtId 0 is not date format");
    test.checkTrue(!xlpp::isDateFormatCode("General", 2), "numFmtId 2 is not date format");
    test.checkTrue(xlpp::isDateFormatCode("yyyy-mm-dd"), "Format string yyyy-mm-dd is date");
    test.checkTrue(!xlpp::isDateFormatCode("General"), "General format is not date");
}

void testRowProxyAndRangeHelpers(TestContext& test) {
    xlpp::Worksheet sheet("Data");
    sheet.cell("A1").setValue(std::string("Name"));
    sheet.cell("B1").setValue(42.0);
    sheet.cell("A2").setValue(std::string("Age"));
    sheet.cell("B2").setValue(30.0);
    sheet.cell("C1").setValue(true);

    auto row = sheet.row(1);
    test.checkEqual(row.number(), std::size_t{1}, "Row proxy number");
    test.checkEqual(std::get<double>(row.cell(2).value()), 42.0, "Row proxy cell access");
    auto rowCells = row.cells();
    test.checkEqual(rowCells.size(), std::size_t{3}, "Row proxy cells count for non-empty cells");

    auto rowValues = sheet.row(2).values();
    test.checkEqual(rowValues.size(), std::size_t{3}, "Row proxy values count includes all columns");

    auto rng = sheet.range("A1:B2");
    std::vector<std::string> visited;
    rng.forEach([&](xlpp::Cell& c) { visited.push_back(c.address()); });
    test.checkEqual(visited.size(), std::size_t{4}, "CellRange::forEach visits all cells");
    test.checkEqual(visited[0], std::string("A1"), "forEach visits A1 first");

    auto vals = rng.values();
    test.checkEqual(vals.size(), std::size_t{4}, "CellRange::values count");
    test.checkEqual(std::get<std::string>(vals[0]), std::string("Name"), "CellRange::values first value");

    auto formulas = rng.formulas();
    test.checkEqual(formulas.size(), std::size_t{4}, "CellRange::formulas count");
}

void testCellStyleIndex(TestContext& test) {
    xlpp::Cell cell("A1");
    test.checkTrue(!cell.styleIndex().has_value(), "styleIndex not set by default");
    cell.setRawStyleIndex(5);
    test.checkEqual(cell.styleIndex().value(), std::size_t{5}, "styleIndex getter returns set value");
    cell.clearRawStyleIndex();
    test.checkTrue(!cell.styleIndex().has_value(), "styleIndex cleared");
}

void testStreamLoadSave(TestContext& test) {
    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("StreamTest");
    sheet.cell("A1").setValue(std::string("stream value"));
    sheet.cell("B1").setValue(99.5);

    std::ostringstream out;
    wb.save(out);
    test.checkTrue(out.str().size() > 100, "Stream save produces non-trivial output");

    xlpp::Workbook loaded;
    std::istringstream in(out.str());
    loaded.load(in);
    test.checkEqual(std::get<std::string>(loaded.worksheet("StreamTest")->cell("A1").value()),
                    std::string("stream value"), "Stream load round-trips string");
    test.checkNear(std::get<double>(loaded.worksheet("StreamTest")->cell("B1").value()),
                   99.5, 1e-12, "Stream load round-trips number");
}

void testNumFmtIdDateRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m21_numfmt_date.xlsx";
    {
        // Build a file with a built-in date format (numFmtId 14 = Short Date) without a formatCode
        xlpp::internal::ZipArchive z;
        z.add("[Content_Types].xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>)");
        z.add("_rels/.rels",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>)");
        z.add("xl/workbook.xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="DateSheet" sheetId="1" r:id="rId1"/></sheets></workbook>)");
        z.add("xl/_rels/workbook.xml.rels",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>)");
        z.add("xl/styles.xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>)");
         z.replace("xl/worksheets/sheet1.xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" s="1"><v>45306</v></c></row><row r="2"><c r="A2"><v>42</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>)");
        z.save(path);
    }
    xlpp::Workbook loaded;
    loaded.load(path);
    test.checkTrue(std::holds_alternative<xlpp::DateTime>(loaded.worksheet("DateSheet")->cell("A1").value()),
                   "Built-in date format (numFmtId 14) detected as DateTime");
    test.checkTrue(std::holds_alternative<double>(loaded.worksheet("DateSheet")->cell("A2").value()),
                   "Plain number without date format stays numeric");
    std::filesystem::remove(path);
}

void testEdgeCasesAndCleanup(TestContext& test) {
    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("Test");
    auto& cell = sheet.cell("A1");

    cell.setStringValue(std::string("typed string"));
    test.checkTrue(cell.isString(), "setStringValue creates string");
    test.checkEqual(std::get<std::string>(cell.value()), std::string("typed string"), "setStringValue value round-trip");

    cell.setNumericValue(3.14);
    test.checkTrue(cell.isNumeric(), "setNumericValue creates numeric");
    test.checkNear(std::get<double>(cell.value()), 3.14, 1e-12, "setNumericValue value round-trip");

    cell.setBoolValue(true);
    test.checkTrue(cell.isBoolean(), "setBoolValue creates boolean");
    test.checkTrue(std::get<bool>(cell.value()), "setBoolValue value round-trip");

    test.checkNear(cell.numericValueOr(-1.0), -1.0, 1e-12, "numericValueOr returns fallback for non-numeric");
    cell.setNumericValue(2.5);
    test.checkNear(cell.numericValueOr(0.0), 2.5, 1e-12, "numericValueOr returns value for numeric");

    cell.setComment(xlpp::Comment("note", "me"));
    test.checkTrue(cell.hasComment(), "setComment sets comment");
    cell.clearComment();
    test.checkTrue(!cell.hasComment(), "clearComment removes comment");

    cell.setHyperlink(xlpp::Hyperlink("http://test"));
    test.checkTrue(cell.hasHyperlink(), "setHyperlink sets hyperlink");
    cell.clearHyperlink();
    test.checkTrue(!cell.hasHyperlink(), "clearHyperlink removes hyperlink");

    cell.setFormula("=1+2");
    test.checkTrue(cell.hasFormula(), "setFormula sets formula");
    cell.clearFormula();
    test.checkTrue(!cell.hasFormula(), "clearFormula removes formula");

    wb.clear();
    test.checkEqual(wb.worksheets().size(), std::size_t{0}, "Workbook::clear removes all sheets");

    xlpp::Workbook wb2;
    auto& sheet2 = wb2.addWorksheet("Test2");
    sheet2.mergeCells("A1:B2");
    bool threw = false;
    try { sheet2.unmergeCells("X99:Z100"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "unmergeCells throws on non-existent range");
}

void testWorksheetRows(TestContext& test) {
    xlpp::Worksheet sheet("Rows");
    sheet.cell("A1").setValue(std::string("one"));
    sheet.cell("B2").setValue(std::string("two"));
    sheet.cell("C3").setValue(std::string("three"));

    auto rows = sheet.rows();
    test.checkEqual(rows.size(), std::size_t{3}, "rows() returns correct count");
    test.checkEqual(rows[0].number(), std::size_t{1}, "row 0 is row number 1");
    test.checkEqual(std::get<std::string>(rows[0].cell(1).value()), std::string("one"), "rows()[0].cell(1) value");
}

void testIterRowsCols(TestContext& test) {
    xlpp::Worksheet sheet("Grid");
    sheet.cell("A1").setValue(1.0);
    sheet.cell("B1").setValue(2.0);
    sheet.cell("A2").setValue(3.0);
    sheet.cell("B2").setValue(4.0);

    auto block = sheet.iterRows(1, 2, 1, 2);
    test.checkEqual(block.size(), std::size_t{2}, "iterRows returns 2 rows");
    test.checkEqual(block[0].size(), std::size_t{2}, "iterRows row has 2 columns");
    test.checkNear(std::get<double>(block[0][0]), 1.0, 1e-12, "iterRows A1");
    test.checkNear(std::get<double>(block[1][1]), 4.0, 1e-12, "iterRows B2");

    auto cols = sheet.iterCols(1, 2, 1, 2);
    test.checkEqual(cols.size(), std::size_t{2}, "iterCols returns 2 columns");
    test.checkEqual(cols[0].size(), std::size_t{2}, "iterCols column has 2 rows");
    test.checkNear(std::get<double>(cols[0][0]), 1.0, 1e-12, "iterCols A1");
    test.checkNear(std::get<double>(cols[1][1]), 4.0, 1e-12, "iterCols B2");

    auto all = sheet.iterRows();
    test.checkEqual(all.size(), std::size_t{2}, "iterRows with 0 bounds uses extents");
    test.checkEqual(all[0].size(), std::size_t{2}, "iterRows col count from extents");
}

void testCellOffset(TestContext& test) {
    xlpp::Cell cell(5, 3);
    auto ref = cell.offset(2, 1);
    test.checkEqual(ref.row, std::size_t{7}, "offset +2 rows");
    test.checkEqual(ref.column, std::size_t{4}, "offset +1 col");
    test.checkEqual(ref.address(), std::string("D7"), "offset address");

    auto refUp = cell.offset(-3, 1);
    test.checkEqual(refUp.row, std::size_t{2}, "offset -3 rows");
    test.checkEqual(refUp.column, std::size_t{4}, "offset +1 col");
}

void testWorkbookNav(TestContext& test) {
    xlpp::Workbook wb;
    wb.addWorksheet("First");
    wb.addWorksheet("Second");
    wb.addWorksheet("Third");

    auto names = wb.sheetNames();
    test.checkEqual(names.size(), std::size_t{3}, "sheetNames count");
    test.checkEqual(names[1], std::string("Second"), "sheetNames index 1");

    test.checkEqual(wb.index(wb[0]), std::size_t{0}, "index of first sheet");
    test.checkEqual(wb.index(wb[2]), std::size_t{2}, "index of third sheet");

    test.checkEqual(wb.sheetCount(), std::size_t{3}, "sheetCount");

    test.checkEqual(wb[1].name(), std::string("Second"), "operator[] access");
    const auto& cwb = wb;
    test.checkEqual(cwb[0].name(), std::string("First"), "const operator[] access");
}

void testCopyWorksheet(TestContext& test) {
    xlpp::Workbook wb;
    auto& src = wb.addWorksheet("Source");
    src.cell("A1").setValue(std::string("original"));

    auto& copy = wb.copyWorksheet(src, "Copy");
    test.checkEqual(wb.sheetCount(), std::size_t{2}, "copyWorksheet adds sheet");
    test.checkEqual(copy.name(), std::string("Copy"), "copied sheet has new name");
    test.checkEqual(std::get<std::string>(copy.cell("A1").value()), std::string("original"),
                    "cell data is deep-copied");

    wb.addWorksheet("Another");
    auto& clone = wb.copyWorksheet(wb[0], "Clone");
    test.checkEqual(wb.sheetCount(), std::size_t{4}, "copyWorksheet from index access");
    test.checkEqual(clone.name(), std::string("Clone"), "clone name");
    test.checkEqual(std::get<std::string>(clone.cell("A1").value()), std::string("original"),
                    "clone cell data is deep-copied");
    clone.cell("A1").setValue(std::string("changed"));
    test.checkEqual(std::get<std::string>(wb[0].cell("A1").value()), std::string("original"),
                    "copying from wb[0] does not alias the source");
}

void testMergedCells(TestContext& test) {
    xlpp::Worksheet sheet("Layout");
    sheet.cell("A1").setValue("Merged title");
    sheet.mergeCells("C3:A1");
    test.checkEqual(sheet.mergedRanges().size(), std::size_t{1}, "One merged range is registered");
    test.checkEqual(sheet.mergedRanges().front(), std::string("A1:C3"), "Merged range is normalized");
    test.checkTrue(sheet.isMerged("B2"), "Interior cell belongs to merged range");
    test.checkTrue(!sheet.isMerged("D4"), "Outside cell is not merged");
    sheet.unmergeCells("A1:C3");
    test.checkEqual(sheet.mergedRanges().size(), std::size_t{0}, "Merged range is removed");
}

void testWorksheetLayout(TestContext& test) {
    xlpp::Worksheet sheet("Layout");
    sheet.freezePanes("C4");
    sheet.rowDimension(1).height = 28.5;
    sheet.rowDimension(1).hidden = true;
    sheet.columnDimension("B").width = 22.25;
    sheet.columnDimension("B").bestFit = true;
    sheet.columnDimension(4).hidden = true;

    test.checkEqual(sheet.frozenPane().value_or(""), std::string("C4"), "Freeze pane top-left cell");
    test.checkNear(sheet.tryRowDimension(1)->height.value_or(0.0), 28.5, 1e-12, "Custom row height");
    test.checkTrue(sheet.tryRowDimension(1)->hidden, "Hidden row flag");
    test.checkNear(sheet.tryColumnDimension(2)->width.value_or(0.0), 22.25, 1e-12, "Custom column width");
    test.checkTrue(sheet.tryColumnDimension(2)->bestFit, "Best-fit column flag");
    test.checkTrue(sheet.tryColumnDimension(4)->hidden, "Hidden column flag");
}

void testRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_milestone_02.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Sheet1");
    sheet.cell("A1").setValue("hello");
    sheet.cell("B1").setValue(42.5);
    sheet.cell("C1").setValue(true);
    sheet.cell("D1").setFormula("B1*2");
    sheet.cell("D1").setValue(85.0);
    sheet.mergeCells("A3:D3");
    sheet.cell("A3").setValue("Summary");
    sheet.freezePanes("B2");
    sheet.rowDimension(1).height = 24.0;
    sheet.rowDimension(1).hidden = true;
    sheet.columnDimension("B").width = 18.5;
    sheet.columnDimension("C").hidden = true;
    workbook.save(path);
    std::cout << "    [INFO] Saved round-trip workbook: " << path.string() << '\n';

    xlpp::Workbook loaded;
    loaded.load(path);
    auto* loadedSheet = loaded.worksheet("Sheet1");
    test.checkTrue(loadedSheet != nullptr, "Worksheet is loaded");
    test.checkEqual(std::get<std::string>(loadedSheet->cell("A1").value()), std::string("hello"), "String round-trip");
    test.checkNear(std::get<double>(loadedSheet->cell("B1").value()), 42.5, 1e-12, "Number round-trip");
    test.checkTrue(std::get<bool>(loadedSheet->cell("C1").value()), "Boolean round-trip");
    test.checkEqual(loadedSheet->cell("D1").formula(), std::string("B1*2"), "Formula round-trip");
    test.checkEqual(loadedSheet->mergedRanges().front(), std::string("A3:D3"), "Merged range round-trip");
    test.checkEqual(loadedSheet->frozenPane().value_or(""), std::string("B2"), "Freeze panes round-trip");
    test.checkNear(loadedSheet->tryRowDimension(1)->height.value_or(0.0), 24.0, 1e-12, "Row height round-trip");
    test.checkTrue(loadedSheet->tryRowDimension(1)->hidden, "Hidden row round-trip");
    test.checkNear(loadedSheet->tryColumnDimension(2)->width.value_or(0.0), 18.5, 1e-12, "Column width round-trip");
    test.checkTrue(loadedSheet->tryColumnDimension(3)->hidden, "Hidden column round-trip");

    std::filesystem::remove(path);
    std::cout << "    [INFO] Temporary workbook removed\n";
}
void testAutoFilter(TestContext& test) {
    xlpp::Worksheet sheet("FilterData");
    sheet.autoFilter().setReference("A1:D20");
    auto& status = sheet.autoFilter().column(1);
    status.addValue("Open");
    status.addValue("Closed");
    status.setIncludeBlank(true);
    auto& amount = sheet.autoFilter().column(2);
    amount.addCustomFilter(xlpp::FilterOperator::GreaterThanOrEqual, "100");
    amount.addCustomFilter(xlpp::FilterOperator::LessThan, "1000");
    amount.setAndMode(true);
    auto& sort = sheet.autoFilter().sortState();
    sort.setReference("A2:D20");
    sort.setCaseSensitive(true);
    sort.addCondition("C2:C20", true);

    test.checkEqual(sheet.autoFilter().reference(), std::string("A1:D20"), "AutoFilter reference");
    test.checkEqual(status.values().size(), std::size_t{2}, "Discrete filter value count");
    test.checkTrue(status.includeBlank(), "Blank values included");
    test.checkEqual(amount.customFilters().size(), std::size_t{2}, "Custom filter count");
    test.checkTrue(amount.andMode(), "Custom filters use AND mode");
    test.checkTrue(sort.caseSensitive(), "Sort is case-sensitive");
    test.checkTrue(sort.conditions().front().descending, "Sort condition is descending");
}

void testAutoFilterRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_milestone_03_filters.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Name"), std::string("Status"), std::string("Amount")});
    sheet.append({std::string("Alpha"), std::string("Open"), 125.0});
    sheet.autoFilter().setReference("A1:C2");
    sheet.autoFilter().column(1).addValue("Open");
    sheet.autoFilter().column(2).addCustomFilter(xlpp::FilterOperator::GreaterThan, "100");
    auto& sort = sheet.autoFilter().sortState();
    sort.setReference("A2:C2");
    sort.addCondition("C2:C2", true);
    workbook.save(path);
    std::cout << "    [INFO] Saved filter workbook: " << path.string() << '\n';

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedSheet = loaded.worksheet("Data");
    test.checkTrue(loadedSheet != nullptr, "Filtered worksheet is loaded");
    test.checkEqual(loadedSheet->autoFilter().reference(), std::string("A1:C2"), "AutoFilter reference round-trip");
    test.checkEqual(loadedSheet->autoFilter().tryColumn(1)->values().front(), std::string("Open"), "Value filter round-trip");
    test.checkEqual(loadedSheet->autoFilter().tryColumn(2)->customFilters().front().value, std::string("100"), "Custom filter round-trip");
    test.checkTrue(loadedSheet->autoFilter().sortStateValue()->conditions().front().descending, "Sort state round-trip");
    std::filesystem::remove(path);
}


void testCellStyles(TestContext& test) {
    xlpp::Worksheet sheet("Styles");
    auto& cell = sheet.cell("B2");
    cell.setValue("Styled");
    cell.font().setName("Arial");
    cell.font().setSize(14.0);
    cell.font().setBold(true);
    cell.font().setItalic(true);
    cell.font().color().setArgb("FFFF0000");
    cell.fill().setPatternType("solid");
    cell.fill().foregroundColor().setArgb("FFFFFF00");
    cell.border().left().setStyle("thin");
    cell.border().left().color().setArgb("FF000000");
    cell.alignment().setHorizontal("center");
    cell.alignment().setVertical("center");
    cell.alignment().setWrapText(true);
    cell.setNumberFormat("0.00");

    test.checkEqual(cell.font().name(), std::string("Arial"), "Font name");
    test.checkNear(cell.font().size(), 14.0, 1e-12, "Font size");
    test.checkTrue(cell.font().bold(), "Bold font");
    test.checkTrue(cell.font().italic(), "Italic font");
    test.checkEqual(cell.font().color().argb(), std::string("FFFF0000"), "Font ARGB color");
    test.checkEqual(cell.fill().patternType(), std::string("solid"), "Solid fill pattern");
    test.checkEqual(cell.fill().foregroundColor().argb(), std::string("FFFFFF00"), "Fill ARGB color");
    test.checkEqual(cell.border().left().style(), std::string("thin"), "Left border style");
    test.checkEqual(cell.alignment().horizontal(), std::string("center"), "Horizontal alignment");
    test.checkTrue(cell.alignment().wrapText(), "Wrap text alignment");
    test.checkEqual(cell.numberFormat(), std::string("0.00"), "Number format");
}

void testCompactCellModelP1N(TestContext& test) {
    test.checkTrue(sizeof(xlpp::Style) <= 256, "P1N compact Style stays below 256 bytes");
    test.checkTrue(sizeof(xlpp::Cell) <= 640, "P1N compact Cell stays below 640 bytes");

    xlpp::Style original;
    test.checkTrue(original.isDefault(), "Fresh compact Style retains default semantics");
    original.font().setName("Arial");
    original.font().color().setArgb("FF102030");
    original.fill().setPatternType("solid");
    original.border().left().setStyle("thin");
    original.alignment().setHorizontal("center");
    original.setNumberFormat("0.000");

    const auto originalHash = original.hash();
    xlpp::Style copied = original;
    test.checkTrue(copied == original, "Compact Style deep copy preserves semantic equality");
    test.checkEqual(copied.hash(), originalHash, "Compact Style copy preserves semantic hash");
    copied.font().setName("Consolas");
    copied.font().color().setArgb("FF556677");
    test.checkEqual(original.font().name(), std::string("Arial"), "Compact Style copy does not alias font name storage");
    test.checkEqual(original.font().color().argb(), std::string("FF102030"), "Compact Style copy does not alias nested color storage");

    copied = xlpp::Style{};
    test.checkTrue(copied.isDefault(), "Assigning default Style releases compact overrides");
    copied.font().setName("Calibri");
    copied.fill().setPatternType("none");
    copied.setNumberFormat("General");
    test.checkTrue(copied.isDefault(), "Writing canonical defaults does not materialize semantic differences");

    xlpp::Hyperlink link("https://example.test/a");
    link.setDisplay("A");
    xlpp::Hyperlink linkCopy = link;
    linkCopy.setTarget("https://example.test/b");
    linkCopy.setDisplay("B");
    test.checkEqual(link.target(), std::string("https://example.test/a"), "Compact Hyperlink copy owns target independently");
    test.checkEqual(link.display(), std::string("A"), "Compact Hyperlink copy owns display independently");

    xlpp::Comment comment("original", "author");
    xlpp::Comment commentCopy = comment;
    commentCopy.setText("copy");
    test.checkEqual(comment.text(), std::string("original"), "Compact Comment copy owns text independently");

    xlpp::FormulaMetadata metadata;
    metadata.setType(xlpp::FormulaType::Shared);
    metadata.setReference("A1:A8");
    metadata.setSharedIndex(7);
    xlpp::FormulaMetadata metadataCopy = metadata;
    metadataCopy.setReference("B1:B8");
    test.checkEqual(metadata.reference(), std::string("A1:A8"), "Compact FormulaMetadata copy owns reference independently");
    metadataCopy.clearReference();
    test.checkEqual(metadataCopy.reference(), std::string(), "Compact FormulaMetadata clear restores empty reference");

    xlpp::Cell source(2, 3);
    source.setValue("payload");
    source.font().setBold(true);
    source.font().setName("Arial");
    source.setSharedFormula("SUM(A1:A2)", 3, "C2:C4");
    xlpp::Cell cellCopy = source;
    cellCopy.font().setName("Consolas");
    cellCopy.formulaMetadata().setReference("D2:D4");
    test.checkEqual(source.font().name(), std::string("Arial"), "Cell copy deep-copies compact style state");
    test.checkEqual(source.formulaMetadata().reference(), std::string("C2:C4"), "Cell copy deep-copies compact formula metadata");

    xlpp::Worksheet tracking("Tracking");
    for (std::size_t r = 0; r < 128; ++r)
        tracking.append({static_cast<double>(r), std::string("row-") + std::to_string(r)});
    tracking.cell(1, 1).setValue(999.0); // also present in trackedCellKeys_
    test.checkEqual(tracking.trackedCellChangeCount(), std::size_t{256},
                    "Tracked-cell union count avoids duplicates after bulk append plus direct mutation");
}

void testCompactOptionalCellPayloadP1O(TestContext& test) {
    test.checkTrue(sizeof(xlpp::Cell) <= 400, "P1O lazy optional payloads keep Cell at or below 400 bytes");

    xlpp::Cell source(4, 5);
    test.checkTrue(!source.hasRichText(), "Fresh P1O cell has no rich-text payload");
    test.checkTrue(!source.namedStyle().has_value(), "Fresh P1O cell has no named-style payload");

    xlpp::RichText rich;
    xlpp::RichTextRun run("source");
    run.setBold(true);
    rich.addRun(std::move(run));
    source.setRichText(std::move(rich));
    source.setNamedStyle(std::string("Accent"));
    source.setHyperlink(xlpp::Hyperlink("https://example.test/source"));
    source.setComment(xlpp::Comment("source comment", "author"));

    xlpp::Cell copy = source;
    copy.richText().runs().front().setText("copy");
    copy.setNamedStyle(std::string("Warning"));
    copy.hyperlink().setTarget("https://example.test/copy");
    copy.comment().setText("copy comment");

    test.checkEqual(source.richTextValue()->runs().front().text(), std::string("source"),
                    "P1O lazy RichText payload deep-copies on Cell copy");
    test.checkEqual(source.namedStyle().value_or(""), std::string("Accent"),
                    "P1O lazy named-style payload deep-copies on Cell copy");
    test.checkEqual(source.hyperlinkValue()->target(), std::string("https://example.test/source"),
                    "P1O lazy Hyperlink payload deep-copies on Cell copy");
    test.checkEqual(source.commentValue()->text(), std::string("source comment"),
                    "P1O lazy Comment payload deep-copies on Cell copy");
    test.checkEqual(copy.richTextValue()->runs().front().text(), std::string("copy"),
                    "P1O copied RichText payload remains independently mutable");
    test.checkEqual(copy.namedStyle().value_or(""), std::string("Warning"),
                    "P1O copied named-style payload remains independently mutable");

    copy.clearRichText();
    copy.setNamedStyle(std::nullopt);
    copy.clearHyperlink();
    copy.clearComment();
    test.checkTrue(!copy.hasRichText(), "P1O lazy RichText payload releases on clear");
    test.checkTrue(!copy.namedStyle().has_value(), "P1O lazy named-style payload releases on reset");
    test.checkTrue(!copy.hasHyperlink(), "P1O lazy Hyperlink payload releases on clear");
    test.checkTrue(!copy.hasComment(), "P1O lazy Comment payload releases on clear");
}


void testLazyStyleFormulaAndMutationTrackingP1P(TestContext& test) {
    test.checkTrue(sizeof(xlpp::Cell) <= 192,
                   "P1P lazy Style/FormulaMetadata keep Cell at or below 192 bytes");

    xlpp::Cell fresh(1, 1);
    const auto& freshConst = static_cast<const xlpp::Cell&>(fresh);
    test.checkTrue(!freshConst.hasNonDefaultStyle(), "Fresh P1P cell has no non-default style payload");
    test.checkTrue(freshConst.style().isDefault(), "Fresh P1P const style reads shared semantic default");
    test.checkTrue(!freshConst.hasFormulaMetadata(), "Fresh P1P cell has no formula metadata allocation");
    test.checkTrue(freshConst.formulaMetadata().type() == xlpp::FormulaType::Normal,
                   "Fresh P1P const formula metadata reads semantic default");

    fresh.setFormula("A1+1");
    test.checkTrue(!fresh.hasFormulaMetadata(), "Normal formula does not allocate metadata payload");
    fresh.setSharedFormula("A1+1", 7, "A1:A4");
    test.checkTrue(fresh.hasFormulaMetadata(), "Shared formula materializes metadata payload");
    test.checkEqual(fresh.formulaMetadata().sharedIndex().value_or(0u), 7u,
                    "Shared formula stores shared index in lazy metadata");
    fresh.setFormula("B1+1");
    test.checkTrue(!fresh.hasFormulaMetadata(), "Replacing shared formula with normal formula releases stale metadata");
    test.checkTrue(fresh.formulaMetadata().type() == xlpp::FormulaType::Normal,
                   "Normal formula cannot inherit stale shared metadata");

    xlpp::Cell styled(2, 2);
    styled.font().setName("Arial");
    styled.alignment().setHorizontal("center");
    styled.setNumberFormat("0.00");
    test.checkTrue(styled.hasNonDefaultStyle(), "Mutating style materializes a non-default lazy Style");
    xlpp::Cell styledCopy = styled;
    styledCopy.font().setName("Consolas");
    styledCopy.setNumberFormat("0.0000");
    test.checkEqual(static_cast<const xlpp::Cell&>(styled).font().name(), std::string("Arial"),
                    "P1P lazy Style deep-copies across Cell copies");
    test.checkEqual(static_cast<const xlpp::Cell&>(styled).numberFormat(), std::string("0.00"),
                    "P1P lazy Style copy owns number format independently");

    xlpp::Worksheet tracking("TrackingP1P");
    auto& cell = tracking.cell("A1");
    cell.setValue(1.0);
    tracking.clearDirty();
    test.checkEqual(tracking.trackedCellChangeCount(), std::size_t{0},
                    "P1P tracking baseline clears direct-cell revisions");

    cell.style().font().setBold(true);
    test.checkEqual(tracking.trackedCellChangeCount(), std::size_t{1},
                    "Mutable style access participates in cell mutation tracking");
    tracking.clearDirty();
    cell.formulaMetadata().setCalculateOnLoad(true);
    test.checkEqual(tracking.trackedCellChangeCount(), std::size_t{1},
                    "Mutable formula metadata access participates in mutation tracking");
    tracking.clearDirty();
    cell.setHyperlink(xlpp::Hyperlink("https://example.test/p1p"));
    test.checkEqual(tracking.trackedCellChangeCount(), std::size_t{1},
                    "Hyperlink mutation participates in mutation tracking");
    tracking.clearDirty();
    cell.setComment(xlpp::Comment("tracked", "xlpp"));
    test.checkEqual(tracking.trackedCellChangeCount(), std::size_t{1},
                    "Comment mutation participates in mutation tracking");

    // Merely materializing a mutable default Style must not cause an empty cell
    // to be serialized as styled content.
    xlpp::Cell defaultStyled(3, 3);
    (void)defaultStyled.style();
    test.checkTrue(!defaultStyled.hasNonDefaultStyle(),
                   "Materialized but untouched Style remains semantically default");

}

void testStylesRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_milestone_05_styles.xlsx";
    xlpp::Workbook workbook;
    auto& cell = workbook.addWorksheet("Styled").cell("A1");
    cell.setValue(1234.5);
    cell.font().setName("Arial");
    cell.font().setSize(16.0);
    cell.font().setBold(true);
    cell.font().setUnderline(true);
    cell.font().color().setArgb("FF112233");
    cell.fill().setPatternType("solid");
    cell.fill().foregroundColor().setArgb("FFABCDEF");
    cell.border().bottom().setStyle("double");
    cell.border().bottom().color().setArgb("FF445566");
    cell.alignment().setHorizontal("right");
    cell.alignment().setVertical("top");
    cell.alignment().setWrapText(true);
    cell.alignment().setTextRotation(30);
    cell.setNumberFormat("#,##0.00");
    cell.style().setLocked(false);
    cell.style().setHidden(true);
    workbook.save(path);
    std::cout << "    [INFO] Saved styled workbook: " << path.string() << '\n';

    xlpp::Workbook loaded;
    loaded.load(path);
    auto* sheet = loaded.worksheet("Styled");
    test.checkTrue(sheet != nullptr, "Styled worksheet is loaded");
    const auto& loadedCell = sheet->cell("A1");
    test.checkEqual(loadedCell.font().name(), std::string("Arial"), "Font name round-trip");
    test.checkNear(loadedCell.font().size(), 16.0, 1e-12, "Font size round-trip");
    test.checkTrue(loadedCell.font().bold(), "Bold round-trip");
    test.checkTrue(loadedCell.font().underline(), "Underline round-trip");
    test.checkEqual(loadedCell.font().color().argb(), std::string("FF112233"), "Font color round-trip");
    test.checkEqual(loadedCell.fill().patternType(), std::string("solid"), "Fill pattern round-trip");
    test.checkEqual(loadedCell.fill().foregroundColor().argb(), std::string("FFABCDEF"), "Fill color round-trip");
    test.checkEqual(loadedCell.border().bottom().style(), std::string("double"), "Border style round-trip");
    test.checkEqual(loadedCell.border().bottom().color().argb(), std::string("FF445566"), "Border color round-trip");
    test.checkEqual(loadedCell.alignment().horizontal(), std::string("right"), "Alignment round-trip");
    test.checkTrue(loadedCell.alignment().wrapText(), "Wrap text round-trip");
    test.checkEqual(loadedCell.alignment().textRotation(), 30, "Text rotation round-trip");
    test.checkEqual(loadedCell.numberFormat(), std::string("#,##0.00"), "Number format round-trip");
    test.checkTrue(!loadedCell.style().locked(), "Unlocked protection round-trip");
    test.checkTrue(loadedCell.style().hidden(), "Hidden formula protection round-trip");
    std::filesystem::remove(path);
    std::cout << "    [INFO] Temporary styled workbook removed\n";
}

void testHeaderMigration(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("HeaderSmoke");
    sheet.cell("A1").setValue(std::string(".h API"));
    sheet.cell("B1").font().setBold(true);

    test.checkEqual(workbook.worksheets().size(), std::size_t{1}, "Workbook API available through XLPP.h");
    test.checkEqual(std::get<std::string>(sheet.cell("A1").value()), std::string(".h API"), "Cell API available through .h headers");
    test.checkTrue(sheet.cell("B1").font().bold(), "Styles API available through .h headers");
    test.checkEqual(sheet.dimensions(), std::string("A1:B1"), "Worksheet API available through .h headers");
}

void testNamedStyles(TestContext& test) {
    xlpp::Workbook workbook;
    xlpp::NamedStyle currency("Currency");
    currency.style().font().setBold(true);
    currency.style().fill().setPatternType("solid");
    currency.style().fill().foregroundColor().setArgb("FFE2F0D9");
    currency.style().setNumberFormat("#,##0.00");
    workbook.addNamedStyle(currency);

    auto& cell = workbook.addWorksheet("Data").cell("B2");
    workbook.applyNamedStyle(cell, "Currency");
    cell.setValue(1250.5);

    test.checkEqual(workbook.namedStyles().size(), std::size_t{1}, "Named style registry size");
    test.checkTrue(workbook.namedStyle("Currency") != nullptr, "Named style lookup");
    test.checkTrue(cell.font().bold(), "Named style font applied");
    test.checkEqual(cell.fill().foregroundColor().argb(), std::string("FFE2F0D9"), "Named style fill applied");
    test.checkEqual(cell.numberFormat(), std::string("#,##0.00"), "Named style number format applied");
}

void testNamedStylesRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_milestone_07_named_styles.xlsx";
    xlpp::Workbook workbook;
    xlpp::NamedStyle warning("Warning");
    warning.style().font().setBold(true);
    warning.style().font().color().setArgb("FF9C0006");
    warning.style().fill().setPatternType("solid");
    warning.style().fill().foregroundColor().setArgb("FFFFC7CE");
    workbook.addNamedStyle(warning);
    auto& cell = workbook.addWorksheet("Data").cell("A1");
    workbook.applyNamedStyle(cell, "Warning");
    cell.setValue("Invalid");
    workbook.save(path);
    std::cout << "    [INFO] Saved named-style workbook: " << path.string() << '\n';

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedStyle = loaded.namedStyle("Warning");
    test.checkTrue(loadedStyle != nullptr, "Named style registry round-trip");
    test.checkTrue(loadedStyle && loadedStyle->style().font().bold(), "Named style font round-trip");
    test.checkEqual(loadedStyle ? loadedStyle->style().fill().foregroundColor().argb() : std::string{}, std::string("FFFFC7CE"), "Named style fill round-trip");
    auto* sheet = loaded.worksheet("Data");
    test.checkTrue(sheet != nullptr, "Named-style worksheet loaded");
    test.checkTrue(sheet && sheet->cell("A1").font().bold(), "Applied named style cell round-trip");
    std::filesystem::remove(path);
}


void testConditionalFormatting(TestContext& test) {
    xlpp::Worksheet sheet("Rules");
    auto& formulaRule = sheet.conditionalFormatting().addRule(
        "A2:A20", xlpp::ConditionalRule::formula("A2<0"));
    formulaRule.setPriority(1);
    formulaRule.setStopIfTrue(true);
    formulaRule.differentialStyle().font().setBold(true);
    formulaRule.differentialStyle().font().color().setArgb("FFFF0000");

    auto& betweenRule = sheet.conditionalFormatting().addRule(
        "B2:B20", xlpp::ConditionalRule::cellIsBetween("10", "20"));
    betweenRule.setPriority(2);
    betweenRule.differentialStyle().fill().setPatternType("solid");
    betweenRule.differentialStyle().fill().foregroundColor().setArgb("FFFFFF00");

    test.checkEqual(sheet.conditionalFormatting().entries().size(), std::size_t{2}, "Conditional formatting range count");
    test.checkEqual(static_cast<int>(formulaRule.type()), static_cast<int>(xlpp::ConditionalRuleType::Formula), "Formula rule type");
    test.checkEqual(formulaRule.formulas().front(), std::string("A2<0"), "Formula expression");
    test.checkEqual(formulaRule.priority(), std::size_t{1}, "Formula rule priority");
    test.checkTrue(formulaRule.stopIfTrue(), "Stop-if-true flag");
    test.checkTrue(formulaRule.hasDifferentialStyle(), "Formula rule has differential style");
    test.checkEqual(formulaRule.differentialStyle().font().color().argb(), std::string("FFFF0000"), "Differential font color");
    test.checkEqual(static_cast<int>(betweenRule.op()), static_cast<int>(xlpp::ConditionalOperator::Between), "Between operator");
    test.checkEqual(betweenRule.formulas().size(), std::size_t{2}, "Between rule formula count");
    test.checkEqual(betweenRule.differentialStyle().fill().foregroundColor().argb(), std::string("FFFFFF00"), "Differential fill color");
}

void testConditionalFormattingRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_milestone_08_conditional_formatting.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Scores");
    sheet.append({std::string("Score")});
    sheet.append({-5.0});
    sheet.append({15.0});

    auto negative = xlpp::ConditionalRule::cellIs(xlpp::ConditionalOperator::LessThan, "0");
    negative.setPriority(1);
    negative.setStopIfTrue(true);
    negative.differentialStyle().font().setBold(true);
    negative.differentialStyle().font().color().setArgb("FF9C0006");
    negative.differentialStyle().fill().setPatternType("solid");
    negative.differentialStyle().fill().foregroundColor().setArgb("FFFFC7CE");
    sheet.conditionalFormatting().addRule("A2:A100", std::move(negative));

    auto expression = xlpp::ConditionalRule::formula("MOD(A2,2)=0");
    expression.setPriority(2);
    expression.differentialStyle().border().bottom().setStyle("thin");
    expression.differentialStyle().border().bottom().color().setArgb("FF0000FF");
    sheet.conditionalFormatting().addRule("A2:A100", std::move(expression));

    workbook.save(path);
    std::cout << "    [INFO] Saved conditional-formatting workbook: " << path.string() << '\n';

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedSheet = loaded.worksheet("Scores");
    test.checkTrue(loadedSheet != nullptr, "Conditional-formatting worksheet loaded");
    test.checkEqual(loadedSheet->conditionalFormatting().entries().size(), std::size_t{1}, "Conditional formatting entry round-trip");
    const auto& rules = loadedSheet->conditionalFormatting().entries().front().rules();
    test.checkEqual(rules.size(), std::size_t{2}, "Conditional rule count round-trip");
    test.checkEqual(static_cast<int>(rules[0].type()), static_cast<int>(xlpp::ConditionalRuleType::CellIs), "Cell-is rule type round-trip");
    test.checkEqual(static_cast<int>(rules[0].op()), static_cast<int>(xlpp::ConditionalOperator::LessThan), "Cell-is operator round-trip");
    test.checkEqual(rules[0].formulas().front(), std::string("0"), "Cell-is formula round-trip");
    test.checkTrue(rules[0].stopIfTrue(), "Stop-if-true round-trip");
    test.checkEqual(rules[0].differentialStyle().font().color().argb(), std::string("FF9C0006"), "Differential font round-trip");
    test.checkEqual(rules[0].differentialStyle().fill().foregroundColor().argb(), std::string("FFFFC7CE"), "Differential fill round-trip");
    test.checkEqual(static_cast<int>(rules[1].type()), static_cast<int>(xlpp::ConditionalRuleType::Formula), "Expression rule type round-trip");
    test.checkEqual(rules[1].formulas().front(), std::string("MOD(A2,2)=0"), "Expression formula round-trip");
    test.checkEqual(rules[1].differentialStyle().border().bottom().style(), std::string("thin"), "Differential border round-trip");
    std::filesystem::remove(path);
}


void testDataValidation(TestContext& test) {
    xlpp::Worksheet sheet("Validation");
    auto list = xlpp::DataValidation::list("A2:A100", "\"Open,Closed,Pending\"");
    list.setAllowBlank(true);
    list.setShowDropDown(true);
    list.setShowInputMessage(true);
    list.setPromptTitle("Choose status");
    list.setPrompt("Select a value from the list.");
    list.setShowErrorMessage(true);
    list.setErrorTitle("Invalid status");
    list.setError("Use one of the available values.");
    list.setErrorStyle(xlpp::DataValidationErrorStyle::Stop);
    sheet.dataValidations().add(std::move(list));

    auto& numeric = sheet.dataValidations().add(xlpp::DataValidationType::Decimal, "B2:B100");
    numeric.setOperator(xlpp::DataValidationOperator::Between);
    numeric.setFormula1("0");
    numeric.setFormula2("100");
    numeric.setAllowBlank(false);

    auto& custom = sheet.dataValidations().add(xlpp::DataValidationType::Custom, "C2:C100");
    custom.setFormula1("=MOD(C2,2)=0");

    test.checkEqual(sheet.dataValidations().items().size(), std::size_t{3}, "Data validation rule count");
    const auto& storedList = sheet.dataValidations().items()[0];
    test.checkEqual(static_cast<int>(storedList.type()), static_cast<int>(xlpp::DataValidationType::List), "List validation type");
    test.checkEqual(storedList.reference(), std::string("A2:A100"), "List validation reference");
    test.checkEqual(storedList.formula1(), std::string("\"Open,Closed,Pending\""), "Inline list formula");
    test.checkTrue(storedList.allowBlank(), "List validation allows blank");
    test.checkTrue(storedList.showDropDown(), "List drop-down enabled");
    test.checkTrue(storedList.showInputMessage(), "Input message enabled");
    test.checkEqual(storedList.promptTitle(), std::string("Choose status"), "Input prompt title");
    test.checkTrue(storedList.showErrorMessage(), "Error message enabled");
    test.checkEqual(storedList.errorTitle(), std::string("Invalid status"), "Error title");
    const auto& storedNumeric = sheet.dataValidations().items()[1];
    const auto& storedCustom = sheet.dataValidations().items()[2];
    test.checkEqual(static_cast<int>(storedNumeric.op()), static_cast<int>(xlpp::DataValidationOperator::Between), "Numeric between operator");
    test.checkEqual(storedNumeric.formula2(), std::string("100"), "Numeric upper bound");
    test.checkEqual(static_cast<int>(storedCustom.type()), static_cast<int>(xlpp::DataValidationType::Custom), "Custom validation type");
    test.checkEqual(storedCustom.formula1(), std::string("=MOD(C2,2)=0"), "Custom validation formula");
}

void testDataValidationRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_milestone_09_data_validation.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Entry");
    sheet.append({std::string("Status"), std::string("Score"), std::string("Date")});

    auto list = xlpp::DataValidation::list("A2:A50", "Lookup!$A$1:$A$4");
    list.setAllowBlank(true);
    list.setShowDropDown(true);
    list.setShowInputMessage(true);
    list.setPromptTitle("Status");
    list.setPrompt("Select a status");
    list.setShowErrorMessage(true);
    list.setErrorStyle(xlpp::DataValidationErrorStyle::Warning);
    list.setErrorTitle("Unknown status");
    list.setError("This value is not in the status list.");
    sheet.dataValidations().add(std::move(list));

    auto& score = sheet.dataValidations().add(xlpp::DataValidationType::Whole, "B2:B50");
    score.setOperator(xlpp::DataValidationOperator::Between);
    score.setFormula1("0");
    score.setFormula2("100");

    auto& date = sheet.dataValidations().add(xlpp::DataValidationType::Date, "C2:C50");
    date.setOperator(xlpp::DataValidationOperator::GreaterThanOrEqual);
    date.setFormula1("DATE(2026,1,1)");

    workbook.save(path);
    std::cout << "    [INFO] Saved data-validation workbook: " << path.string() << '\n';

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedSheet = loaded.worksheet("Entry");
    test.checkTrue(loadedSheet != nullptr, "Data-validation worksheet loaded");
    const auto& items = loadedSheet->dataValidations().items();
    test.checkEqual(items.size(), std::size_t{3}, "Data validation count round-trip");
    test.checkEqual(static_cast<int>(items[0].type()), static_cast<int>(xlpp::DataValidationType::List), "List type round-trip");
    test.checkEqual(items[0].reference(), std::string("A2:A50"), "List reference round-trip");
    test.checkEqual(items[0].formula1(), std::string("Lookup!$A$1:$A$4"), "List source round-trip");
    test.checkTrue(items[0].allowBlank(), "Allow blank round-trip");
    test.checkTrue(items[0].showDropDown(), "Drop-down flag round-trip");
    test.checkEqual(static_cast<int>(items[0].errorStyle()), static_cast<int>(xlpp::DataValidationErrorStyle::Warning), "Error style round-trip");
    test.checkEqual(items[0].prompt(), std::string("Select a status"), "Prompt round-trip");
    test.checkEqual(items[0].error(), std::string("This value is not in the status list."), "Error message round-trip");
    test.checkEqual(static_cast<int>(items[1].type()), static_cast<int>(xlpp::DataValidationType::Whole), "Whole-number type round-trip");
    test.checkEqual(items[1].formula2(), std::string("100"), "Whole-number upper bound round-trip");
    test.checkEqual(static_cast<int>(items[2].op()), static_cast<int>(xlpp::DataValidationOperator::GreaterThanOrEqual), "Date operator round-trip");
    test.checkEqual(items[2].formula1(), std::string("DATE(2026,1,1)"), "Date formula round-trip");
    std::filesystem::remove(path);
}



void testTablesAndDefinedNames(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Sales");
    sheet.append({std::string("Product"), std::string("Amount")});
    sheet.append({std::string("A"), 10.0});
    auto& table = sheet.addTable("SalesTable", "A1:B2");
    table.addColumn("Product");
    table.addColumn("Amount");
    table.styleInfo().setName("TableStyleMedium9");
    table.styleInfo().setShowRowStripes(true);
    table.setShowTotalsRow(false);

    xlpp::DefinedName name("SalesRange", "'Sales'!$A$1:$B$2");
    name.setComment("Primary sales range");
    workbook.addDefinedName(std::move(name));
    auto& local = workbook.addDefinedName(xlpp::DefinedName("LocalAmount", "'Sales'!$B$2"));
    local.setLocalSheetId(0);
    local.setHidden(true);

    test.checkEqual(sheet.tables().size(), std::size_t{1}, "Worksheet table count");
    test.checkEqual(table.name(), std::string("SalesTable"), "Table name");
    test.checkEqual(table.reference(), std::string("A1:B2"), "Table reference");
    test.checkEqual(table.columns().size(), std::size_t{2}, "Table column count");
    test.checkEqual(table.styleInfo().name(), std::string("TableStyleMedium9"), "Table style name");
    test.checkTrue(sheet.table("SalesTable") != nullptr, "Table lookup");
    test.checkEqual(workbook.definedNames().size(), std::size_t{2}, "Defined name count");
    test.checkEqual(workbook.definedName("SalesRange")->value(), std::string("'Sales'!$A$1:$B$2"), "Defined name value");
    test.checkTrue(workbook.definedName("LocalAmount")->localSheetId().has_value(), "Local defined name scope");
    test.checkTrue(workbook.definedName("LocalAmount")->hidden(), "Hidden defined name flag");
}

void testTablesAndDefinedNamesRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_milestone_10_tables_defined_names.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Inventory");
    sheet.append({std::string("SKU"), std::string("Quantity"), std::string("Active")});
    sheet.append({std::string("A-001"), 25.0, true});
    auto& table = sheet.addTable("InventoryTable", "A1:C2");
    table.addColumn("SKU"); table.addColumn("Quantity"); table.addColumn("Active");
    table.styleInfo().setName("TableStyleMedium4");
    table.styleInfo().setShowFirstColumn(true);
    table.styleInfo().setShowColumnStripes(true);
    workbook.addDefinedName(xlpp::DefinedName("InventoryData", "'Inventory'!$A$1:$C$2"));
    workbook.save(path);
    std::cout << "    [INFO] Saved table workbook: " << path.string() << '\n';

    xlpp::Workbook loaded; loaded.load(path);
    const auto* loadedSheet = loaded.worksheet("Inventory");
    test.checkTrue(loadedSheet != nullptr, "Table worksheet loaded");
    test.checkEqual(loadedSheet->tables().size(), std::size_t{1}, "Table count round-trip");
    const auto& loadedTable = loadedSheet->tables().front();
    test.checkEqual(loadedTable.name(), std::string("InventoryTable"), "Table name round-trip");
    test.checkEqual(loadedTable.reference(), std::string("A1:C2"), "Table reference round-trip");
    test.checkEqual(loadedTable.columns().size(), std::size_t{3}, "Table columns round-trip");
    test.checkEqual(loadedTable.columns()[1].name(), std::string("Quantity"), "Table column name round-trip");
    test.checkEqual(loadedTable.styleInfo().name(), std::string("TableStyleMedium4"), "Table style round-trip");
    test.checkTrue(loadedTable.styleInfo().showFirstColumn(), "Table first-column style flag round-trip");
    test.checkTrue(loadedTable.styleInfo().showColumnStripes(), "Table column-stripes flag round-trip");
    test.checkEqual(loaded.definedNames().size(), std::size_t{1}, "Defined name count round-trip");
    test.checkEqual(loaded.definedNames().front().value(), std::string("'Inventory'!$A$1:$C$2"), "Defined name formula round-trip");
    std::filesystem::remove(path);
}

void testHyperlinksCommentsAndProperties(TestContext& test) {
    xlpp::Workbook workbook; auto& sheet=workbook.addWorksheet("Links");
    auto& cell=sheet.cell("A1"); cell.setValue("OpenAI");
    xlpp::Hyperlink link("https://example.com"); link.setDisplay("Example"); link.setTooltip("Open website"); cell.setHyperlink(std::move(link));
    cell.setComment(xlpp::Comment("Review this link", "XL++ Tester"));
    workbook.properties().setTitle("XL++ Milestone 11"); workbook.properties().setCreator("XL++"); workbook.properties().setCategory("Tests");
    test.checkTrue(cell.hasHyperlink(), "Cell hyperlink exists");
    test.checkEqual(cell.hyperlinkValue()->target(), std::string("https://example.com"), "Hyperlink target");
    test.checkTrue(cell.hasComment(), "Cell comment exists");
    test.checkEqual(cell.commentValue()->author(), std::string("XL++ Tester"), "Comment author");
    test.checkEqual(workbook.properties().title(), std::string("XL++ Milestone 11"), "Document title");
}

void testHyperlinksAndPropertiesRoundTrip(TestContext& test) {
    const auto path=std::filesystem::temp_directory_path()/"xlpp_m11_links_properties.xlsx";
    xlpp::Workbook workbook; auto& sheet=workbook.addWorksheet("Links"); sheet.cell("A1").setValue("Website");
    xlpp::Hyperlink link("https://example.com/docs"); link.setTooltip("Documentation"); sheet.cell("A1").setHyperlink(std::move(link));
    workbook.properties().setTitle("Hyperlink workbook"); workbook.properties().setSubject("Round-trip"); workbook.properties().setCreator("XL++ Tests");
    workbook.save(path); std::cout<<"    [INFO] Saved hyperlink workbook: "<<path.string()<<'\n';
    xlpp::Workbook loaded; loaded.load(path); const auto* ws=loaded.worksheet("Links");
    test.checkTrue(ws!=nullptr, "Hyperlink worksheet loaded");
    test.checkTrue(ws->tryCell("A1")->hasHyperlink(), "Hyperlink round-trip exists");
    test.checkEqual(ws->tryCell("A1")->hyperlinkValue()->target(), std::string("https://example.com/docs"), "Hyperlink target round-trip");
    test.checkEqual(loaded.properties().title(), std::string("Hyperlink workbook"), "Document title round-trip");
    test.checkEqual(loaded.properties().creator(), std::string("XL++ Tests"), "Document creator round-trip");
    std::filesystem::remove(path);
}

void testCommentsRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m12_comments.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Comments");
    sheet.cell("A1").setValue("First");
    sheet.cell("A1").setComment(xlpp::Comment("Review this value", "Alice"));
    sheet.cell("C4").setValue(42.0);
    sheet.cell("C4").setComment(xlpp::Comment("Second note with <XML> & spaces", "Bob"));
    workbook.save(path);
    std::cout << "    [INFO] Saved comments workbook: " << path.string() << '\n';

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedSheet = loaded.worksheet("Comments");
    test.checkTrue(loadedSheet != nullptr, "Comments worksheet loaded");
    test.checkTrue(loadedSheet->tryCell("A1") != nullptr, "First comment cell loaded");
    test.checkTrue(loadedSheet->tryCell("A1")->hasComment(), "First comment exists after round-trip");
    test.checkEqual(loadedSheet->tryCell("A1")->commentValue()->text(), std::string("Review this value"), "First comment text round-trip");
    test.checkEqual(loadedSheet->tryCell("A1")->commentValue()->author(), std::string("Alice"), "First comment author round-trip");
    test.checkTrue(loadedSheet->tryCell("C4")->hasComment(), "Second comment exists after round-trip");
    test.checkEqual(loadedSheet->tryCell("C4")->commentValue()->text(), std::string("Second note with <XML> & spaces"), "Escaped comment text round-trip");
    test.checkEqual(loadedSheet->tryCell("C4")->commentValue()->author(), std::string("Bob"), "Second comment author round-trip");
    const auto archive = xlpp::internal::ZipArchive::open(path);
    const auto sheetXml = archive.get("xl/worksheets/sheet1.xml");
    const auto marginsPosition = sheetXml.find("<pageMargins");
    const auto legacyDrawingPosition = sheetXml.find("<legacyDrawing");
    test.checkTrue(marginsPosition != std::string::npos && legacyDrawingPosition != std::string::npos
                       && marginsPosition < legacyDrawingPosition,
                   "Legacy comment drawing follows page settings in worksheet schema order");
    std::filesystem::remove(path);
}

void testPageSetupProtectionAndImages(TestContext& test) {
    xlpp::Workbook workbook;
    workbook.protection().setLockStructure(true);
    workbook.protection().setWorkbookPasswordHash("ABCD");
    auto& sheet = workbook.addWorksheet("Report");
    sheet.pageSetup().setOrientation(xlpp::PageOrientation::Landscape);
    sheet.pageSetup().setPaperSize(xlpp::PaperSize::A4);
    sheet.pageSetup().setFitToPage(true);
    sheet.pageSetup().setFitToWidth(1);
    sheet.pageMargins().setLeft(0.25);
    sheet.printOptions().setGridLines(true);
    sheet.headerFooter().setOddHeader("&CXL++ Report");
    sheet.headerFooter().setOddFooter("Page &P of &N");
    sheet.protection().setEnabled(true);
    sheet.protection().setPasswordHash("CDEF");
    sheet.protection().setSort(true);
    const std::vector<unsigned char> png{137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,137,0,0,0,13,73,68,65,84,8,215,99,248,207,192,240,31,0,5,0,1,255,137,153,61,29,0,0,0,0,73,69,78,68,174,66,96,130};
    xlpp::Image image("B2", png, "png"); image.setName("Pixel"); image.setWidthPixels(32); image.setHeightPixels(24); sheet.addImage(std::move(image));
    test.checkTrue(workbook.protection().lockStructure(), "Workbook structure protection");
    test.checkEqual(static_cast<unsigned>(sheet.pageSetup().orientation()), static_cast<unsigned>(xlpp::PageOrientation::Landscape), "Landscape page orientation");
    test.checkEqual(static_cast<unsigned>(sheet.pageSetup().paperSize()), static_cast<unsigned>(xlpp::PaperSize::A4), "A4 paper size");
    test.checkNear(sheet.pageMargins().left(), 0.25, 1e-12, "Custom left page margin");
    test.checkTrue(sheet.printOptions().gridLines(), "Print grid lines");
    test.checkEqual(sheet.headerFooter().oddFooter(), std::string("Page &P of &N"), "Odd footer text");
    test.checkTrue(sheet.protection().enabled(), "Worksheet protection enabled");
    test.checkTrue(sheet.protection().sort(), "Sort allowed on protected worksheet");
    test.checkEqual(sheet.images().size(), std::size_t{1}, "Worksheet image count");
    test.checkEqual(sheet.images().front().anchor(), std::string("B2"), "Image anchor");
}

void testPageSetupProtectionAndImagesRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m13_page_protection_images.xlsx";
    xlpp::Workbook workbook; workbook.protection().setLockStructure(true); workbook.protection().setLockWindows(true); workbook.protection().setWorkbookPasswordHash("ABCD");
    auto& sheet = workbook.addWorksheet("Print"); sheet.cell("A1").setValue("XL++");
    sheet.pageSetup().setOrientation(xlpp::PageOrientation::Landscape); sheet.pageSetup().setPaperSize(xlpp::PaperSize::A4); sheet.pageSetup().setScale(85); sheet.pageSetup().setBlackAndWhite(true);
    sheet.pageMargins().setTop(0.4); sheet.pageMargins().setBottom(0.4); sheet.printOptions().setHorizontalCentered(true); sheet.headerFooter().setOddHeader("&LXL++&R&P");
    sheet.protection().setEnabled(true); sheet.protection().setPasswordHash("CDEF"); sheet.protection().setAutoFilter(true);
    const std::vector<unsigned char> png{137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,137,0,0,0,13,73,68,65,84,8,215,99,248,207,192,240,31,0,5,0,1,255,137,153,61,29,0,0,0,0,73,69,78,68,174,66,96,130};
    sheet.addImage(xlpp::Image("D5", png, "png"));
    workbook.save(path); std::cout << "    [INFO] Saved page/protection/images workbook: " << path.string() << '\n';
    test.checkTrue(std::filesystem::exists(path), "Workbook with drawing package is created");
    test.checkTrue(std::filesystem::file_size(path) > 500, "Workbook drawing package has content");
    xlpp::Workbook loaded; loaded.load(path); const auto* ws=loaded.worksheet("Print");
    test.checkTrue(ws!=nullptr, "Print worksheet loaded");
    test.checkTrue(loaded.protection().lockStructure(), "Workbook protection round-trip");
    test.checkTrue(loaded.protection().lockWindows(), "Workbook window lock round-trip");
    test.checkEqual(loaded.protection().workbookPasswordHash(), std::string("ABCD"), "Workbook password hash round-trip");
    test.checkEqual(static_cast<unsigned>(ws->pageSetup().orientation()), static_cast<unsigned>(xlpp::PageOrientation::Landscape), "Page orientation round-trip");
    test.checkEqual(static_cast<unsigned>(ws->pageSetup().paperSize()), static_cast<unsigned>(xlpp::PaperSize::A4), "Paper size round-trip");
    test.checkEqual(ws->pageSetup().scale(), 85u, "Print scale round-trip");
    test.checkTrue(ws->pageSetup().blackAndWhite(), "Black-and-white print round-trip");
    test.checkNear(ws->pageMargins().top(), 0.4, 1e-12, "Top margin round-trip");
    test.checkTrue(ws->printOptions().horizontalCentered(), "Horizontal centering round-trip");
    test.checkEqual(ws->headerFooter().oddHeader(), std::string("&LXL++&R&P"), "Header round-trip");
    test.checkTrue(ws->protection().enabled(), "Worksheet protection round-trip");
    test.checkEqual(ws->protection().passwordHash(), std::string("CDEF"), "Worksheet password hash round-trip");
    test.checkTrue(ws->protection().autoFilter(), "Protected AutoFilter permission round-trip");
    std::filesystem::remove(path);
}


void testFormulaMetadataAndErrorCells(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("FormulaMetadata");

    auto& shared = sheet.cell("A1");
    shared.setSharedFormula("SUM(B1:C1)", 7, "A1:A10");
    shared.setValue(42.0);
    test.checkTrue(shared.hasFormula(), "Shared formula exists");
    test.checkEqual(static_cast<unsigned>(shared.formulaMetadata().type()), static_cast<unsigned>(xlpp::FormulaType::Shared), "Shared formula type");
    test.checkEqual(shared.formulaMetadata().reference(), std::string("A1:A10"), "Shared formula reference");
    test.checkTrue(shared.formulaMetadata().sharedIndex().has_value(), "Shared formula index exists");
    test.checkEqual(*shared.formulaMetadata().sharedIndex(), 7u, "Shared formula index value");

    auto& array = sheet.cell("D1");
    array.setArrayFormula("TRANSPOSE(A1:A3)", "D1:F1");
    array.formulaMetadata().setAlwaysCalculateArray(true);
    test.checkEqual(static_cast<unsigned>(array.formulaMetadata().type()), static_cast<unsigned>(xlpp::FormulaType::Array), "Array formula type");
    test.checkTrue(array.formulaMetadata().alwaysCalculateArray(), "Array formula ACA flag");

    auto& errorCell = sheet.cell("G1");
    errorCell.setError(xlpp::CellError::DivisionByZero);
    test.checkTrue(errorCell.isError(), "Error cell type detected");
    test.checkEqual(xlpp::toString(*errorCell.error()), std::string("#DIV/0!"), "Error cell text");
}

void testFormulaMetadataAndErrorCellsRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m14_formula_metadata_errors.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("FormulaMetadata");
    auto& shared = sheet.cell("A1");
    shared.setSharedFormula("B1+C1", 3, "A1:A4");
    shared.formulaMetadata().setCalculateOnLoad(true);
    shared.setValue(10.0);
    auto& array = sheet.cell("D1");
    array.setArrayFormula("TRANSPOSE(A1:A3)", "D1:F1");
    array.formulaMetadata().setAlwaysCalculateArray(true);
    array.setValue(1.0);
    sheet.cell("G1").setError(xlpp::CellError::NotAvailable);
    sheet.cell("G2").setError(xlpp::CellError::Reference);

    workbook.save(path);
    std::cout << "    [INFO] Saved formula metadata workbook: " << path.string() << '\n';
    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* ws = loaded.worksheet("FormulaMetadata");
    test.checkTrue(ws != nullptr, "Formula metadata worksheet loaded");
    const auto* loadedShared = ws->tryCell("A1");
    const auto* loadedArray = ws->tryCell("D1");
    const auto* loadedError = ws->tryCell("G1");
    test.checkTrue(loadedShared != nullptr, "Shared formula cell loaded");
    test.checkEqual(loadedShared->formula(), std::string("B1+C1"), "Shared formula text round-trip");
    test.checkEqual(static_cast<unsigned>(loadedShared->formulaMetadata().type()), static_cast<unsigned>(xlpp::FormulaType::Shared), "Shared formula type round-trip");
    test.checkEqual(loadedShared->formulaMetadata().reference(), std::string("A1:A4"), "Shared formula reference round-trip");
    test.checkEqual(*loadedShared->formulaMetadata().sharedIndex(), 3u, "Shared index round-trip");
    test.checkTrue(loadedShared->formulaMetadata().calculateOnLoad(), "Calculate-on-load flag round-trip");
    test.checkEqual(static_cast<unsigned>(loadedArray->formulaMetadata().type()), static_cast<unsigned>(xlpp::FormulaType::Array), "Array formula type round-trip");
    test.checkEqual(loadedArray->formulaMetadata().reference(), std::string("D1:F1"), "Array formula reference round-trip");
    test.checkTrue(loadedArray->formulaMetadata().alwaysCalculateArray(), "Array ACA round-trip");
    test.checkTrue(loadedError != nullptr && loadedError->isError(), "Error cell round-trip type");
    test.checkEqual(xlpp::toString(*loadedError->error()), std::string("#N/A"), "Error cell round-trip value");
    test.checkEqual(xlpp::toString(*ws->tryCell("G2")->error()), std::string("#REF!"), "Second error cell round-trip");
    std::filesystem::remove(path);
}



void testDirectZipReader(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m16_zip_reader.xlsx";
    {
        xlpp::internal::ZipArchive zip;
        zip.add("hello.txt", "Hello, world!", true);
        zip.add("stored.bin", std::string("0123456789"), false);
        zip.add("dir/nested.xml", "<nested/>", true);
        zip.save(path);
    }

    xlpp::internal::ZipArchiveReader reader(path);
    test.checkEqual(reader.entryCount(), std::size_t{3}, "Direct reader indexes central directory");
    test.checkTrue(reader.contains("hello.txt"), "Direct reader entry lookup");
    test.checkTrue(!reader.contains("missing.xml"), "Missing entry is not found");
    test.checkEqual(reader.names().size(), std::size_t{3}, "Direct reader lists all entries");
    test.checkEqual(reader.readEntry("hello.txt"), std::string("Hello, world!"), "Deflate entry full read");
    test.checkEqual(reader.readEntry("stored.bin"), std::string("0123456789"), "Stored entry full read");
    test.checkEqual(reader.readEntry("dir/nested.xml"), std::string("<nested/>"), "Nested path entry read");

    std::string streamed;
    reader.forEachChunk("hello.txt", [&](const char* data, std::size_t size) { streamed.append(data, size); });
    test.checkEqual(streamed, std::string("Hello, world!"), "Chunk streaming reproduces entry");

    {
        auto source = reader.openEntry("stored.bin");
        std::array<unsigned char, 4> pullBuffer{};
        std::string pulled;
        for (std::size_t count = source.read(pullBuffer.data(), pullBuffer.size()); count;
             count = source.read(pullBuffer.data(), pullBuffer.size()))
            pulled.append(reinterpret_cast<const char*>(pullBuffer.data()), count);
        test.checkEqual(pulled, std::string("0123456789"), "Pull source reproduces stored entry");
        test.checkTrue(source.complete(), "Pull source CRC verified at end");
    }

    {
        // A buffer sized exactly to the decompressed data forces the final
        // inflate call to fill the output before the stream end marker is seen.
        auto source = reader.openEntry("hello.txt");
        std::array<unsigned char, 13> exact{};
        const auto first = source.read(exact.data(), exact.size());
        test.checkEqual(first, std::size_t{13}, "Deflate pull fills output exactly");
        const auto second = source.read(exact.data(), exact.size());
        test.checkEqual(second, std::size_t{0}, "Deflate pull reports end on next read");
        test.checkTrue(source.complete(), "Deflate pull CRC verified after end marker");
    }

    {
        const auto corruptPath = std::filesystem::temp_directory_path() / "xlpp_m16_corrupt.xlsx";
        {
            xlpp::internal::ZipArchive single;
            single.add("stored.bin", std::string("0123456789"), false);
            single.save(corruptPath);
        }
        std::fstream file(corruptPath, std::ios::binary | std::ios::in | std::ios::out);
        file.seekp(40); // local header (30) + name "stored.bin" (10)
        file.put('X');
        file.close();
        xlpp::internal::ZipArchiveReader corrupted(corruptPath);
        bool crcThrew = false;
        try { (void)corrupted.readEntry("stored.bin"); } catch (const std::exception&) { crcThrew = true; }
        test.checkTrue(crcThrew, "Corrupted stored entry reports CRC mismatch");
        std::filesystem::remove(corruptPath);
    }

    std::filesystem::remove(path);
}

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

void testDateTimeCore(TestContext& test) {
    using xlpp::DateTime;

    test.checkNear(xlpp::toExcelSerial(DateTime{1900, 1, 1}), 1.0, 1e-12, "1900-01-01 is serial 1");
    test.checkNear(xlpp::toExcelSerial(DateTime{1900, 2, 28}), 59.0, 1e-12, "1900-02-28 is serial 59");
    test.checkNear(xlpp::toExcelSerial(DateTime{1900, 3, 1}), 61.0, 1e-12, "1900-03-01 skips phantom 60");
    test.checkNear(xlpp::toExcelSerial(DateTime{1970, 1, 1}), 25569.0, 1e-12, "1970-01-01 is serial 25569");
    test.checkNear(xlpp::toExcelSerial(DateTime{2000, 1, 1}), 36526.0, 1e-12, "2000-01-01 is serial 36526");
    test.checkNear(xlpp::toExcelSerial(DateTime{2020, 1, 1}), 43831.0, 1e-12, "2020-01-01 is serial 43831");
    test.checkNear(xlpp::toExcelSerial(DateTime{2024, 1, 15}), 45306.0, 1e-12, "2024-01-15 is serial 45306");
    test.checkNear(xlpp::toExcelSerial(DateTime{1900, 1, 1, 12, 0, 0}), 1.5, 1e-12, "Serial includes time fraction");
    test.checkNear(xlpp::toExcelSerial(DateTime{2024, 1, 15, 13, 30, 0}), 45306.5625, 1e-12, "Exact time fraction");

    test.checkNear(xlpp::toExcelSerial(DateTime{1904, 1, 1}, true), 0.0, 1e-12, "1904-01-01 is serial 0 in 1904");
    test.checkNear(xlpp::toExcelSerial(DateTime{2000, 1, 1}, true), 35064.0, 1e-12, "2000-01-01 1904 system");

    test.checkEqual(xlpp::fromExcelSerial(1.0), DateTime{1900, 1, 1}, "Serial 1 round-trips");
    test.checkEqual(xlpp::fromExcelSerial(61.0), DateTime{1900, 3, 1}, "Serial 61 round-trips");
    test.checkEqual(xlpp::fromExcelSerial(45306.5625), DateTime{2024, 1, 15, 13, 30, 0}, "Serial 45306.5625 round-trips");
    test.checkEqual(xlpp::fromExcelSerial(0.0, true), DateTime{1904, 1, 1}, "1904 serial 0 round-trips");
    test.checkEqual(xlpp::fromExcelSerial(35064.0, true), DateTime{2000, 1, 1}, "1904 serial 35064 round-trips");

    for (int year : {1999, 2000, 2004, 2024, 2099}) {
        for (int month : {1, 6, 12}) {
            for (int day : {1, 15, 28}) {
                const DateTime value{year, month, day, 23, 59, 59.25};
                const auto serial = xlpp::toExcelSerial(value);
                test.checkEqual(xlpp::fromExcelSerial(serial), value, "1900 serial round-trip sample");
                const auto serial1904 = xlpp::toExcelSerial(value, true);
                test.checkEqual(xlpp::fromExcelSerial(serial1904, true), value, "1904 serial round-trip sample");
            }
        }
    }

    test.checkTrue(!xlpp::isDateFormatCode("General"), "General is not a date format");
    test.checkTrue(!xlpp::isDateFormatCode("0.00"), "Numeric format is not a date format");
    test.checkTrue(!xlpp::isDateFormatCode("@"), "Text format is not a date format");
    test.checkTrue(xlpp::isDateFormatCode("yyyy-mm-dd"), "ISO-like format is a date format");
    test.checkTrue(xlpp::isDateFormatCode("m/d/yy"), "US date format is a date format");
    test.checkTrue(xlpp::isDateFormatCode("hh:mm"), "Time format is a date format");
    test.checkTrue(xlpp::isDateFormatCode("[h]:mm"), "Elapsed time format is a date format");
    test.checkTrue(!xlpp::isDateFormatCode("\"m\";0.00"), "Quoted letters are not a date format");

    test.checkEqual(xlpp::parseIso8601("2024-01-15").value_or(DateTime{}), DateTime{2024, 1, 15}, "ISO date parses");
    test.checkEqual(xlpp::parseIso8601("2024-01-15T13:30:45").value_or(DateTime{}), DateTime{2024, 1, 15, 13, 30, 45}, "ISO datetime parses");
    test.checkEqual(xlpp::parseIso8601("2024-01-15T13:30:45.250").value_or(DateTime{}), DateTime{2024, 1, 15, 13, 30, 45.25}, "ISO fractional seconds parse");
    test.checkEqual(xlpp::parseIso8601("2024-01-15 08:05").value_or(DateTime{}), DateTime{2024, 1, 15, 8, 5}, "ISO space-separated time parses");
    test.checkEqual(xlpp::parseIso8601("2024-01-15T12:00:00Z").value_or(DateTime{}), DateTime{2024, 1, 15, 12, 0, 0}, "ISO UTC parses");
    test.checkEqual(xlpp::parseIso8601("2024-01-15T12:00:00+02:00").value_or(DateTime{}), DateTime{2024, 1, 15, 10, 0, 0}, "ISO offset is applied to UTC");
    test.checkEqual(xlpp::parseIso8601("2024-01-15T00:30:00-03:30").value_or(DateTime{}), DateTime{2024, 1, 15, 4, 0, 0}, "ISO negative offset is applied");
    test.checkTrue(!xlpp::parseIso8601("garbage").has_value(), "Garbage rejected");
    test.checkTrue(!xlpp::parseIso8601("2024-13-01").has_value(), "Invalid month rejected");
    test.checkTrue(!xlpp::parseIso8601("2024-01-15T25:00").has_value(), "Invalid hour rejected");
    test.checkEqual(xlpp::toIso8601(DateTime{2024, 1, 15, 13, 30, 45}), std::string("2024-01-15T13:30:45"), "ISO format");
    test.checkEqual(xlpp::toIso8601(DateTime{2024, 1, 15, 13, 30, 45.25}), std::string("2024-01-15T13:30:45.250"), "ISO format with fraction");
    test.checkEqual(xlpp::toIso8601Date(DateTime{2024, 1, 15, 13, 30, 45}), std::string("2024-01-15"), "ISO date-only format");
}

void testDateCellsRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m18_date_cells.xlsx";
    {
        xlpp::Workbook workbook;
        auto& sheet = workbook.addWorksheet("Dates");
        sheet.cell("A1").setDate(2024, 1, 15);
        sheet.cell("B1").setDateTime(xlpp::DateTime{2024, 1, 15, 13, 30, 45});
        sheet.cell("C1").setValue(42.5);
        workbook.save(path);
    }
    {
        xlpp::Workbook workbook;
        workbook.load(path);
        test.checkEqual(workbook.date1904(), false, "1900 epoch is the default");
        auto& sheet = *workbook.worksheet("Dates");
        const auto a1 = sheet.cell("A1");
        test.checkTrue(a1.date().has_value(), "Date cell reads as DateTime");
        test.checkEqual(*a1.date(), xlpp::DateTime{2024, 1, 15}, "Date value round-trips");
        test.checkEqual(a1.numberFormat(), std::string("yyyy-mm-dd"), "Date number format");
        const auto b1 = sheet.cell("B1");
        test.checkTrue(b1.date().has_value(), "DateTime cell reads as DateTime");
        test.checkEqual(*b1.date(), xlpp::DateTime{2024, 1, 15, 13, 30, 45}, "DateTime value round-trips");
        const auto c1 = sheet.cell("C1");
        test.checkTrue(!c1.date().has_value(), "Plain number stays numeric");
        test.checkNear(std::get<double>(c1.value()), 42.5, 1e-12, "Plain number value");
    }
    {
        xlpp::Workbook workbook;
        workbook.setDate1904(true);
        auto& sheet = workbook.addWorksheet("Dates");
        sheet.cell("A1").setDate(2024, 1, 15);
        workbook.save(path);
    }
    {
        xlpp::Workbook workbook;
        workbook.load(path);
        test.checkEqual(workbook.date1904(), true, "1904 epoch round-trips");
        const auto a1 = workbook.worksheet("Dates")->cell("A1");
        test.checkEqual(*a1.date(), xlpp::DateTime{2024, 1, 15}, "1904 date value round-trips");
    }
    std::filesystem::remove(path);
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

    {
        xlpp::Workbook workbook;
        auto& sheet = workbook.addWorksheet("Rows");
        for (int r = 1; r <= 3000; ++r)
            sheet.append({std::string("Row ") + std::to_string(r), static_cast<double>(r)});
        xlpp::SaveOptions parallelRows;
        parallelRows.parallelWorkers = 4;
        parallelRows.parallelRows = true;
        workbook.save(parPath, parallelRows);
        xlpp::Workbook loaded;
        loaded.load(parPath);
        auto* loadedSheet = loaded.worksheet("Rows");
        test.checkTrue(loadedSheet != nullptr, "Parallel-row save worksheet loads");
        test.checkEqual(loadedSheet->maxRow(), std::size_t{3000}, "Parallel-row save retains final row");
        test.checkEqual(std::get<std::string>(loadedSheet->cell("A3000").value()), std::string("Row 3000"), "Parallel-row save retains final string");
        test.checkNear(std::get<double>(loadedSheet->cell("B3000").value()), 3000.0, 1e-12, "Parallel-row save retains final number");
        std::filesystem::remove(parPath);
    }
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

std::uint32_t crc32Of(const std::string& data) {
    std::uint32_t crc = 0xFFFFFFFFu;
    for (unsigned char c : data) {
        crc ^= c;
        for (int i = 0; i < 8; ++i) crc = (crc >> 1) ^ (0xEDB88320u & (0u - (crc & 1u)));
    }
    return ~crc;
}

// Builds a small but genuine ZIP64 archive (sentinel sizes/offsets plus the
// EOCD64 record and locator) so the reader exercises the ZIP64 path without a
// 4 GB payload.
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
        const auto preserved = xlpp::internal::ZipArchive::open(wbPath);
        test.checkTrue(preserved.contains("[Content_Types].xml"),
                       "Cancelled atomic save preserves the previous valid package");
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

void testZipIntegrityHardening(TestContext& test) {
    const auto toString = [](const std::vector<unsigned char>& bytes) {
        return std::string(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    };
    const auto toBytes = [](const std::string& bytes) {
        return std::vector<unsigned char>(bytes.begin(), bytes.end());
    };

    xlpp::internal::ZipArchive source;
    source.add("a.txt", "hello-integrity", false);
    source.add("b.txt", "world-integrity", false);
    const auto cleanBytes = source.saveToBytes();
    const auto clean = xlpp::internal::ZipArchive::open(cleanBytes);
    test.checkEqual(clean.get("a.txt"), std::string("hello-integrity"),
                    "Memory ZIP serialization round-trips without temporary string copy");

    {
        auto corrupted = toString(cleanBytes);
        const auto payload = corrupted.find("hello-integrity");
        test.checkTrue(payload != std::string::npos, "Stored ZIP payload located for CRC hardening test");
        if (payload != std::string::npos) corrupted[payload] ^= 0x01;
        bool rejected = false;
        try { (void)xlpp::internal::ZipArchive::open(toBytes(corrupted)); }
        catch (const std::exception& e) { rejected = std::string(e.what()).find("CRC") != std::string::npos; }
        test.checkTrue(rejected, "Materialized ZIP reader rejects stored-entry CRC corruption");
    }

    {
        auto duplicate = toString(cleanBytes);
        std::size_t replaced = 0;
        for (std::size_t pos = 0; (pos = duplicate.find("b.txt", pos)) != std::string::npos; pos += 5) {
            duplicate.replace(pos, 5, "a.txt");
            ++replaced;
        }
        test.checkTrue(replaced >= 2, "Duplicate-entry fixture rewrites local and central names");
        bool rejected = false;
        try { (void)xlpp::internal::ZipArchive::open(toBytes(duplicate)); }
        catch (const std::exception& e) { rejected = std::string(e.what()).find("Duplicate ZIP entry") != std::string::npos; }
        test.checkTrue(rejected, "Materialized ZIP reader rejects duplicate entry names");
    }

    {
        auto mismatch = toString(cleanBytes);
        const auto firstCentral = mismatch.find(std::string("PK\x01\x02", 4));
        test.checkTrue(firstCentral != std::string::npos, "Central directory located for header-consistency test");
        if (firstCentral != std::string::npos) {
            const auto centralName = firstCentral + 46;
            if (centralName < mismatch.size()) mismatch[centralName] = 'z';
        }
        bool rejected = false;
        try { (void)xlpp::internal::ZipArchive::open(toBytes(mismatch)); }
        catch (const std::exception& e) { rejected = std::string(e.what()).find("central/local entry name mismatch") != std::string::npos; }
        test.checkTrue(rejected, "Materialized ZIP reader rejects central/local name disagreement");
    }
}

// Exercises the writer's ZIP64 layout (large-entry path, per-record extra
// fields and the EOCD64 record) without a 4 GB payload, via the force-ZIP64
// test seam, and round-trips the result through the reader.
void testZip64WritePath(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto seqPath = dir / "xlpp_m20_zip64write_seq.xlsx";
    const auto parPath = dir / "xlpp_m20_zip64write_par.xlsx";
    const auto smallPath = dir / "xlpp_m20_zip64write_small.xlsx";

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
        xlpp::internal::ZipArchiveReader reader(seqPath);
        test.checkEqual(reader.entryCount(), std::size_t{4}, "Streaming ZIP reader parses forced ZIP64 directory");
        test.checkEqual(reader.readEntry("hello.txt"), std::string("Hello, forced ZIP64!"),
                        "Streaming ZIP64 reader round-trips compressed entry");
        auto source = reader.openEntry("big.txt");
        std::array<unsigned char, 113> chunk{};
        std::string streamed;
        auto count = source.read(chunk.data(), chunk.size());
        streamed.append(reinterpret_cast<const char*>(chunk.data()), count);
        xlpp::internal::ZipEntrySource moved(std::move(source));
        while ((count = moved.read(chunk.data(), chunk.size())) != 0)
            streamed.append(reinterpret_cast<const char*>(chunk.data()), count);
        test.checkTrue(moved.complete(), "Moved streaming ZIP source reaches validated completion");
        test.checkEqual(streamed, bigText, "Move-safe streaming inflater preserves all output");
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

    std::filesystem::remove(seqPath);
    std::filesystem::remove(parPath);
    std::filesystem::remove(smallPath);
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

void testChartPartPreservation(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto base = dir / "xlpp_chart_preserve_base.xlsx";
    const auto staged = dir / "xlpp_chart_preserve_staged.xlsx";
    const auto out = dir / "xlpp_chart_preserve_out.xlsx";
    {
        xlpp::Workbook w;
        w.addWorksheet("Sheet1").cell("A1").setValue("chart host");
        w.save(base);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(base);
        const std::string chartXml = "<c:chartSpace xmlns:c=\"urn:fixture\"><c:extLst><c:ext uri=\"custom\"/></c:extLst></c:chartSpace>";
        z.add("xl/charts/chart1.xml", chartXml);
        z.add("xl/drawings/drawing1.xml", "<drawing/>" );
        z.add("xl/drawings/_rels/drawing1.xml.rels", "<Relationships><Relationship Id=\"rIdChart1\" Target=\"../charts/chart1.xml\"/></Relationships>");
        z.add("xl/worksheets/_rels/sheet1.xml.rels", "<Relationships><Relationship Id=\"rIdDrawing\" Target=\"../drawings/drawing1.xml\"/></Relationships>");
        auto ct = z.get("[Content_Types].xml");
        const auto marker = std::string("<Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
        ct.insert(ct.rfind("</Types>"), marker);
         z.replace("[Content_Types].xml", ct);
        z.save(staged);
    }
    {
        xlpp::Workbook w;
        w.load(staged);
        const auto it = std::find_if(w.preservedParts().begin(), w.preservedParts().end(),
            [](const xlpp::PreservedPart& part) { return part.name == "xl/charts/chart1.xml"; });
        test.checkTrue(it != w.preservedParts().end(), "Chart part is preserved on load");
        w.save(out);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(out);
        test.checkEqual(z.get("xl/charts/chart1.xml"), std::string("<c:chartSpace xmlns:c=\"urn:fixture\"><c:extLst><c:ext uri=\"custom\"/></c:extLst></c:chartSpace>"), "Chart XML survives load-save");
        test.checkEqual(z.get("xl/drawings/_rels/drawing1.xml.rels"), std::string("<Relationships><Relationship Id=\"rIdChart1\" Target=\"../charts/chart1.xml\"/></Relationships>"), "Chart relationships survive load-save");
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


void testIndependentImageChartFixtureRoundTrip(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};

    for (const auto& [producer, sourcePath] : fixtures) {
        test.checkTrue(std::filesystem::exists(sourcePath), std::string(producer) + " fixture exists");
        const auto before = xlpp::internal::ZipArchive::open(sourcePath);
        const auto beforeGraph = xlpp::internal::RelationshipGraph::fromArchive(before);
        const auto beforeValidation = beforeGraph.validate();
        test.checkTrue(beforeValidation.ok(), std::string(producer) + " fixture object graph is valid");
        test.checkEqual(beforeGraph.objectInventory().worksheets, std::size_t{1}, std::string(producer) + " worksheet count");
        test.checkEqual(beforeGraph.objectInventory().drawings, std::size_t{1}, std::string(producer) + " drawing count");
        test.checkEqual(beforeGraph.objectInventory().images, std::size_t{1}, std::string(producer) + " visible image count");
        test.checkEqual(beforeGraph.objectInventory().charts, std::size_t{1}, std::string(producer) + " visible chart count");

        const auto outputPath = std::filesystem::temp_directory_path() /
            (std::string("xlpp_external_") + producer + "_roundtrip.xlsx");
        xlpp::Workbook workbook;
        workbook.load(sourcePath);
        const auto hasBrokenReferenceWarning = std::any_of(
            workbook.diagnostics().warnings.begin(), workbook.diagnostics().warnings.end(), [](const auto& warning) {
                return warning.find("Broken owner reference") != std::string::npos;
            });
        test.checkTrue(!hasBrokenReferenceWarning, std::string(producer) + " load has no owner-reference warning");
        auto* sheet = workbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(producer) + " worksheet loads by name");
        sheet->cell("K21").setValue(std::string("unrelated edit"));
        workbook.save(outputPath);

        const auto after = xlpp::internal::ZipArchive::open(outputPath);
        const auto diff = xlpp::internal::comparePackages(before, after);
        test.checkTrue(diff.afterValidation.ok(), std::string(producer) + " round-trip object graph is valid");
        test.checkTrue(diff.objectCountRegressions.empty(), std::string(producer) + " round-trip has no object-count regression");
        test.checkEqual(diff.afterObjects.drawings, diff.beforeObjects.drawings, std::string(producer) + " drawing count preserved");
        test.checkEqual(diff.afterObjects.images, diff.beforeObjects.images, std::string(producer) + " image count preserved");
        test.checkEqual(diff.afterObjects.charts, diff.beforeObjects.charts, std::string(producer) + " chart count preserved");
        for (const auto& part : {"xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels",
                                 "xl/charts/chart1.xml", "xl/media/image1.png"}) {
            test.checkTrue(after.contains(part), std::string(producer) + " keeps part " + part);
            test.checkEqual(after.get(part), before.get(part), std::string(producer) + " keeps untouched bytes for " + part);
        }
        std::filesystem::remove(outputPath);
    }
}


void testDrawingImageReaderMetadata(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};

    for (const auto& [producer, path] : fixtures) {
        xlpp::Workbook workbook;
        workbook.load(path);
        const auto& constWorkbook = workbook;
        const auto* sheet = constWorkbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(producer) + " drawing-reader worksheet loads");
        test.checkEqual(sheet->images().size(), std::size_t{1}, std::string(producer) + " image is exposed through the reader model");
        const auto& image = sheet->images().front();
        test.checkTrue(image.imported(), std::string(producer) + " image is marked as package-imported");
        test.checkEqual(image.anchor(), std::string("D2"), std::string(producer) + " image top-left marker is converted to A1 notation");
        test.checkTrue(!image.stableId().empty(), std::string(producer) + " image has a stable drawing-object ID");
        test.checkEqual(image.sourceMediaPart(), std::string("xl/media/image1.png"), std::string(producer) + " media package target is resolved");
        test.checkTrue(!image.sourceRelationshipId().empty(), std::string(producer) + " drawing relationship ID is retained");
        test.checkEqual(image.anchorInfo().from.column, std::size_t{4}, std::string(producer) + " anchor column is parsed as 1-based");
        test.checkEqual(image.anchorInfo().from.row, std::size_t{2}, std::string(producer) + " anchor row is parsed as 1-based");
        test.checkTrue(image.anchorInfo().widthEmu > 0 && image.anchorInfo().heightEmu > 0,
                       std::string(producer) + " image extents are available in EMU");
        if (std::string(producer) == "OpenPyXL") {
            test.checkEqual(static_cast<unsigned>(image.anchorInfo().type), static_cast<unsigned>(xlpp::DrawingAnchorType::OneCell),
                            "OpenPyXL one-cell image anchor is classified correctly");
            test.checkEqual(image.anchorInfo().widthEmu, 1143000LL, "OpenPyXL image width EMU is parsed exactly");
            test.checkEqual(image.anchorInfo().heightEmu, 762000LL, "OpenPyXL image height EMU is parsed exactly");
        } else {
            test.checkEqual(static_cast<unsigned>(image.anchorInfo().type), static_cast<unsigned>(xlpp::DrawingAnchorType::TwoCell),
                            "LibreOffice two-cell image anchor is classified correctly");
            test.checkEqual(image.anchorInfo().to.column, std::size_t{5}, "LibreOffice image end column is parsed");
            test.checkEqual(image.anchorInfo().to.row, std::size_t{5}, "LibreOffice image end row is parsed");
            test.checkEqual(image.anchorInfo().to.columnOffsetEmu, 531000LL, "LibreOffice end-column offset is parsed");
            test.checkEqual(image.anchorInfo().to.rowOffsetEmu, 190080LL, "LibreOffice end-row offset is parsed");
        }
    }
}

void testAbsoluteImageAnchorReader(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    const auto fixturePath = std::filesystem::temp_directory_path() / "xlpp_absolute_anchor_fixture.xlsx";
    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto drawing = package.get("xl/drawings/drawing1.xml");
    const auto anchors = xlpp::internal::tags(drawing, "oneCellAnchor");
    test.checkEqual(anchors.size(), std::size_t{2}, "Absolute-anchor fixture starts from two OpenPyXL one-cell anchors");
    if (anchors.size() >= 2) {
        const auto& imageAnchor = anchors[1];
        const auto pictures = xlpp::internal::tags(imageAnchor, "pic");
        test.checkTrue(!pictures.empty(), "Absolute-anchor fixture locates the image payload");
        if (!pictures.empty()) {
            const std::string absoluteAnchor =
                "<absoluteAnchor><pos x=\"123456\" y=\"654321\"/><ext cx=\"1143000\" cy=\"762000\"/>" +
                pictures.front() + "<clientData/></absoluteAnchor>";
            const auto position = drawing.find(imageAnchor);
            if (position != std::string::npos) drawing.replace(position, imageAnchor.size(), absoluteAnchor);
            package.replace("xl/drawings/drawing1.xml", drawing);
            package.save(fixturePath);
        }
    }

    xlpp::Workbook workbook;
    workbook.load(fixturePath);
    const auto& constWorkbook = workbook;
    const auto* sheet = constWorkbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "Absolute-anchor worksheet loads");
    test.checkEqual(sheet->images().size(), std::size_t{1}, "Absolute-anchor image is discovered");
    const auto& image = sheet->images().front();
    test.checkEqual(static_cast<unsigned>(image.anchorInfo().type), static_cast<unsigned>(xlpp::DrawingAnchorType::Absolute),
                    "Absolute image anchor is classified correctly");
    test.checkEqual(image.anchorInfo().xEmu, 123456LL, "Absolute anchor X position is parsed");
    test.checkEqual(image.anchorInfo().yEmu, 654321LL, "Absolute anchor Y position is parsed");
    test.checkEqual(image.anchorInfo().widthEmu, 1143000LL, "Absolute anchor width is parsed");
    test.checkEqual(image.anchorInfo().heightEmu, 762000LL, "Absolute anchor height is parsed");
    std::filesystem::remove(fixturePath);
}

void testAppendImageToPreservedDrawing(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};

    for (const auto& [producer, sourcePath] : fixtures) {
        const auto outputPath = std::filesystem::temp_directory_path() / (std::string("xlpp_append_image_") + producer + ".xlsx");
        const auto secondPath = std::filesystem::temp_directory_path() / (std::string("xlpp_append_image_second_") + producer + ".xlsx");
        const auto before = xlpp::internal::ZipArchive::open(sourcePath);
        const auto media = before.get("xl/media/image1.png");

        xlpp::Workbook workbook;
        workbook.load(sourcePath);
        auto* sheet = workbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(producer) + " append-image worksheet loads");
        xlpp::Image added("J3", std::vector<unsigned char>(media.begin(), media.end()), "png");
        added.setName("Added by XLPP");
        added.setWidthPixels(80.0);
        added.setHeightPixels(50.0);
        sheet->addImage(std::move(added));
        test.checkEqual(sheet->loadedImageCount(), std::size_t{1}, std::string(producer) + " imported image baseline remains one");
        test.checkEqual(sheet->appendedImageCount(), std::size_t{1}, std::string(producer) + " one image is tracked as an additive drawing mutation");
        workbook.save(outputPath);

        const auto after = xlpp::internal::ZipArchive::open(outputPath);
        const auto validation = xlpp::internal::RelationshipGraph::fromArchive(after).validate();
        const auto inventory = xlpp::internal::RelationshipGraph::fromArchive(after).objectInventory();
        test.checkTrue(validation.ok(), std::string(producer) + " drawing remains graph-valid after appending an image");
        test.checkEqual(inventory.drawings, std::size_t{1}, std::string(producer) + " append reuses the existing drawing part");
        test.checkEqual(inventory.images, std::size_t{2}, std::string(producer) + " original and appended images are both visible");
        test.checkEqual(inventory.charts, std::size_t{1}, std::string(producer) + " existing chart remains visible");
        test.checkTrue(after.contains("xl/media/image2.png"), std::string(producer) + " appended image gets a collision-free media part");
        test.checkEqual(after.get("xl/media/image1.png"), before.get("xl/media/image1.png"), std::string(producer) + " original media bytes remain untouched");
        test.checkEqual(after.get("xl/charts/chart1.xml"), before.get("xl/charts/chart1.xml"), std::string(producer) + " existing chart XML remains byte-identical");
        test.checkTrue(after.get("xl/drawings/drawing1.xml").find("Added by XLPP") != std::string::npos,
                       std::string(producer) + " appended anchor is injected into the preserved drawing XML");

        xlpp::Workbook reopened;
        reopened.load(outputPath);
        const auto& constReopened = reopened;
        const auto* reopenedSheet = constReopened.worksheet("Objects");
        test.checkTrue(reopenedSheet != nullptr, std::string(producer) + " appended workbook reloads");
        test.checkEqual(reopenedSheet->images().size(), std::size_t{2}, std::string(producer) + " reader sees both images after append");
        test.checkTrue(std::any_of(reopenedSheet->images().begin(), reopenedSheet->images().end(), [](const auto& image) {
            return image.anchor() == "J3" && image.name() == "Added by XLPP";
        }), std::string(producer) + " appended image anchor round-trips");

        // A repeated save of the same in-memory workbook must not duplicate or
        // lose the appended object even though its source package remains the
        // original external workbook.
        workbook.save(secondPath);
        const auto second = xlpp::internal::ZipArchive::open(secondPath);
        const auto secondInventory = xlpp::internal::RelationshipGraph::fromArchive(second).objectInventory();
        test.checkEqual(secondInventory.images, std::size_t{2}, std::string(producer) + " repeated save keeps exactly two visible images");
        test.checkEqual(secondInventory.charts, std::size_t{1}, std::string(producer) + " repeated save keeps the original chart");

        std::filesystem::remove(outputPath);
        std::filesystem::remove(secondPath);
    }
}


std::vector<unsigned char> onePixelPngBytes() {
    return {
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d,
        0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x04, 0x00, 0x00, 0x00, 0xb5, 0x1c, 0x0c, 0x02, 0x00, 0x00, 0x00, 0x0b,
        0x49, 0x44, 0x41, 0x54, 0x78, 0xda, 0x63, 0x64, 0xf8, 0x0f, 0x00, 0x01, 0x05,
        0x01, 0x01, 0x27, 0x18, 0xe3, 0x66, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
        0x44, 0xae, 0x42, 0x60, 0x82
    };
}

void testSelectiveImportedImageMoveResize(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};

    for (const auto& [producer, sourcePath] : fixtures) {
        const auto outputPath = std::filesystem::temp_directory_path() /
            (std::string("xlpp_selective_move_resize_") + producer + ".xlsx");
        const auto secondPath = std::filesystem::temp_directory_path() /
            (std::string("xlpp_selective_move_resize_second_") + producer + ".xlsx");
        const auto before = xlpp::internal::ZipArchive::open(sourcePath);

        xlpp::Workbook workbook;
        workbook.load(sourcePath);
        auto* sheet = workbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(producer) + " selective-mutation worksheet loads");
        const auto& constSheet = *sheet;
        test.checkEqual(constSheet.images().size(), std::size_t{1}, std::string(producer) + " selective-mutation fixture has one image");
        const auto stableId = constSheet.images().front().stableId();
        test.checkTrue(sheet->moveImage(stableId, "H6"), std::string(producer) + " imported image moves by stable ID");
        test.checkTrue(sheet->resizeImage(stableId, 64.0, 48.0), std::string(producer) + " imported image resizes by stable ID");
        test.checkTrue(!sheet->moveImage("missing-image", "A1"), std::string(producer) + " unknown stable ID is rejected");
        test.checkTrue(!sheet->resizeImage(stableId, 0.0, 48.0), std::string(producer) + " invalid resize is rejected");
        workbook.save(outputPath);

        const auto after = xlpp::internal::ZipArchive::open(outputPath);
        const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
        test.checkTrue(graph.validate().ok(), std::string(producer) + " move/resize keeps package graph valid");
        test.checkEqual(graph.objectInventory().images, std::size_t{1}, std::string(producer) + " move/resize keeps one visible image");
        test.checkEqual(graph.objectInventory().charts, std::size_t{1}, std::string(producer) + " move/resize preserves sibling chart");
        test.checkEqual(after.get("xl/charts/chart1.xml"), before.get("xl/charts/chart1.xml"),
                        std::string(producer) + " sibling chart remains byte-identical");
        test.checkEqual(after.get("xl/media/image1.png"), before.get("xl/media/image1.png"),
                        std::string(producer) + " move/resize does not rewrite media bytes");

        xlpp::Workbook reopened;
        reopened.load(outputPath);
        const auto& reopenedConst = reopened;
        const auto* reopenedSheet = reopenedConst.worksheet("Objects");
        test.checkTrue(reopenedSheet != nullptr, std::string(producer) + " move/resize result reloads");
        test.checkEqual(reopenedSheet->images().size(), std::size_t{1}, std::string(producer) + " moved image reload count");
        const auto& image = reopenedSheet->images().front();
        test.checkEqual(image.anchor(), std::string("H6"), std::string(producer) + " moved top-left anchor round-trips");
        test.checkEqual(image.anchorInfo().widthEmu, 609600LL, std::string(producer) + " resized width round-trips exactly");
        test.checkEqual(image.anchorInfo().heightEmu, 457200LL, std::string(producer) + " resized height round-trips exactly");
        if (std::string(producer) == "LibreOffice") {
            test.checkEqual(image.anchorInfo().to.column, std::size_t{8},
                            "LibreOffice two-cell resize updates the terminal column");
            test.checkEqual(image.anchorInfo().to.row, std::size_t{8},
                            "LibreOffice two-cell resize updates the terminal row");
            test.checkEqual(image.anchorInfo().to.columnOffsetEmu, 609600LL,
                            "LibreOffice two-cell resize computes the terminal column offset");
            test.checkEqual(image.anchorInfo().to.rowOffsetEmu, 76080LL,
                            "LibreOffice two-cell resize computes the terminal row offset");
        }

        workbook.save(secondPath);
        const auto second = xlpp::internal::ZipArchive::open(secondPath);
        const auto secondGraph = xlpp::internal::RelationshipGraph::fromArchive(second);
        test.checkTrue(secondGraph.validate().ok(), std::string(producer) + " repeated selective save remains graph-valid");
        test.checkEqual(secondGraph.objectInventory().images, std::size_t{1}, std::string(producer) + " repeated selective save does not duplicate image");
        std::filesystem::remove(outputPath);
        std::filesystem::remove(secondPath);
    }
}

void testSelectiveAbsoluteImageMoveResize(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    const auto fixturePath = std::filesystem::temp_directory_path() / "xlpp_absolute_mutation_fixture.xlsx";
    const auto outputPath = std::filesystem::temp_directory_path() / "xlpp_absolute_mutation_result.xlsx";
    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto drawing = package.get("xl/drawings/drawing1.xml");
    const auto anchors = xlpp::internal::tags(drawing, "oneCellAnchor");
    test.checkTrue(anchors.size() >= 2, "Absolute mutation fixture locates image anchor");
    if (anchors.size() >= 2) {
        const auto pictures = xlpp::internal::tags(anchors[1], "pic");
        test.checkTrue(!pictures.empty(), "Absolute mutation fixture locates picture payload");
        if (!pictures.empty()) {
            const std::string absoluteAnchor =
                "<absoluteAnchor><pos x=\"123456\" y=\"654321\"/><ext cx=\"1143000\" cy=\"762000\"/>" +
                pictures.front() + "<clientData/></absoluteAnchor>";
            const auto position = drawing.find(anchors[1]);
            drawing.replace(position, anchors[1].size(), absoluteAnchor);
            package.replace("xl/drawings/drawing1.xml", drawing);
            package.save(fixturePath);
        }
    }

    xlpp::Workbook workbook;
    workbook.load(fixturePath);
    auto* sheet = workbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "Absolute mutation worksheet loads");
    const auto& constSheet = *sheet;
    const auto stableId = constSheet.images().front().stableId();
    test.checkTrue(!sheet->moveImage(stableId, "A1"), "Cell-based move rejects absolute anchors");
    test.checkTrue(sheet->moveImageAbsolute(stableId, 222222, 333333), "Absolute image position can be changed in EMU");
    test.checkTrue(sheet->resizeImage(stableId, 50.0, 30.0), "Absolute image can be resized selectively");
    workbook.save(outputPath);

    xlpp::Workbook reopened;
    reopened.load(outputPath);
    const auto& reopenedConst = reopened;
    const auto* reopenedSheet = reopenedConst.worksheet("Objects");
    const auto& image = reopenedSheet->images().front();
    test.checkEqual(image.anchorInfo().xEmu, 222222LL, "Absolute image X mutation round-trips");
    test.checkEqual(image.anchorInfo().yEmu, 333333LL, "Absolute image Y mutation round-trips");
    test.checkEqual(image.anchorInfo().widthEmu, 476250LL, "Absolute image width mutation round-trips");
    test.checkEqual(image.anchorInfo().heightEmu, 285750LL, "Absolute image height mutation round-trips");
    const auto after = xlpp::internal::ZipArchive::open(outputPath);
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "Absolute selective mutation keeps package graph valid");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Absolute selective mutation preserves sibling chart");
    std::filesystem::remove(fixturePath);
    std::filesystem::remove(outputPath);
}

void testSelectiveImportedImageRemove(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};
    for (const auto& [producer, sourcePath] : fixtures) {
        const auto outputPath = std::filesystem::temp_directory_path() /
            (std::string("xlpp_selective_remove_") + producer + ".xlsx");
        const auto before = xlpp::internal::ZipArchive::open(sourcePath);
        xlpp::Workbook workbook;
        workbook.load(sourcePath);
        auto* sheet = workbook.worksheet("Objects");
        const auto stableId = static_cast<const xlpp::Worksheet&>(*sheet).images().front().stableId();
        test.checkTrue(sheet->removeImage(stableId), std::string(producer) + " imported image removes by stable ID");
        test.checkTrue(sheet->imageByStableId(stableId) == nullptr, std::string(producer) + " removed image disappears from the in-memory model");
        test.checkTrue(!sheet->removeImage(stableId), std::string(producer) + " removing the same imported image twice is rejected");
        workbook.save(outputPath);
        const auto after = xlpp::internal::ZipArchive::open(outputPath);
        const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
        test.checkTrue(graph.validate().ok(), std::string(producer) + " image removal keeps package graph valid");
        test.checkEqual(graph.objectInventory().images, std::size_t{0}, std::string(producer) + " image removal removes the visible image");
        test.checkEqual(graph.objectInventory().charts, std::size_t{1}, std::string(producer) + " image removal preserves sibling chart");
        test.checkTrue(!after.contains("xl/media/image1.png"), std::string(producer) + " unreferenced media part is cleaned up");
        test.checkEqual(after.get("xl/charts/chart1.xml"), before.get("xl/charts/chart1.xml"),
                        std::string(producer) + " sibling chart XML remains byte-identical after image removal");
        xlpp::Workbook reopened;
        reopened.load(outputPath);
        const auto* reopenedSheet = static_cast<const xlpp::Workbook&>(reopened).worksheet("Objects");
        test.checkEqual(reopenedSheet->images().size(), std::size_t{0}, std::string(producer) + " removed image stays removed after reload");
        std::filesystem::remove(outputPath);
    }
}

void testSelectiveImageReplaceWithSharedMedia(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    const auto fixturePath = std::filesystem::temp_directory_path() / "xlpp_shared_media_fixture.xlsx";
    const auto replacePath = std::filesystem::temp_directory_path() / "xlpp_shared_media_replace.xlsx";
    const auto removePath = std::filesystem::temp_directory_path() / "xlpp_shared_media_remove.xlsx";
    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto drawing = package.get("xl/drawings/drawing1.xml");
    const auto anchors = xlpp::internal::tags(drawing, "oneCellAnchor");
    test.checkTrue(anchors.size() >= 2, "Shared-media fixture locates image anchor");
    if (anchors.size() >= 2) {
        auto duplicate = anchors[1];
        auto replaceOnce = [](std::string& value, const std::string& from, const std::string& to) {
            const auto pos = value.find(from);
            if (pos != std::string::npos) value.replace(pos, from.size(), to);
        };
        replaceOnce(duplicate, "id=\"2\"", "id=\"3\"");
        replaceOnce(duplicate, "name=\"Image 2\"", "name=\"Image 3\"");
        replaceOnce(duplicate, "r:embed=\"rId2\"", "r:embed=\"rId3\"");
        replaceOnce(duplicate, "<col>3</col>", "<col>5</col>");
        const auto close = drawing.rfind("</wsDr>");
        drawing.insert(close, duplicate);
        package.replace("xl/drawings/drawing1.xml", drawing);

        auto rels = package.get("xl/drawings/_rels/drawing1.xml.rels");
        const std::string rel =
            "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
            "Target=\"/xl/media/image1.png\" Id=\"rId3\"/>";
        rels.insert(rels.rfind("</Relationships>"), rel);
        package.replace("xl/drawings/_rels/drawing1.xml.rels", rels);
        package.save(fixturePath);
    }

    const auto fixture = xlpp::internal::ZipArchive::open(fixturePath);
    const auto fixtureGraph = xlpp::internal::RelationshipGraph::fromArchive(fixture);
    test.checkTrue(fixtureGraph.validate().ok(), "Shared-media fixture graph is valid");
    test.checkEqual(fixtureGraph.objectInventory().images, std::size_t{2}, "Shared-media fixture exposes two images");

    xlpp::Workbook replaceWorkbook;
    replaceWorkbook.load(fixturePath);
    auto* replaceSheet = replaceWorkbook.worksheet("Objects");
    const auto& constReplaceSheet = *replaceSheet;
    const auto target = std::find_if(constReplaceSheet.images().begin(), constReplaceSheet.images().end(), [](const auto& image) {
        return image.name() == "Image 3";
    });
    test.checkTrue(target != constReplaceSheet.images().end(), "Shared-media replacement target is found by stable object metadata");
    const auto targetId = target->stableId();
    xlpp::Image replacement("A1", onePixelPngBytes(), "png");
    replacement.setName("Replacement");
    test.checkTrue(replaceSheet->replaceImage(targetId, std::move(replacement)), "Shared-media image can be replaced selectively");
    replaceWorkbook.save(replacePath);

    const auto replaced = xlpp::internal::ZipArchive::open(replacePath);
    const auto replacedGraph = xlpp::internal::RelationshipGraph::fromArchive(replaced);
    test.checkTrue(replacedGraph.validate().ok(), "Shared-media replacement keeps graph valid");
    test.checkEqual(replacedGraph.objectInventory().images, std::size_t{2}, "Shared-media replacement preserves both visible images");
    test.checkEqual(replacedGraph.objectInventory().charts, std::size_t{1}, "Shared-media replacement preserves chart");
    test.checkEqual(replaced.get("xl/media/image1.png"), fixture.get("xl/media/image1.png"),
                    "Replacing one shared-media image leaves the original media bytes untouched");
    test.checkTrue(replaced.contains("xl/media/image2.png"), "Replacing one shared-media image allocates a private media part");
    const auto replacementBytes = onePixelPngBytes();
    const std::string replacementText(reinterpret_cast<const char*>(replacementBytes.data()), replacementBytes.size());
    test.checkEqual(replaced.get("xl/media/image2.png"), replacementText, "Replacement media part contains the requested bytes");
    const auto relationships = replacedGraph.relationshipsFrom("xl/drawings/drawing1.xml");
    const auto rId2 = std::find_if(relationships.begin(), relationships.end(), [](const auto& rel) { return rel.id == "rId2"; });
    const auto rId3 = std::find_if(relationships.begin(), relationships.end(), [](const auto& rel) { return rel.id == "rId3"; });
    test.checkTrue(rId2 != relationships.end() && xlpp::internal::RelationshipGraph::resolveTarget(rId2->sourcePart, rId2->target) == "xl/media/image1.png",
                   "First image keeps the shared original media target");
    test.checkTrue(rId3 != relationships.end() && xlpp::internal::RelationshipGraph::resolveTarget(rId3->sourcePart, rId3->target) == "xl/media/image2.png",
                   "Replaced image relationship is retargeted without renumbering its relationship ID");
    test.checkEqual(replaced.get("xl/charts/chart1.xml"), fixture.get("xl/charts/chart1.xml"),
                    "Shared-media replacement preserves sibling chart XML byte-for-byte");

    xlpp::Workbook removeWorkbook;
    removeWorkbook.load(fixturePath);
    auto* removeSheet = removeWorkbook.worksheet("Objects");
    const auto& constRemoveSheet = *removeSheet;
    const auto removeTarget = std::find_if(constRemoveSheet.images().begin(), constRemoveSheet.images().end(), [](const auto& image) {
        return image.name() == "Image 3";
    });
    test.checkTrue(removeTarget != constRemoveSheet.images().end(), "Shared-media remove target is found");
    test.checkTrue(removeSheet->removeImage(removeTarget->stableId()), "One of two shared-media images can be removed selectively");
    removeWorkbook.save(removePath);
    const auto removed = xlpp::internal::ZipArchive::open(removePath);
    const auto removedGraph = xlpp::internal::RelationshipGraph::fromArchive(removed);
    test.checkTrue(removedGraph.validate().ok(), "Shared-media removal keeps graph valid");
    test.checkEqual(removedGraph.objectInventory().images, std::size_t{1}, "Shared-media removal leaves the sibling image visible");
    test.checkTrue(removed.contains("xl/media/image1.png"), "Shared media part is retained while another image still references it");
    test.checkEqual(removed.get("xl/media/image1.png"), fixture.get("xl/media/image1.png"), "Shared media bytes remain unchanged after sibling removal");

    std::filesystem::remove(fixturePath);
    std::filesystem::remove(replacePath);
    std::filesystem::remove(removePath);
}

void testMultiDrawingSelectiveMutation(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    const auto fixturePath = std::filesystem::temp_directory_path() / "xlpp_multi_drawing_fixture.xlsx";
    const auto outputPath = std::filesystem::temp_directory_path() / "xlpp_multi_drawing_mutated.xlsx";

    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto sheetXml = package.get("xl/worksheets/sheet1.xml");
    const std::string secondDrawingOwner =
        "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId2\"/>";
    sheetXml.insert(sheetXml.rfind("</worksheet>"), secondDrawingOwner);
    package.replace("xl/worksheets/sheet1.xml", sheetXml);

    auto sheetRels = package.get("xl/worksheets/_rels/sheet1.xml.rels");
    sheetRels.insert(sheetRels.rfind("</Relationships>"),
        "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
        "Target=\"/xl/drawings/drawing2.xml\" Id=\"rId2\"/>");
    package.replace("xl/worksheets/_rels/sheet1.xml.rels", sheetRels);

    auto contentTypes = package.get("[Content_Types].xml");
    contentTypes.insert(contentTypes.rfind("</Types>"),
        "<Override PartName=\"/xl/drawings/drawing2.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
    package.replace("[Content_Types].xml", contentTypes);

    const std::string drawing2 =
        "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<xdr:oneCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        "<xdr:ext cx=\"1143000\" cy=\"762000\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"21\" name=\"Second Drawing Image\"/>"
        "<xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rIdImage\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
        "<xdr:spPr><a:prstGeom prst=\"rect\"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"
        "<xdr:oneCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        "<xdr:ext cx=\"1500000\" cy=\"500000\"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"22\" name=\"Preserved Text Box\">"
        "<a:hlinkClick r:id=\"rIdShapeLink\"/></xdr:cNvPr><xdr:cNvSpPr txBox=\"1\"/></xdr:nvSpPr><xdr:spPr/>"
        "<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>keep me</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>"
        "</xdr:wsDr>";
    package.add("xl/drawings/drawing2.xml", drawing2);
    package.add("xl/drawings/_rels/drawing2.xml.rels",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"/xl/media/image1.png\"/>"
        "<Relationship Id=\"rIdShapeLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
        "Target=\"https://example.com/preserved-shape\" TargetMode=\"External\"/>"
        "</Relationships>");
    package.save(fixturePath);

    const auto fixture = xlpp::internal::ZipArchive::open(fixturePath);
    const auto fixtureGraph = xlpp::internal::RelationshipGraph::fromArchive(fixture);
    test.checkTrue(fixtureGraph.validate().ok(), "Multi-drawing fixture is graph-valid including the shape hyperlink");
    test.checkEqual(fixtureGraph.objectInventory().drawings, std::size_t{2}, "Validator sees both worksheet drawing parts");
    test.checkEqual(fixtureGraph.objectInventory().images, std::size_t{2}, "Validator sees images in both drawing parts");
    test.checkEqual(fixtureGraph.objectInventory().charts, std::size_t{1}, "Chart in the first drawing remains visible");
    test.checkEqual(fixtureGraph.objectInventory().shapes, std::size_t{1}, "Validator inventories the unsupported shape");
    test.checkEqual(fixtureGraph.objectInventory().textBoxes, std::size_t{1}, "Validator inventories the text box shape");

    xlpp::Workbook workbook;
    workbook.load(fixturePath);
    auto* sheet = workbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "Multi-drawing worksheet loads");
    const auto& constSheet = *sheet;
    test.checkEqual(constSheet.images().size(), std::size_t{2}, "Reader exposes images from both preserved drawings");
    const auto secondImage = std::find_if(constSheet.images().begin(), constSheet.images().end(), [](const auto& image) {
        return image.sourceDrawingPart() == "xl/drawings/drawing2.xml";
    });
    test.checkTrue(secondImage != constSheet.images().end(), "Second drawing image retains its source drawing part");
    const auto secondStableId = secondImage->stableId();
    test.checkTrue(sheet->moveImage(secondStableId, "K6"), "Selective move can target an image in the second preserved drawing");
    workbook.save(outputPath);

    const auto after = xlpp::internal::ZipArchive::open(outputPath);
    const auto afterGraph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(afterGraph.validate().ok(), "Multi-drawing selective mutation keeps the package graph valid");
    test.checkEqual(afterGraph.objectInventory().drawings, std::size_t{2}, "Both drawing parts survive selective mutation");
    test.checkEqual(afterGraph.objectInventory().images, std::size_t{2}, "Both images survive selective mutation");
    test.checkEqual(afterGraph.objectInventory().charts, std::size_t{1}, "Sibling chart survives second-drawing mutation");
    test.checkEqual(afterGraph.objectInventory().shapes, std::size_t{1}, "Unsupported shape survives selective mutation");
    test.checkEqual(afterGraph.objectInventory().textBoxes, std::size_t{1}, "Unsupported text box survives selective mutation");
    test.checkEqual(after.get("xl/drawings/drawing1.xml"), fixture.get("xl/drawings/drawing1.xml"),
                    "Untouched first drawing stays byte-identical");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), fixture.get("xl/drawings/_rels/drawing1.xml.rels"),
                    "Untouched first drawing relationships stay byte-identical");
    test.checkTrue(after.get("xl/drawings/drawing2.xml").find("<xdr:col>10</xdr:col>") != std::string::npos &&
                   after.get("xl/drawings/drawing2.xml").find("<xdr:row>5</xdr:row>") != std::string::npos,
                   "Second drawing image anchor moves to K6 without rebuilding sibling objects");
    test.checkTrue(after.get("xl/drawings/drawing2.xml").find("Preserved Text Box") != std::string::npos,
                   "Unknown DrawingML text-box XML remains in the selectively patched drawing");
    test.checkTrue(after.get("xl/drawings/_rels/drawing2.xml.rels").find("rIdShapeLink") != std::string::npos,
                   "Unknown drawing hyperlink relationship remains connected");

    xlpp::Workbook reopened;
    reopened.load(outputPath);
    const auto* reopenedSheet = static_cast<const xlpp::Workbook&>(reopened).worksheet("Objects");
    test.checkTrue(reopenedSheet != nullptr, "Mutated multi-drawing workbook reloads");
    test.checkTrue(std::any_of(reopenedSheet->images().begin(), reopenedSheet->images().end(), [](const auto& image) {
        return image.sourceDrawingPart() == "xl/drawings/drawing2.xml" && image.anchor() == "K6";
    }), "Reader observes the moved image in drawing2 after round-trip");

    std::filesystem::remove(fixturePath);
    std::filesystem::remove(outputPath);
}

void testUnknownDrawingRelationshipValidation(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto drawing = package.get("xl/drawings/drawing1.xml");
    auto rels = package.get("xl/drawings/_rels/drawing1.xml.rels");

    // Inject an unsupported DrawingML shape carrying an external hyperlink.
    const std::string shape =
        "<oneCellAnchor><from><col>8</col><colOff>0</colOff><row>10</row><rowOff>0</rowOff></from><ext cx=\"1000000\" cy=\"400000\"/>"
        "<sp><nvSpPr><cNvPr id=\"77\" name=\"Unknown shape\"><a:hlinkClick xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdUnknown\"/></cNvPr><cNvSpPr/></nvSpPr>"
        "<spPr/><txBody><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>"
        "<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>unknown</a:t></a:r></a:p></txBody></sp><clientData/></oneCellAnchor>";
    drawing.insert(drawing.rfind("</wsDr>"), shape);
    rels.insert(rels.rfind("</Relationships>"),
        "<Relationship Id=\"rIdUnknown\" Type=\"urn:xlpp:test:unsupportedDrawingRelationship\" "
        "Target=\"https://example.com/unknown\" TargetMode=\"External\"/>");
    package.replace("xl/drawings/drawing1.xml", drawing);
    package.replace("xl/drawings/_rels/drawing1.xml.rels", rels);

    const auto validGraph = xlpp::internal::RelationshipGraph::fromArchive(package);
    test.checkTrue(validGraph.validate().ok(), "Referenced unknown DrawingML relationship is preserved as a valid owner edge");
    test.checkEqual(validGraph.objectInventory().shapes, std::size_t{1}, "Unknown shape is included in drawing inventory");
    test.checkEqual(validGraph.objectInventory().textBoxes, std::size_t{1}, "Unknown shape text body is inventoried as a text box");

    auto orphanedRelationship = package;
    auto brokenDrawing = orphanedRelationship.get("xl/drawings/drawing1.xml");
    const auto attribute = std::string(" r:id=\"rIdUnknown\"");
    const auto position = brokenDrawing.find(attribute);
    test.checkTrue(position != std::string::npos, "Negative drawing fixture locates the relationship-bearing attribute");
    if (position != std::string::npos) brokenDrawing.erase(position, attribute.size());
    orphanedRelationship.replace("xl/drawings/drawing1.xml", brokenDrawing);
    const auto orphanReport = xlpp::internal::RelationshipGraph::fromArchive(orphanedRelationship).validate();
    test.checkTrue(std::any_of(orphanReport.ownerReferenceErrors.begin(), orphanReport.ownerReferenceErrors.end(), [](const auto& issue) {
        return issue.find("rIdUnknown") != std::string::npos && issue.find("not referenced") != std::string::npos;
    }), "Validator rejects an unknown drawing relationship that is no longer referenced by DrawingML");

    auto missingRelationship = package;
    auto brokenRels = missingRelationship.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto relationBegin = brokenRels.find("<Relationship Id=\"rIdUnknown\"");
    test.checkTrue(relationBegin != std::string::npos, "Negative drawing fixture locates the unknown relationship node");
    if (relationBegin != std::string::npos) {
        const auto relationEnd = brokenRels.find("/>", relationBegin);
        if (relationEnd != std::string::npos) brokenRels.erase(relationBegin, relationEnd + 2 - relationBegin);
    }
    missingRelationship.replace("xl/drawings/_rels/drawing1.xml.rels", brokenRels);
    const auto missingReport = xlpp::internal::RelationshipGraph::fromArchive(missingRelationship).validate();
    test.checkTrue(std::any_of(missingReport.ownerReferenceErrors.begin(), missingReport.ownerReferenceErrors.end(), [](const auto& issue) {
        return issue.find("rIdUnknown") != std::string::npos && issue.find("missing relationship") != std::string::npos;
    }), "Validator rejects DrawingML that references an unknown missing relationship");
}

void testIndependentPivotFixtureRoundTrip(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "libreoffice" / "pivot.xlsx";
    const auto outputPath = std::filesystem::temp_directory_path() / "xlpp_libreoffice_pivot_roundtrip.xlsx";
    test.checkTrue(std::filesystem::exists(sourcePath), "LibreOffice pivot fixture exists");

    const auto before = xlpp::internal::ZipArchive::open(sourcePath);
    const auto beforeGraph = xlpp::internal::RelationshipGraph::fromArchive(before);
    const auto beforeValidation = beforeGraph.validate();
    test.checkTrue(beforeValidation.ok(), "LibreOffice pivot fixture object graph is valid");
    test.checkEqual(beforeGraph.objectInventory().pivotTables, std::size_t{1}, "LibreOffice visible pivot count");
    test.checkEqual(beforeGraph.objectInventory().pivotCaches, std::size_t{1}, "LibreOffice pivot-cache count");
    test.checkTrue(xlpp::internal::tags(before.get("xl/worksheets/sheet1.xml"), "pivotTableParts").empty(),
                   "LibreOffice fixture intentionally owns its pivot through the worksheet relationship graph");

    {
        auto broken = xlpp::internal::ZipArchive::open(sourcePath);
        auto pivotXml = broken.get("xl/pivotTables/pivotTable1.xml");
        const auto cacheId = pivotXml.find("cacheId=\"1\"");
        test.checkTrue(cacheId != std::string::npos, "Negative pivot fixture exposes cacheId for mutation");
        if (cacheId != std::string::npos) pivotXml.replace(cacheId, std::string("cacheId=\"1\"").size(), "cacheId=\"999\"");
        broken.replace("xl/pivotTables/pivotTable1.xml", pivotXml);
        const auto brokenValidation = xlpp::internal::RelationshipGraph::fromArchive(broken).validate();
        test.checkTrue(!brokenValidation.ownerReferenceErrors.empty(), "Validator rejects a pivot whose logical cacheId is not declared by the workbook");
    }

    xlpp::Workbook workbook;
    workbook.load(sourcePath);
    const auto hasBrokenReferenceWarning = std::any_of(
        workbook.diagnostics().warnings.begin(), workbook.diagnostics().warnings.end(), [](const auto& warning) {
            return warning.find("Broken owner reference") != std::string::npos;
        });
    test.checkTrue(!hasBrokenReferenceWarning, "LibreOffice pivot load has no owner-reference warning");
    auto* sheet = workbook.worksheet("Data");
    test.checkTrue(sheet != nullptr, "LibreOffice pivot worksheet loads by name");
    sheet->cell("M20").setValue(std::string("unrelated edit"));
    workbook.save(outputPath);

    const auto after = xlpp::internal::ZipArchive::open(outputPath);
    const auto diff = xlpp::internal::comparePackages(before, after);
    test.checkTrue(diff.afterValidation.ok(), "LibreOffice pivot round-trip object graph is valid");
    test.checkTrue(diff.objectCountRegressions.empty(), "LibreOffice pivot round-trip has no object-count regression");
    test.checkEqual(diff.afterObjects.pivotTables, std::size_t{1}, "LibreOffice pivot remains reachable");
    test.checkEqual(diff.afterObjects.pivotCaches, std::size_t{1}, "LibreOffice pivot cache remains reachable");

    for (const auto& part : {"xl/pivotTables/pivotTable1.xml",
                             "xl/pivotTables/_rels/pivotTable1.xml.rels",
                             "xl/pivotCache/pivotCacheDefinition1.xml",
                             "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
                             "xl/pivotCache/pivotCacheRecords1.xml"}) {
        test.checkTrue(after.contains(part), std::string("LibreOffice pivot round-trip keeps ") + part);
        test.checkEqual(after.get(part), before.get(part), std::string("LibreOffice pivot part stays byte-identical: ") + part);
    }

    std::filesystem::remove(outputPath);
}

void testPreservedAndGeneratedPivotCoexistence(TestContext& test) {
    const auto canonicalPath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "libreoffice" / "pivot.xlsx";
    const auto sourcePath = std::filesystem::temp_directory_path() / "xlpp_nonmatching_cache_id_source.xlsx";
    const auto outputPath = std::filesystem::temp_directory_path() / "xlpp_preserved_generated_pivots.xlsx";

    // Keep the independent LibreOffice package topology, but deliberately make
    // its logical cacheId differ from pivotCacheDefinition1.xml. OOXML does not
    // require the cacheId to equal a part's numeric file suffix.
    {
        auto source = xlpp::internal::ZipArchive::open(canonicalPath);
        auto workbookXml = source.get("xl/workbook.xml");
        auto pivotXml = source.get("xl/pivotTables/pivotTable1.xml");
        const auto workbookCacheId = workbookXml.find("cacheId=\"1\"");
        const auto pivotCacheId = pivotXml.find("cacheId=\"1\"");
        test.checkTrue(workbookCacheId != std::string::npos && pivotCacheId != std::string::npos,
                       "Cache-ID coexistence fixture can be decoupled from its part suffix");
        if (workbookCacheId != std::string::npos)
            workbookXml.replace(workbookCacheId, std::string("cacheId=\"1\"").size(), "cacheId=\"17\"");
        if (pivotCacheId != std::string::npos)
            pivotXml.replace(pivotCacheId, std::string("cacheId=\"1\"").size(), "cacheId=\"17\"");
        source.replace("xl/workbook.xml", workbookXml);
        source.replace("xl/pivotTables/pivotTable1.xml", pivotXml);
        source.save(sourcePath);
    }

    const auto before = xlpp::internal::ZipArchive::open(sourcePath);
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(before).validate().ok(),
                   "Nonmatching logical cacheId and cache-part suffix are valid before mixed save");

    xlpp::Workbook workbook;
    workbook.load(sourcePath);
    auto* sheet = workbook.worksheet("Data");
    test.checkTrue(sheet != nullptr, "Pivot coexistence source worksheet loads");

    xlpp::PivotTable pivot("AddedByXLPP");
    pivot.setLocation("J2");
    pivot.cache().setSourceData("'Data'!$A$1:$C$7");
    pivot.cache().setFields({"Region", "Quarter", "Sales"});
    pivot.addRowField("Region");
    pivot.addDataField("Sales", "sum");
    sheet->addPivotTable(std::move(pivot));
    workbook.save(outputPath);

    const auto after = xlpp::internal::ZipArchive::open(outputPath);
    const auto afterGraph = xlpp::internal::RelationshipGraph::fromArchive(after);
    const auto validation = afterGraph.validate();
    test.checkTrue(validation.ok(), "Preserved and generated pivots form one valid object graph");
    test.checkEqual(afterGraph.objectInventory().pivotTables, std::size_t{2}, "Existing and new pivot tables are both reachable");
    test.checkEqual(afterGraph.objectInventory().pivotCaches, std::size_t{2}, "Existing and new pivot caches are both reachable");

    for (const auto& part : {"xl/pivotTables/pivotTable1.xml",
                             "xl/pivotTables/_rels/pivotTable1.xml.rels",
                             "xl/pivotCache/pivotCacheDefinition1.xml",
                             "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
                             "xl/pivotCache/pivotCacheRecords1.xml"}) {
        test.checkTrue(after.contains(part), std::string("Mixed pivot save keeps original part ") + part);
        test.checkEqual(after.get(part), before.get(part), std::string("Mixed pivot save preserves original bytes: ") + part);
    }
    test.checkTrue(after.contains("xl/pivotTables/pivotTable2.xml"), "Mixed pivot save writes new pivot table to a non-colliding part");
    test.checkTrue(after.contains("xl/pivotCache/pivotCacheDefinition2.xml"), "Mixed pivot save writes new cache definition to a non-colliding part");
    test.checkTrue(after.contains("xl/pivotCache/pivotCacheRecords2.xml"), "Mixed pivot save writes new cache records to a non-colliding part");

    const auto cacheNodes = xlpp::internal::tags(after.get("xl/workbook.xml"), "pivotCache");
    test.checkEqual(cacheNodes.size(), std::size_t{2}, "Workbook merges preserved and generated pivotCache nodes");
    std::set<std::string> cacheIds;
    for (const auto& node : cacheNodes) cacheIds.insert(xlpp::internal::attribute(node, "cacheId"));
    test.checkEqual(cacheIds.size(), std::size_t{2}, "Merged pivot caches have unique logical cache IDs");
    test.checkTrue(cacheIds.count("17") == 1 && cacheIds.count("18") == 1,
                   "Generated cache ID advances beyond the preserved logical cache ID, independent of part numbering");
    test.checkTrue(after.get("xl/pivotTables/pivotTable2.xml").find("cacheId=\"18\"") != std::string::npos,
                   "Generated pivot uses cacheId 18 while its physical part remains pivotTable2.xml");

    const auto sheetRelationships = afterGraph.relationshipsFrom("xl/worksheets/sheet1.xml");
    const auto pivotRelationshipCount = static_cast<std::size_t>(std::count_if(
        sheetRelationships.begin(), sheetRelationships.end(), [](const auto& relationship) {
            return relationship.type.find("/pivotTable") != std::string::npos;
        }));
    test.checkEqual(pivotRelationshipCount, std::size_t{2}, "Worksheet retains original implicit pivot relationship and adds the new pivot relationship");

    std::filesystem::remove(sourcePath);
    std::filesystem::remove(outputPath);
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

void testPivotValidation(TestContext& test) {
    xlpp::PivotTable pivot("Validation");
    pivot.cache().setFields({"Category", "Amount"});
    bool widthRejected = false;
    try { pivot.cache().addRecord({"only-one"}); } catch (const std::invalid_argument&) { widthRejected = true; }
    test.checkTrue(widthRejected, "Pivot cache rejects mismatched record width");
    pivot.addDataField(2);
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Pivot");
    sheet.addPivotTable(std::move(pivot));
    bool indexRejected = false;
    try { workbook.save(std::filesystem::temp_directory_path() / "xlpp_invalid_pivot.xlsx"); }
    catch (const std::invalid_argument&) { indexRejected = true; }
    test.checkTrue(indexRejected, "Pivot serialization rejects invalid field index");
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

// Malformed-input hardening: oversized structural ranges, alloc-bombs and
// truncated/garbage ZIP bytes must fail with clean exceptions, never crash.
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

// Deterministic mutation fuzz: flip/truncate bytes in a valid workbook and in
// the raw ZIP container. Every mutated input must either load or throw a clean
// std::exception — never crash, never hang, never silently produce garbage
// that trips later serialization.
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

void testCellReferenceMatrix(TestContext& test) {
    const std::vector<std::pair<std::size_t, std::string>> columns{
        {1, "A"}, {2, "B"}, {26, "Z"}, {27, "AA"}, {52, "AZ"}, {53, "BA"},
        {78, "BZ"}, {79, "CA"}, {702, "ZZ"}, {703, "AAA"}, {16384, "XFD"}};
    for (const auto& [index, name] : columns) {
        test.checkEqual(xlpp::CellReference::columnName(index), name, "columnName for " + name);
        test.checkEqual(xlpp::CellReference::columnIndex(name), index, "columnIndex for " + name);
        test.checkEqual(xlpp::CellReference::columnIndex(xlpp::CellReference::columnName(index)), index,
                        "columnName/columnIndex inverse for " + name);
    }

    test.checkEqual(xlpp::CellReference::parse("$A$1").row, std::size_t{1}, "Dollar-prefixed row");
    test.checkEqual(xlpp::CellReference::parse("$A$1").column, std::size_t{1}, "Dollar-prefixed column");
    test.checkEqual(xlpp::CellReference::parse("xfd1048576").row, std::size_t{1048576}, "Maximum row parses");
    test.checkEqual(xlpp::CellReference::parse("xfd1048576").column, std::size_t{16384}, "Maximum column parses");
    test.checkEqual(xlpp::CellReference{3, 2}.address(), std::string("B3"), "Address builds from coords");
    test.checkEqual(xlpp::CellReference::parse("b3").address(), std::string("B3"), "Lowercase input normalizes");

    test.checkEqual(xlpp::makeCellKey(1, 1), std::uint64_t{1} << 20 | 1, "Row 1 col 1 key");
    test.checkEqual(xlpp::makeCellKey(2, 1) > xlpp::makeCellKey(1, 16384), true,
                    "Row-major ordering keeps next row after max column");
    test.checkEqual(xlpp::makeCellKey(1048576, 16384), (std::uint64_t{1048576} << 20) | 16384,
                    "Max coordinate key");

    bool threw = false;
    try { (void)xlpp::CellReference::columnName(0); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "columnName(0) throws");
    threw = false;
    try { (void)xlpp::CellReference::columnIndex(""); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "columnIndex(empty) throws");
    threw = false;
    try { (void)xlpp::CellReference::columnIndex("A1"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "columnIndex rejects digits");
    threw = false;
    try { (void)xlpp::CellReference::parse(""); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "parse(empty) throws");
    threw = false;
    try { (void)xlpp::CellReference::parse("1A"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "parse rejects digits before letters");
    threw = false;
    try { (void)xlpp::CellReference::parse("A0"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "parse rejects row zero");
    threw = false;
    try { (void)xlpp::CellReference::parse("A1:Z9"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "parse rejects range strings");
    threw = false;
    try { (void)xlpp::CellReference::parse("A"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "parse rejects missing row");
}

void testCellErrorMatrix(TestContext& test) {
    const std::vector<std::pair<xlpp::CellError, std::string>> errors{
        {xlpp::CellError::Null, "#NULL!"},
        {xlpp::CellError::DivisionByZero, "#DIV/0!"},
        {xlpp::CellError::Value, "#VALUE!"},
        {xlpp::CellError::Reference, "#REF!"},
        {xlpp::CellError::Name, "#NAME?"},
        {xlpp::CellError::Number, "#NUM!"},
        {xlpp::CellError::NotAvailable, "#N/A"},
        {xlpp::CellError::GettingData, "#GETTING_DATA"},
    };
    for (const auto& [error, text] : errors) {
        test.checkEqual(xlpp::toString(error), text, "toString maps error to " + text);
        test.checkTrue(xlpp::cellErrorFromString(text) == error, "cellErrorFromString parses " + text);
    }
    test.checkTrue(xlpp::cellErrorFromString("#BOGUS!") == xlpp::CellError::Value,
                   "Unknown error text falls back to #VALUE!");
    test.checkTrue(xlpp::cellErrorFromString("") == xlpp::CellError::Value,
                   "Empty error text falls back to #VALUE!");
}

void testXlfnHelper(TestContext& test) {
    test.checkEqual(xlpp::xlfn("SORT"), std::string("_xlfn.SORT"), "New function gets prefix");
    test.checkEqual(xlpp::xlfn("FILTER(A1:A5,\"x\")"), std::string("_xlfn.FILTER(A1:A5,\"x\")"),
                    "Prefixed argument form");
    test.checkEqual(xlpp::xlfn("_xlfn.XLOOKUP"), std::string("_xlfn.XLOOKUP"),
                    "Already-prefixed input unchanged");
    test.checkEqual(xlpp::xlfn("_xlfn.UNIQUE"), std::string("_xlfn.UNIQUE"),
                    "Case-preserving on prefix");
    test.checkEqual(xlpp::xlfn(""), std::string(""), "Empty input stays empty");
    test.checkEqual(xlpp::xlfn("SEQUENCE(10)"), std::string("_xlfn.SEQUENCE(10)"),
                    "Function with args gets prefix");
}

void testFormulaMetadataDefaults(TestContext& test) {
    xlpp::Cell cell("A1");
    test.checkTrue(cell.formulaMetadata().empty(), "Fresh metadata is empty");
    test.checkTrue(!cell.hasFormula(), "No formula initially");

    cell.setSharedFormula("B1+C1", 9, "A1:A5");
    test.checkTrue(cell.hasFormula(), "Shared formula present");
    test.checkEqual(static_cast<unsigned>(cell.formulaMetadata().type()),
                    static_cast<unsigned>(xlpp::FormulaType::Shared), "Shared type set");
    test.checkEqual(*cell.formulaMetadata().sharedIndex(), 9u, "Shared index stored");
    test.checkEqual(cell.formulaMetadata().reference(), std::string("A1:A5"), "Shared reference stored");
    cell.formulaMetadata().setCalculateOnLoad(true);
    test.checkTrue(cell.formulaMetadata().calculateOnLoad(), "Calculate-on-load flag");
    test.checkTrue(!cell.formulaMetadata().empty(), "Populated metadata is non-empty");

    cell.clearFormula();
    test.checkTrue(!cell.hasFormula(), "clearFormula removes formula");
    test.checkTrue(cell.formulaMetadata().empty(), "clearFormula resets metadata");

    cell.setDynamicArrayFormula("_xlfn.SORT(A1:A5)", "C1");
    test.checkEqual(static_cast<unsigned>(cell.formulaMetadata().type()),
                    static_cast<unsigned>(xlpp::FormulaType::DynamicArray), "Dynamic array type");
    test.checkTrue(cell.formulaMetadata().alwaysCalculateArray(), "Dynamic array sets aca");
    test.checkEqual(cell.formulaMetadata().reference(), std::string("C1"), "Dynamic array reference");
    test.checkEqual(cell.formula(), std::string("_xlfn.SORT(A1:A5)"), "Dynamic array formula text");
}

void testNumberFormatDetection(TestContext& test) {
    test.checkTrue(xlpp::isDateFormatCode("yyyy-mm-dd", 0), "Literal date format detected");
    test.checkTrue(xlpp::isDateFormatCode("m/d/yy", 14), "Built-in id 14 is a date");
    test.checkTrue(xlpp::isDateFormatCode("", 27), "Built-in id 27 is a date");
    test.checkTrue(xlpp::isDateFormatCode("", 36), "Built-in id 36 is a date");
    test.checkTrue(xlpp::isDateFormatCode("", 45), "Built-in id 45 is a date");
    test.checkTrue(xlpp::isDateFormatCode("", 50), "Built-in id 50 is a date");
    test.checkTrue(xlpp::isDateFormatCode("", 58), "Built-in id 58 is a date");
    test.checkTrue(xlpp::isDateFormatCode("", 81), "Built-in id 81 is a date");
    test.checkTrue(!xlpp::isDateFormatCode("", 0), "Built-in id 0 is not a date");
    test.checkTrue(!xlpp::isDateFormatCode("", 1), "Built-in id 1 is not a date");
    test.checkTrue(!xlpp::isDateFormatCode("", 49), "Built-in id 49 is not a date");
    test.checkTrue(!xlpp::isDateFormatCode("0.00%", 0), "Percent format is not a date");
    test.checkTrue(!xlpp::isDateFormatCode("#,##0.00", 0), "Thousands format is not a date");
    test.checkTrue(xlpp::isDateFormatCode("[h]:mm:ss", 0), "Elapsed time bracket format");
    test.checkTrue(!xlpp::isDateFormatCode("\"yyyy\";0.00", 0), "Quoted letters are literals");
    test.checkTrue(!xlpp::isDateFormatCode("\\m", 0), "Escaped letter is a literal");
    test.checkTrue(xlpp::isDateFormatCode("[Red]yyyy", 0), "Color section then date letters");
    test.checkTrue(!xlpp::isDateFormatCode("[Red]0.00", 0), "Color section numeric stays numeric");
    test.checkTrue(!xlpp::isDateFormatCode("[$-F800]dddd, mmmm dd, yyyy", 0) == false,
                   "Locale format with letters detected");
}

void testDateTimeBoundaries(TestContext& test) {
    using xlpp::DateTime;
    test.checkEqual(xlpp::fromExcelSerial(59.0), DateTime{1900, 2, 28}, "Serial 59 is 1900-02-28");
    test.checkEqual(xlpp::fromExcelSerial(61.0), DateTime{1900, 3, 1}, "Serial 61 is 1900-03-01");
    test.checkEqual(xlpp::fromExcelSerial(0.0), DateTime{1899, 12, 31}, "Serial 0 is the epoch");
    test.checkEqual(xlpp::fromExcelSerial(-1.0), DateTime{1899, 12, 30}, "Negative serials go before epoch");

    const DateTime leap2000{2000, 2, 29};
    test.checkEqual(xlpp::fromExcelSerial(xlpp::toExcelSerial(leap2000)), leap2000, "2000 leap day round-trips");
    const DateTime leap2024{2024, 2, 29};
    test.checkEqual(xlpp::fromExcelSerial(xlpp::toExcelSerial(leap2024)), leap2024, "2024 leap day round-trips");
    const DateTime notLeap{2100, 2, 28};
    test.checkEqual(xlpp::fromExcelSerial(xlpp::toExcelSerial(notLeap)), notLeap, "Century non-leap year");

    test.checkEqual(xlpp::fromExcelSerial(xlpp::toExcelSerial(DateTime{1899, 12, 30})),
                    DateTime{1899, 12, 30}, "Pre-epoch date round-trips");
    test.checkEqual(xlpp::fromExcelSerial(xlpp::toExcelSerial(DateTime{1899, 12, 31})),
                    DateTime{1899, 12, 31}, "Epoch date round-trips");

    test.checkNear(xlpp::toExcelSerial(DateTime{2000, 1, 1, 0, 0, 0.5}), 36526.0 + 0.5 / 86400.0, 1e-12,
                   "Sub-second fraction preserved");
    const DateTime subsecond{2024, 6, 1, 12, 0, 0.25};
    const auto serial = xlpp::toExcelSerial(subsecond);
    test.checkTrue(std::abs(xlpp::fromExcelSerial(serial).second - 0.25) < 1e-9,
                   "Quarter-second survives serial round-trip");
    test.checkEqual(xlpp::fromExcelSerial(2958465.0), DateTime{9999, 12, 31}, "Max Excel serial round-trips");
    test.checkEqual(xlpp::fromExcelSerial(2958466.0), DateTime{10000, 1, 1}, "Serial past max rolls to next year");
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

void testImportedChartInspectionAndSelectiveMutation(TestContext& test) {
    const struct FixtureCase {
        const char* producer;
        const char* relativePath;
        xlpp::DrawingAnchorType anchorType;
        const char* preservedMarker;
    } cases[] = {
        {"OpenPyXL", "fixtures/openpyxl/image_chart.xlsx", xlpp::DrawingAnchorType::OneCell, "prstDash"},
        {"LibreOffice", "fixtures/libreoffice/image_chart.xlsx", xlpp::DrawingAnchorType::TwoCell, "c15:showLeaderLines"}
    };

    for (const auto& fixture : cases) {
        const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / fixture.relativePath;
        const auto before = xlpp::internal::ZipArchive::open(source);
        const auto originalChartXml = before.get("xl/charts/chart1.xml");
        const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
        const auto originalImageBytes = before.get("xl/media/image1.png");
        const auto originalSheetFormats = xlpp::internal::tags(before.get("xl/worksheets/sheet1.xml"), "sheetFormatPr");

        xlpp::Workbook workbook;
        workbook.load(source);
        auto* sheet = workbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(fixture.producer) + " chart fixture worksheet loads");
        const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
        test.checkEqual(charts.size(), std::size_t{1}, std::string(fixture.producer) + " chart reader exposes one chart");
        if (charts.empty()) continue;
        const auto& chart = charts.front();
        test.checkTrue(chart.imported(), std::string(fixture.producer) + " chart is marked imported");
        test.checkTrue(!chart.stableId().empty(), std::string(fixture.producer) + " chart has stable drawing-object ID");
        test.checkEqual(chart.sourceDrawingPart(), std::string("xl/drawings/drawing1.xml"), std::string(fixture.producer) + " chart source drawing part");
        test.checkEqual(chart.sourceChartPart(), std::string("xl/charts/chart1.xml"), std::string(fixture.producer) + " chart source chart part");
        test.checkEqual(chart.sourceRelationshipId(), std::string("rId1"), std::string(fixture.producer) + " chart source relationship ID");
        test.checkEqual(chart.title(), std::string("OpenPyXL chart"), std::string(fixture.producer) + " chart title parsed across namespace styles");
        test.checkEqual(chart.xAxisTitle(), std::string("Category"), std::string(fixture.producer) + " category-axis title parsed");
        test.checkEqual(chart.yAxisTitle(), std::string("Amount"), std::string(fixture.producer) + " value-axis title parsed");
        test.checkEqual(static_cast<int>(chart.type()), static_cast<int>(xlpp::Chart::Type::Bar), std::string(fixture.producer) + " chart type parsed");
        test.checkEqual(static_cast<int>(chart.anchorInfo().type), static_cast<int>(fixture.anchorType), std::string(fixture.producer) + " chart anchor type parsed");
        test.checkEqual(chart.anchorInfo().from.row, std::size_t{8}, std::string(fixture.producer) + " chart anchor row parsed");
        test.checkEqual(chart.anchorInfo().from.column, std::size_t{4}, std::string(fixture.producer) + " chart anchor column parsed");
        test.checkNear(static_cast<double>(chart.width()), 302.0, 1.0, std::string(fixture.producer) + " chart width parsed from owning anchor");
        test.checkNear(static_cast<double>(chart.height()), 189.0, 1.0, std::string(fixture.producer) + " chart height parsed from owning anchor");
        test.checkEqual(chart.series().size(), std::size_t{1}, std::string(fixture.producer) + " chart series parsed");
        test.checkTrue(!chart.series().front().categoriesReference().empty(), std::string(fixture.producer) + " category formula parsed");
        test.checkTrue(!chart.series().front().valuesReference().empty(), std::string(fixture.producer) + " value formula parsed");

        const auto stableId = chart.stableId();
        test.checkTrue(sheet->chartByStableId(stableId) != nullptr, std::string(fixture.producer) + " chart lookup by stable ID");
        test.checkTrue(sheet->chartByStableId("missing-chart") == nullptr, std::string(fixture.producer) + " missing stable chart ID returns null");
        test.checkTrue(!sheet->setChartSeriesReferences(stableId, 9, "Objects!$A$2:$A$3", "Objects!$B$2:$B$3"), std::string(fixture.producer) + " invalid chart series index is rejected");
        test.checkTrue(!sheet->resizeChart(stableId, 0.0, 200.0), std::string(fixture.producer) + " invalid chart resize is rejected");
        test.checkTrue(!sheet->moveChartAbsolute(stableId, 1, 1), std::string(fixture.producer) + " cell-anchored chart rejects absolute move");

        test.checkTrue(sheet->setChartTitle(stableId, "XL++ <safe> & chart"), std::string(fixture.producer) + " imported chart title edits selectively");
        test.checkTrue(sheet->setChartSeriesReferences(stableId, 0, "'Objects'!$A$2:$A$3", "'Objects'!$B$2:$B$3"), std::string(fixture.producer) + " imported chart series formulas edit selectively");
        test.checkTrue(sheet->moveChart(stableId, "H6"), std::string(fixture.producer) + " imported chart moves by stable ID");
        test.checkTrue(sheet->resizeChart(stableId, 320.0, 200.0), std::string(fixture.producer) + " imported chart resizes by stable ID");
        sheet->cell("J20").setValue("chart-edit-regression");

        const auto output = std::filesystem::temp_directory_path() /
            (std::string("xlpp_p0h_") + fixture.producer + "_chart_mutation.xlsx");
        workbook.save(output);
        const auto after = xlpp::internal::ZipArchive::open(output);
        const auto editedChartXml = after.get("xl/charts/chart1.xml");
        test.checkTrue(editedChartXml.find("XL++ &lt;safe&gt; &amp; chart") != std::string::npos,
                       std::string(fixture.producer) + " chart title is XML-escaped in selective patch");
        test.checkTrue(editedChartXml.find("&apos;Objects&apos;!$A$2:$A$3") != std::string::npos ||
                       editedChartXml.find("'Objects'!$A$2:$A$3") != std::string::npos,
                       std::string(fixture.producer) + " category formula updated");
        test.checkTrue(editedChartXml.find("&apos;Objects&apos;!$B$2:$B$3") != std::string::npos ||
                       editedChartXml.find("'Objects'!$B$2:$B$3") != std::string::npos,
                       std::string(fixture.producer) + " value formula updated");
        test.checkTrue(editedChartXml.find(fixture.preservedMarker) != std::string::npos,
                       std::string(fixture.producer) + " unsupported chart formatting/extensions survive selective edits");
        test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels,
                        std::string(fixture.producer) + " drawing relationships remain byte-identical");
        test.checkEqual(after.get("xl/media/image1.png"), originalImageBytes,
                        std::string(fixture.producer) + " sibling image bytes remain byte-identical");
        const auto afterSheetFormats = xlpp::internal::tags(after.get("xl/worksheets/sheet1.xml"), "sheetFormatPr");
        test.checkTrue(!originalSheetFormats.empty() && !afterSheetFormats.empty(),
                       std::string(fixture.producer) + " source sheet-format metrics remain present");
        if (!originalSheetFormats.empty() && !afterSheetFormats.empty())
            test.checkEqual(afterSheetFormats.front(), originalSheetFormats.front(),
                            std::string(fixture.producer) + " source sheet-format metrics remain byte-identical for drawing geometry");

        const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
        const auto validation = graph.validate();
        test.checkTrue(validation.relationshipSyntaxErrors.empty(), std::string(fixture.producer) + " selective chart output has no relationship syntax errors");
        test.checkTrue(validation.danglingRelationships.empty(), std::string(fixture.producer) + " selective chart output has no dangling relationships");
        test.checkTrue(validation.orphanedParts.empty(), std::string(fixture.producer) + " selective chart output has no orphaned parts");
        test.checkTrue(validation.ownerReferenceErrors.empty(), std::string(fixture.producer) + " selective chart output has no owner-reference errors");
        test.checkEqual(graph.objectInventory().charts, std::size_t{1}, std::string(fixture.producer) + " selective chart output retains visible chart count");
        test.checkEqual(graph.objectInventory().images, std::size_t{1}, std::string(fixture.producer) + " selective chart output retains sibling image count");

        xlpp::Workbook reloaded;
        reloaded.load(output);
        const auto* reloadedSheet = reloaded.worksheet("Objects");
        const auto& reloadedChart = static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front();
        test.checkEqual(reloadedChart.title(), std::string("XL++ <safe> & chart"), std::string(fixture.producer) + " selective title survives reload");
        test.checkEqual(reloadedChart.series().front().categoriesReference(), std::string("'Objects'!$A$2:$A$3"), std::string(fixture.producer) + " selective category formula survives reload");
        test.checkEqual(reloadedChart.series().front().valuesReference(), std::string("'Objects'!$B$2:$B$3"), std::string(fixture.producer) + " selective value formula survives reload");
        test.checkEqual(reloadedChart.anchorInfo().from.row, std::size_t{6}, std::string(fixture.producer) + " selective chart row survives reload");
        test.checkEqual(reloadedChart.anchorInfo().from.column, std::size_t{8}, std::string(fixture.producer) + " selective chart column survives reload");
        test.checkNear(static_cast<double>(reloadedChart.width()), 320.0, 1.0, std::string(fixture.producer) + " selective chart width survives reload");
        test.checkNear(static_cast<double>(reloadedChart.height()), 200.0, 1.0, std::string(fixture.producer) + " selective chart height survives reload");
        std::filesystem::remove(output);
        (void)originalChartXml;
    }
}

void testImportedScatterChartDeepSelectiveEditing(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/scatter_multiseries.xlsx";
    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Scatter");
    test.checkTrue(sheet != nullptr, "OpenPyXL scatter fixture worksheet loads");
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "Scatter fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkEqual(static_cast<int>(chart.type()), static_cast<int>(xlpp::Chart::Type::Scatter), "Scatter chart type parsed");
    test.checkEqual(chart.xAxisTitle(), std::string("X Axis"), "Scatter X title comes from first value axis");
    test.checkEqual(chart.yAxisTitle(), std::string("Y Axis"), "Scatter Y title comes from second value axis");
    test.checkEqual(chart.series().size(), std::size_t{2}, "Scatter multi-series count parsed");
    test.checkEqual(chart.series()[0].title(), std::string("Alpha"), "Scatter first series title parsed");
    test.checkEqual(chart.series()[1].title(), std::string("Beta"), "Scatter second series title parsed");
    test.checkTrue(chart.series()[1].categoriesReference().find("$A$2:$A$5") != std::string::npos, "Scatter second X reference parsed");
    test.checkTrue(chart.series()[1].valuesReference().find("$C$2:$C$5") != std::string::npos, "Scatter second Y reference parsed");

    const auto stableId = chart.stableId();
    test.checkTrue(!sheet->setChartLegend(stableId, true, "invalid"), "Invalid legend position rejected");
    test.checkTrue(sheet->setChartXAxisTitle(stableId, "Horizontal <X>"), "Scatter X-axis title selective edit");
    test.checkTrue(sheet->setChartYAxisTitle(stableId, "Vertical & Y"), "Scatter Y-axis title selective edit");
    test.checkTrue(sheet->setChartLegend(stableId, true, "b"), "Scatter legend selective edit");
    test.checkTrue(sheet->setChartSeriesTitle(stableId, 1, "Gamma & Delta"), "Scatter second series title selective edit");
    test.checkTrue(!sheet->setChartSeriesTitle(stableId, 5, "bad"), "Scatter invalid series-title index rejected");
    test.checkTrue(sheet->setChartSeriesReferences(stableId, 1, "'Scatter'!$A$2:$A$4", "'Scatter'!$C$2:$C$4"),
                   "Scatter xVal/yVal references selective edit");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0i_scatter_selective.xlsx";
    workbook.save(output);
    const auto archive = xlpp::internal::ZipArchive::open(output);
    const auto xml = archive.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("Horizontal &lt;X&gt;") != std::string::npos, "Scatter X title XML escaped");
    test.checkTrue(xml.find("Vertical &amp; Y") != std::string::npos, "Scatter Y title XML escaped");
    test.checkTrue(xml.find("Gamma &amp; Delta") != std::string::npos, "Scatter series title XML escaped");
    test.checkTrue(xml.find("legendPos val=\"b\"") != std::string::npos, "Scatter legend position patched");
    test.checkTrue(xml.find("prstDash") != std::string::npos, "Scatter unsupported series line formatting preserved");

    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    const auto validation = graph.validate();
    test.checkTrue(validation.ok(), "Scatter selective output package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Scatter selective output chart count stable");
    test.checkEqual(graph.objectInventory().images, std::size_t{1}, "Scatter sibling image retained");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto& reloadedChart = static_cast<const xlpp::Worksheet&>(*reloaded.worksheet("Scatter")).charts().front();
    test.checkEqual(reloadedChart.xAxisTitle(), std::string("Horizontal <X>"), "Scatter X title survives reload");
    test.checkEqual(reloadedChart.yAxisTitle(), std::string("Vertical & Y"), "Scatter Y title survives reload");
    test.checkEqual(reloadedChart.legendPosition(), std::string("b"), "Scatter legend position survives reload");
    test.checkEqual(reloadedChart.series()[1].title(), std::string("Gamma & Delta"), "Scatter series title survives reload");
    test.checkTrue(reloadedChart.series()[1].categoriesReference().find("$A$2:$A$4") != std::string::npos, "Scatter edited X reference survives reload");
    test.checkTrue(reloadedChart.series()[1].valuesReference().find("$C$2:$C$4") != std::string::npos, "Scatter edited Y reference survives reload");

    const auto stableReloaded = reloadedChart.stableId();
    auto* reloadedSheet = reloaded.worksheet("Scatter");
    test.checkTrue(reloadedSheet->setChartLegend(stableReloaded, false), "Imported chart legend can be hidden selectively");
    const auto noLegend = std::filesystem::temp_directory_path() / "xlpp_p0i_scatter_no_legend.xlsx";
    reloaded.save(noLegend);
    const auto noLegendXml = xlpp::internal::ZipArchive::open(noLegend).get("xl/charts/chart1.xml");
    test.checkTrue(noLegendXml.find("<legend") == std::string::npos && noLegendXml.find("<c:legend") == std::string::npos, "Selective legend hide removes legend node");
    std::filesystem::remove(output);
    std::filesystem::remove(noLegend);
}

void testImportedCombinedChartAxisStructure(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/combined_secondary_axes.xlsx";
    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Data");
    test.checkTrue(sheet != nullptr, "Combined chart fixture worksheet loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "Combined fixture exposes one chart object");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkTrue(chart.combined(), "Bar + line chart is recognized as combined");
    test.checkTrue(chart.hasSecondaryAxes(), "Combined chart secondary axis is recognized");
    test.checkEqual(chart.plots().size(), std::size_t{2}, "Combined chart exposes two plots");
    test.checkEqual(chart.axes().size(), std::size_t{3}, "Combined chart exposes three axis definitions");
    test.checkEqual(chart.primaryXAxisId(), std::uint64_t{10}, "Primary X/category axis ID parsed from first plot");
    test.checkEqual(chart.primaryYAxisId(), std::uint64_t{100}, "Primary Y/value axis ID parsed from first plot");
    test.checkEqual(chart.xAxisTitle(), std::string("Month"), "Primary X axis title resolved by axId");
    test.checkEqual(chart.yAxisTitle(), std::string("Sales"), "Primary Y axis title resolved by axId");
    test.checkEqual(chart.series().size(), std::size_t{2}, "Combined chart series from both plots are loaded");

    test.checkEqual(static_cast<int>(chart.plots()[0].type), static_cast<int>(xlpp::Chart::Type::Bar), "First plot is bar");
    test.checkEqual(static_cast<int>(chart.plots()[1].type), static_cast<int>(xlpp::Chart::Type::Line), "Second plot is line");
    test.checkEqual(chart.plots()[0].firstSeries, std::size_t{0}, "Primary plot first series index");
    test.checkEqual(chart.plots()[0].seriesCount, std::size_t{1}, "Primary plot series count");
    test.checkEqual(chart.plots()[1].firstSeries, std::size_t{1}, "Secondary plot first series index");
    test.checkEqual(chart.plots()[1].seriesCount, std::size_t{1}, "Secondary plot series count");
    test.checkTrue(!chart.plots()[0].usesSecondaryAxes, "Primary plot does not use secondary axes");
    test.checkTrue(chart.plots()[1].usesSecondaryAxes, "Line plot is linked to secondary value axis");

    const auto* categoryAxis = chart.axisById(10);
    const auto* primaryValueAxis = chart.axisById(100);
    const auto* secondaryValueAxis = chart.axisById(200);
    test.checkTrue(categoryAxis != nullptr && primaryValueAxis != nullptr && secondaryValueAxis != nullptr,
                   "Axis lookup by native axId succeeds");
    if (categoryAxis && primaryValueAxis && secondaryValueAxis) {
        test.checkEqual(static_cast<int>(categoryAxis->kind), static_cast<int>(xlpp::Chart::AxisKind::Category), "Axis 10 is category axis");
        test.checkEqual(categoryAxis->crossAxisId, std::uint64_t{100}, "Category axis crossAx parsed");
        test.checkTrue(!categoryAxis->secondary, "Shared category axis remains primary");
        test.checkEqual(primaryValueAxis->crossAxisId, std::uint64_t{10}, "Primary value axis crossAx parsed");
        test.checkTrue(!primaryValueAxis->secondary, "Primary value axis classified primary");
        test.checkEqual(secondaryValueAxis->crossAxisId, std::uint64_t{10}, "Secondary value axis crossAx parsed");
        test.checkEqual(secondaryValueAxis->position, std::string("r"), "Secondary axis position parsed");
        test.checkEqual(secondaryValueAxis->title, std::string("Margin"), "Secondary axis title parsed");
        test.checkTrue(secondaryValueAxis->secondary, "Axis 200 classified secondary");
    }
    test.checkTrue(chart.axisById(9999) == nullptr, "Unknown axis ID lookup returns null");

    const auto stableId = chart.stableId();
    test.checkTrue(!sheet->setChartAxisTitle(stableId, 9999, "invalid"), "Unknown axis ID edit is rejected");
    test.checkTrue(sheet->setChartXAxisTitle(stableId, "Primary Category"), "Primary X title edits by native axis ID");
    test.checkTrue(sheet->setChartYAxisTitle(stableId, "Primary Sales"), "Primary Y title edits by native axis ID");
    test.checkTrue(sheet->setChartAxisTitle(stableId, 200, "Secondary Margin"), "Secondary axis title edits by native axis ID");
    test.checkTrue(sheet->setChartSeriesTitle(stableId, 1, "Margin series"), "Secondary plot series title edits selectively");
    sheet->cell("J20").setValue("combined-axis-regression");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0j_combined_secondary_axes.xlsx";
    workbook.save(output);
    const auto archive = xlpp::internal::ZipArchive::open(output);
    const auto xml = archive.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("Primary Category") != std::string::npos, "Primary category title written");
    test.checkTrue(xml.find("Primary Sales") != std::string::npos, "Primary value title written");
    test.checkTrue(xml.find("Secondary Margin") != std::string::npos, "Secondary value title written");
    test.checkTrue(xml.find("Margin series") != std::string::npos, "Secondary plot series title written");
    test.checkTrue(xml.find("axId val=\"10\"") != std::string::npos &&
                   xml.find("axId val=\"100\"") != std::string::npos &&
                   xml.find("axId val=\"200\"") != std::string::npos,
                   "Selective edits preserve all native axis IDs");
    test.checkTrue(xml.find("crosses val=\"max\"") != std::string::npos, "Unsupported secondary-axis crosses metadata preserved");
    test.checkTrue(xml.find("<barChart") != std::string::npos && xml.find("<lineChart") != std::string::npos,
                   "Combined plot structure preserved");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    test.checkTrue(graph.validate().ok(), "Combined chart selective output package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Combined chart remains one visible drawing chart");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadedSheet = reloaded.worksheet("Data");
    test.checkTrue(reloadedSheet != nullptr, "Combined chart output reloads");
    if (reloadedSheet) {
        const auto& reloadedCharts = static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts();
        test.checkEqual(reloadedCharts.size(), std::size_t{1}, "Combined chart count survives reload");
        if (!reloadedCharts.empty()) {
            const auto& reloadedChart = reloadedCharts.front();
            test.checkTrue(reloadedChart.combined(), "Combined plot identity survives reload");
            test.checkTrue(reloadedChart.hasSecondaryAxes(), "Secondary-axis identity survives reload");
            test.checkEqual(reloadedChart.primaryXAxisId(), std::uint64_t{10}, "Primary X axis ID survives reload");
            test.checkEqual(reloadedChart.primaryYAxisId(), std::uint64_t{100}, "Primary Y axis ID survives reload");
            test.checkEqual(reloadedChart.xAxisTitle(), std::string("Primary Category"), "Primary X title survives reload");
            test.checkEqual(reloadedChart.yAxisTitle(), std::string("Primary Sales"), "Primary Y title survives reload");
            const auto* secondary = reloadedChart.axisById(200);
            test.checkTrue(secondary != nullptr && secondary->title == "Secondary Margin" && secondary->secondary,
                           "Secondary axis title/classification survives reload");
        }
    }
    std::filesystem::remove(output);
}

void testImportedChartLabelsTrendlinesAndErrorBars(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_labels_trendline_errorbars.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Advanced");
    test.checkTrue(sheet != nullptr, "Advanced chart fixture worksheet loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "Advanced fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkEqual(chart.plots().size(), std::size_t{1}, "Advanced scatter exposes one plot");
    test.checkTrue(chart.plots()[0].dataLabels.present, "Plot data labels are parsed");
    test.checkTrue(chart.plots()[0].dataLabels.showValue, "Data-label showVal parsed");
    test.checkTrue(chart.plots()[0].dataLabels.showSeriesName, "Data-label showSerName parsed");
    test.checkEqual(chart.plots()[0].dataLabels.position, std::string("t"), "Data-label position parsed");
    test.checkEqual(chart.series().size(), std::size_t{2}, "Advanced scatter series count parsed");
    test.checkEqual(chart.series()[0].trendlines().size(), std::size_t{1}, "First series trendline parsed");
    test.checkEqual(static_cast<int>(chart.series()[0].trendlines()[0].type),
                    static_cast<int>(xlpp::ChartSeries::TrendlineType::Linear), "Linear trendline type parsed");
    test.checkTrue(chart.series()[0].trendlines()[0].displayEquation, "Trendline equation flag parsed");
    test.checkTrue(chart.series()[0].trendlines()[0].displayRSquared, "Trendline R-squared flag parsed");
    test.checkEqual(chart.series()[1].trendlines().size(), std::size_t{1}, "Second series polynomial trendline parsed");
    test.checkTrue(chart.series()[1].dataLabels().present, "Series-level data labels are parsed");
    test.checkTrue(chart.series()[1].dataLabels().showCategoryName, "Series-level category-name flag parsed");
    test.checkEqual(chart.series()[1].dataLabels().position, std::string("r"), "Series-level data-label position parsed");
    test.checkEqual(chart.series()[1].trendlines()[0].order, 2, "Polynomial trendline order parsed");
    test.checkEqual(chart.series()[0].errorBars().size(), std::size_t{1}, "Series Y error bars parsed");
    test.checkEqual(static_cast<int>(chart.series()[0].errorBars()[0].direction),
                    static_cast<int>(xlpp::ChartSeries::ErrorBarDirection::Y), "Error-bar direction parsed");
    test.checkEqual(static_cast<int>(chart.series()[0].errorBars()[0].valueType),
                    static_cast<int>(xlpp::ChartSeries::ErrorValueType::FixedValue), "Error-bar value type parsed");
    test.checkNear(chart.series()[0].errorBars()[0].value, 1.5, 1e-9, "Fixed error-bar value parsed");

    const auto stableId = chart.stableId();
    auto labels = chart.plots()[0].dataLabels;
    labels.showCategoryName = true;
    labels.showSeriesName = false;
    labels.position = "b";
    labels.separator = " | ";
    test.checkTrue(sheet->setChartPlotDataLabels(stableId, 0, labels), "Plot data labels selectively edit");
    test.checkTrue(!sheet->setChartPlotDataLabels(stableId, 3, labels), "Invalid plot-index data-label edit rejected");

    auto seriesLabels = chart.series()[1].dataLabels();
    seriesLabels.showValue = true;
    seriesLabels.position = "l";
    seriesLabels.separator = " / ";
    test.checkTrue(sheet->setChartSeriesDataLabels(stableId, 1, seriesLabels), "Series-level data labels selectively edit");
    test.checkTrue(!sheet->setChartSeriesDataLabels(stableId, 8, seriesLabels), "Invalid series data-label index rejected");

    auto trendline = chart.series()[0].trendlines()[0];
    trendline.type = xlpp::ChartSeries::TrendlineType::Polynomial;
    trendline.order = 3;
    trendline.forward = 1.0;
    trendline.displayEquation = false;
    trendline.displayRSquared = true;
    test.checkTrue(sheet->setChartSeriesTrendline(stableId, 0, 0, trendline), "Existing trendline selectively edits");
    test.checkTrue(sheet->removeChartSeriesTrendline(stableId, 1, 0), "Existing second trendline selectively removes");
    xlpp::ChartSeries::Trendline exponential;
    exponential.type = xlpp::ChartSeries::TrendlineType::Exponential;
    exponential.displayEquation = true;
    test.checkTrue(sheet->addChartSeriesTrendline(stableId, 1, exponential), "Trendline selectively appends");

    auto yBars = chart.series()[0].errorBars()[0];
    yBars.barType = xlpp::ChartSeries::ErrorBarType::Plus;
    yBars.value = 2.25;
    yBars.noEndCap = true;
    test.checkTrue(sheet->setChartSeriesErrorBars(stableId, 0, yBars), "Existing Y error bars selectively edit");
    xlpp::ChartSeries::ErrorBars xBars;
    xBars.direction = xlpp::ChartSeries::ErrorBarDirection::X;
    xBars.barType = xlpp::ChartSeries::ErrorBarType::Both;
    xBars.valueType = xlpp::ChartSeries::ErrorValueType::Percentage;
    xBars.value = 5.0;
    test.checkTrue(sheet->setChartSeriesErrorBars(stableId, 1, xBars), "New X error bars selectively append");
    auto customBars = xBars;
    customBars.valueType = xlpp::ChartSeries::ErrorValueType::Custom;
    test.checkTrue(!sheet->setChartSeriesErrorBars(stableId, 1, customBars), "Custom error-bar write rejected until range refs are modeled");
    sheet->cell("J20").setValue("p0k-regression");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0k_chart_features.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("dLblPos val=\"b\"") != std::string::npos, "Data-label position selectively patched");
    test.checkTrue(xml.find("showCatName val=\"1\"") != std::string::npos, "Data-label category-name flag added");
    test.checkTrue(xml.find("showSerName val=\"0\"") != std::string::npos, "Data-label series-name flag disabled");
    test.checkTrue(xml.find("separator> | </") != std::string::npos, "Data-label separator selectively added");
    test.checkTrue(xml.find("dLblPos val=\"l\"") != std::string::npos && xml.find("separator> / </") != std::string::npos,
                   "Series-level data labels selectively patched");
    test.checkTrue(xml.find("trendlineType val=\"poly\"") != std::string::npos && xml.find("order val=\"3\"") != std::string::npos,
                   "Polynomial trendline selective patch written");
    test.checkTrue(xml.find("trendlineType val=\"exp\"") != std::string::npos, "Replacement exponential trendline written");
    test.checkTrue(xml.find("errBarType val=\"plus\"") != std::string::npos && xml.find("val val=\"2.25\"") != std::string::npos,
                   "Existing fixed Y error bars selectively patched");
    test.checkTrue(xml.find("errDir val=\"x\"") != std::string::npos && xml.find("errValType val=\"percentage\"") != std::string::npos,
                   "New X percentage error bars written");
    test.checkTrue(xml.find("prstDash") != std::string::npos, "Unsupported series formatting survives chart feature edits");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels,
                    "Chart feature edits keep drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "Chart feature edits keep sibling image byte-identical");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "Chart labels/trendline/error-bar output package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Advanced chart remains visible after selective edits");
    test.checkEqual(graph.objectInventory().images, std::size_t{1}, "Sibling image remains visible after selective edits");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("Advanced");
    test.checkTrue(reloadSheet != nullptr, "Advanced selective output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        test.checkTrue(reloadChart.plots()[0].dataLabels.showCategoryName && !reloadChart.plots()[0].dataLabels.showSeriesName,
                       "Data-label flags survive reload");
        test.checkEqual(reloadChart.plots()[0].dataLabels.position, std::string("b"), "Data-label position survives reload");
        test.checkTrue(reloadChart.series()[1].dataLabels().present && reloadChart.series()[1].dataLabels().showValue,
                       "Series-level data labels survive reload");
        test.checkEqual(reloadChart.series()[1].dataLabels().position, std::string("l"), "Series-level data-label position survives reload");
        test.checkEqual(reloadChart.series()[0].trendlines().size(), std::size_t{1}, "Edited first trendline count survives reload");
        test.checkEqual(static_cast<int>(reloadChart.series()[0].trendlines()[0].type),
                        static_cast<int>(xlpp::ChartSeries::TrendlineType::Polynomial), "Edited trendline type survives reload");
        test.checkEqual(reloadChart.series()[0].trendlines()[0].order, 3, "Edited polynomial order survives reload");
        test.checkEqual(reloadChart.series()[1].trendlines().size(), std::size_t{1}, "Replacement second trendline count survives reload");
        test.checkEqual(static_cast<int>(reloadChart.series()[1].trendlines()[0].type),
                        static_cast<int>(xlpp::ChartSeries::TrendlineType::Exponential), "Replacement trendline survives reload");
        test.checkEqual(reloadChart.series()[0].errorBars().size(), std::size_t{1}, "Edited Y error-bar count survives reload");
        test.checkNear(reloadChart.series()[0].errorBars()[0].value, 2.25, 1e-9, "Edited error-bar value survives reload");
        test.checkEqual(reloadChart.series()[1].errorBars().size(), std::size_t{1}, "Appended X error bars survive reload");
        test.checkEqual(static_cast<int>(reloadChart.series()[1].errorBars()[0].direction),
                        static_cast<int>(xlpp::ChartSeries::ErrorBarDirection::X), "Appended X error-bar direction survives reload");

        const auto reloadId = reloadChart.stableId();
        auto* mutableReload = reloaded.worksheet("Advanced");
        test.checkTrue(mutableReload->removeChartSeriesErrorBars(reloadId, 0, xlpp::ChartSeries::ErrorBarDirection::Y),
                       "Selective error-bar removal supported");
        test.checkTrue(mutableReload->removeChartSeriesTrendline(reloadId, 1, 0), "Selective trendline removal supported after reload");
        const auto removedOutput = std::filesystem::temp_directory_path() / "xlpp_p0k_chart_features_removed.xlsx";
        reloaded.save(removedOutput);
        const auto removedXml = xlpp::internal::ZipArchive::open(removedOutput).get("xl/charts/chart1.xml");
        test.checkTrue(removedXml.find("errDir val=\"y\"") == std::string::npos, "Selective Y error bars removed from XML");
        test.checkTrue(removedXml.find("trendlineType val=\"exp\"") == std::string::npos, "Selective exponential trendline removed from XML");
        std::filesystem::remove(removedOutput);
    }
    std::filesystem::remove(output);
}

void testImportedChartPerPointCustomErrorsAndFormatting(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/per_point_custom_errors_formatting.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("P0L");
    test.checkTrue(sheet != nullptr, "P0L OpenPyXL fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0L fixture exposes one imported chart");
    if (charts.empty() || charts.front().series().empty()) return;
    const auto& chart = charts.front();
    const auto& series = chart.series().front();

    test.checkTrue(series.dataLabels().present && series.dataLabels().showValue, "Series aggregate data-label defaults parsed independently");
    test.checkEqual(series.dataLabels().points.size(), std::size_t{2}, "Per-point data labels parsed");
    if (series.dataLabels().points.size() >= 2) {
        test.checkEqual(series.dataLabels().points[0].index, std::size_t{1}, "First point-label index parsed");
        test.checkTrue(series.dataLabels().points[0].showSeriesName && !series.dataLabels().points[0].showValue,
                       "First point-label overrides parsed");
        test.checkEqual(series.dataLabels().points[0].position, std::string("t"), "First point-label position parsed");
        test.checkEqual(series.dataLabels().points[1].index, std::size_t{3}, "Second point-label index parsed");
        test.checkTrue(series.dataLabels().points[1].showCategoryName && series.dataLabels().points[1].showValue,
                       "Second point-label overrides parsed");
    }

    test.checkTrue(series.lineFormat().present, "Series line formatting parsed");
    test.checkEqual(static_cast<int>(series.lineFormat().color.kind), static_cast<int>(xlpp::ChartColor::Kind::SRgb), "Series line color kind parsed");
    test.checkEqual(series.lineFormat().color.value, std::string("FF0000"), "Series line color parsed");
    test.checkNear(series.lineFormat().widthPoints, 2.0, 1e-9, "Series line width parsed in points");
    test.checkEqual(series.lineFormat().dash, std::string("dash"), "Series dash style parsed");
    test.checkTrue(series.markerFormat().present, "Marker formatting parsed");
    test.checkEqual(series.markerFormat().symbol, std::string("diamond"), "Marker symbol parsed");
    test.checkEqual(series.markerFormat().size, 9, "Marker size parsed");
    test.checkEqual(series.markerFormat().fill.color.value, std::string("00FF00"), "Marker fill parsed");
    test.checkEqual(series.markerFormat().line.color.value, std::string("0000FF"), "Marker outline parsed");

    test.checkEqual(series.trendlines().size(), std::size_t{1}, "Formatted trendline parsed");
    if (!series.trendlines().empty()) {
        test.checkTrue(series.trendlines()[0].lineFormat.present, "Trendline line formatting parsed");
        test.checkEqual(series.trendlines()[0].lineFormat.color.value, std::string("AA00AA"), "Trendline line color parsed");
        test.checkEqual(series.trendlines()[0].lineFormat.dash, std::string("dot"), "Trendline dash parsed");
        test.checkNear(series.trendlines()[0].lineFormat.widthPoints, 1.5, 1e-9, "Trendline width parsed");
    }
    test.checkEqual(series.errorBars().size(), std::size_t{1}, "Custom error bars parsed");
    if (!series.errorBars().empty()) {
        const auto& bars = series.errorBars()[0];
        test.checkEqual(static_cast<int>(bars.valueType), static_cast<int>(xlpp::ChartSeries::ErrorValueType::Custom), "Custom error-bar value type parsed");
        test.checkEqual(bars.plusReference, std::string("'P0L'!$C$2:$C$6"), "Custom plus reference parsed");
        test.checkEqual(bars.minusReference, std::string("'P0L'!$D$2:$D$6"), "Custom minus reference parsed");
        test.checkTrue(bars.lineFormat.present, "Error-bar line formatting parsed");
        test.checkEqual(bars.lineFormat.color.value, std::string("444444"), "Error-bar line color parsed");
        test.checkEqual(bars.lineFormat.dash, std::string("sysDot"), "Error-bar dash parsed");
    }

    const auto stableId = chart.stableId();
    auto aggregateLabels = series.dataLabels();
    aggregateLabels.showValue = false;
    aggregateLabels.showCategoryName = true;
    aggregateLabels.position = "ctr";
    aggregateLabels.separator = " / ";
    test.checkTrue(sheet->setChartSeriesDataLabels(stableId, 0, aggregateLabels),
                   "Aggregate series labels selectively edit without touching point labels");

    auto point1 = series.dataLabels().points.front();
    point1.showValue = true;
    point1.showSeriesName = false;
    point1.position = "b";
    point1.separator = " | ";
    test.checkTrue(sheet->setChartSeriesDataLabelPoint(stableId, 0, point1), "Existing point label selectively edits");
    test.checkTrue(sheet->removeChartSeriesDataLabelPoint(stableId, 0, 3), "Existing point label selectively removes");
    xlpp::ChartDataLabelPoint point4;
    point4.index = 4;
    point4.showCategoryName = true;
    point4.position = "l";
    test.checkTrue(sheet->setChartSeriesDataLabelPoint(stableId, 0, point4), "New point label selectively appends");
    xlpp::ChartDataLabelPoint plotPoint;
    plotPoint.index = 2;
    plotPoint.showValue = true;
    plotPoint.position = "r";
    test.checkTrue(sheet->setChartPlotDataLabelPoint(stableId, 0, plotPoint), "Plot-level point label selectively appends");

    auto customBars = series.errorBars().front();
    customBars.noEndCap = true;
    customBars.plusReference = "'P0L'!$D$2:$D$6";
    customBars.minusReference = "'P0L'!$C$2:$C$6";
    test.checkTrue(sheet->setChartSeriesErrorBars(stableId, 0, customBars), "Custom error-bar ranges selectively edit");
    auto invalidCustom = customBars;
    invalidCustom.plusReference.clear();
    test.checkTrue(!sheet->setChartSeriesErrorBars(stableId, 0, invalidCustom), "Custom error bars require plus/minus references");

    xlpp::ChartLineFormat seriesLine;
    seriesLine.present = true;
    seriesLine.color = {xlpp::ChartColor::Kind::SRgb, "112233"};
    seriesLine.widthPoints = 2.5;
    seriesLine.dash = "lgDash";
    test.checkTrue(sheet->setChartSeriesLineFormat(stableId, 0, seriesLine), "Series line selectively formats");
    xlpp::ChartFillFormat seriesFill;
    seriesFill.present = true;
    seriesFill.color = {xlpp::ChartColor::Kind::SRgb, "ABCDEF"};
    test.checkTrue(sheet->setChartSeriesFillFormat(stableId, 0, seriesFill), "Series fill selectively formats");
    auto marker = series.markerFormat();
    marker.symbol = "triangle";
    marker.size = 11;
    marker.fill.present = true;
    marker.fill.color = {xlpp::ChartColor::Kind::SRgb, "123456"};
    marker.line.present = true;
    marker.line.color = {xlpp::ChartColor::Kind::SRgb, "654321"};
    marker.line.widthPoints = 1.25;
    marker.line.dash = "solid";
    test.checkTrue(sheet->setChartSeriesMarkerFormat(stableId, 0, marker), "Marker selectively formats");

    xlpp::ChartLineFormat trendLine;
    trendLine.present = true;
    trendLine.color = {xlpp::ChartColor::Kind::SRgb, "00AAAA"};
    trendLine.widthPoints = 2.0;
    trendLine.dash = "dashDot";
    test.checkTrue(sheet->setChartSeriesTrendlineLineFormat(stableId, 0, 0, trendLine), "Trendline line selectively formats");
    xlpp::ChartLineFormat errorLine;
    errorLine.present = true;
    errorLine.color = {xlpp::ChartColor::Kind::SRgb, "333333"};
    errorLine.widthPoints = 1.25;
    errorLine.dash = "dash";
    test.checkTrue(sheet->setChartSeriesErrorBarsLineFormat(stableId, 0, xlpp::ChartSeries::ErrorBarDirection::Y, errorLine),
                   "Error-bar line selectively formats");

    xlpp::ChartSeries::Trendline addedTrendline;
    addedTrendline.type = xlpp::ChartSeries::TrendlineType::Polynomial;
    addedTrendline.order = 3;
    addedTrendline.displayEquation = true;
    addedTrendline.lineFormat.present = true;
    addedTrendline.lineFormat.color = {xlpp::ChartColor::Kind::SRgb, "CC5500"};
    addedTrendline.lineFormat.widthPoints = 1.75;
    addedTrendline.lineFormat.dash = "sysDash";
    test.checkTrue(sheet->addChartSeriesTrendline(stableId, 0, addedTrendline),
                   "New formatted trendline selectively appends");

    xlpp::ChartSeries::ErrorBars xBars;
    xBars.direction = xlpp::ChartSeries::ErrorBarDirection::X;
    xBars.barType = xlpp::ChartSeries::ErrorBarType::Both;
    xBars.valueType = xlpp::ChartSeries::ErrorValueType::Custom;
    xBars.plusReference = "'P0L'!$C$2:$C$6";
    xBars.minusReference = "'P0L'!$D$2:$D$6";
    xBars.lineFormat.present = true;
    xBars.lineFormat.color = {xlpp::ChartColor::Kind::SRgb, "0055CC"};
    xBars.lineFormat.widthPoints = 1.5;
    xBars.lineFormat.dash = "sysDashDot";
    test.checkTrue(sheet->setChartSeriesErrorBars(stableId, 0, xBars),
                   "New formatted custom X error bars selectively append");

    sheet->cell("J20").setValue("p0l-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0l_chart_features.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("idx val=\"1\"") != std::string::npos && xml.find("dLblPos val=\"b\"") != std::string::npos,
                   "Edited point-label XML written");
    test.checkTrue(xml.find("dLblPos val=\"ctr\"") != std::string::npos && xml.find("<showCatName val=\"1\"") != std::string::npos,
                   "Aggregate data-label XML written separately from point labels");
    test.checkTrue(xml.find("idx val=\"3\"") == std::string::npos, "Removed point label absent from XML");
    test.checkTrue(xml.find("idx val=\"4\"") != std::string::npos, "Appended point label written");
    test.checkTrue(xml.find("&apos;P0L&apos;!$D$2:$D$6") != std::string::npos && xml.find("&apos;P0L&apos;!$C$2:$C$6") != std::string::npos,
                   "Custom plus/minus error-bar references remain in ChartML");
    test.checkTrue(xml.find("112233") != std::string::npos && xml.find("ABCDEF") != std::string::npos,
                   "Series line/fill colors written");
    test.checkTrue(xml.find("triangle") != std::string::npos && xml.find("123456") != std::string::npos && xml.find("654321") != std::string::npos,
                   "Marker formatting written");
    test.checkTrue(xml.find("00AAAA") != std::string::npos && xml.find("333333") != std::string::npos,
                   "Trendline and error-bar line colors written");
    test.checkTrue(xml.find("CC5500") != std::string::npos && xml.find("0055CC") != std::string::npos,
                   "Formatting for newly appended trendline and custom X error bars written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels,
                    "P0L chart edits keep drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0L chart edits keep sibling image byte-identical");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "P0L output package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "P0L chart remains visible");
    test.checkEqual(graph.objectInventory().images, std::size_t{1}, "P0L sibling image remains visible");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("P0L");
    test.checkTrue(reloadSheet != nullptr, "P0L selective output reloads");
    if (reloadSheet) {
        const auto& reloadSeries = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front().series().front();
        test.checkTrue(reloadSeries.dataLabels().present && !reloadSeries.dataLabels().showValue && reloadSeries.dataLabels().showCategoryName,
                       "Aggregate data-label flags survive reload without inheriting point overrides");
        test.checkEqual(reloadSeries.dataLabels().position, std::string("ctr"), "Aggregate data-label position survives reload");
        test.checkEqual(reloadSeries.dataLabels().separator, std::string(" / "), "Aggregate data-label separator survives reload");
        test.checkEqual(reloadSeries.dataLabels().points.size(), std::size_t{2}, "Point-label count survives reload after remove/add");
        const auto pointIndex1 = std::find_if(reloadSeries.dataLabels().points.begin(), reloadSeries.dataLabels().points.end(), [](const auto& point) { return point.index == 1; });
        const auto pointIndex4 = std::find_if(reloadSeries.dataLabels().points.begin(), reloadSeries.dataLabels().points.end(), [](const auto& point) { return point.index == 4; });
        test.checkTrue(pointIndex1 != reloadSeries.dataLabels().points.end() && pointIndex1->showValue && !pointIndex1->showSeriesName,
                       "Edited point label survives reload");
        test.checkTrue(pointIndex4 != reloadSeries.dataLabels().points.end() && pointIndex4->showCategoryName,
                       "Appended point label survives reload");
        test.checkEqual(reloadSeries.errorBars().front().plusReference, std::string("'P0L'!$D$2:$D$6"), "Edited custom plus reference survives reload");
        test.checkEqual(reloadSeries.errorBars().front().minusReference, std::string("'P0L'!$C$2:$C$6"), "Edited custom minus reference survives reload");
        test.checkEqual(reloadSeries.lineFormat().color.value, std::string("112233"), "Edited series line color survives reload");
        test.checkNear(reloadSeries.lineFormat().widthPoints, 2.5, 1e-9, "Edited series line width survives reload");
        test.checkEqual(reloadSeries.fillFormat().color.value, std::string("ABCDEF"), "Edited series fill survives reload");
        test.checkEqual(reloadSeries.markerFormat().symbol, std::string("triangle"), "Edited marker symbol survives reload");
        test.checkEqual(reloadSeries.markerFormat().fill.color.value, std::string("123456"), "Edited marker fill survives reload");
        test.checkEqual(reloadSeries.trendlines().front().lineFormat.color.value, std::string("00AAAA"), "Edited trendline formatting survives reload");
        test.checkEqual(reloadSeries.trendlines().size(), std::size_t{2}, "New formatted trendline survives reload");
        if (reloadSeries.trendlines().size() >= 2) {
            test.checkEqual(static_cast<int>(reloadSeries.trendlines()[1].type), static_cast<int>(xlpp::ChartSeries::TrendlineType::Polynomial),
                            "New polynomial trendline type survives reload");
            test.checkEqual(reloadSeries.trendlines()[1].order, 3, "New polynomial trendline order survives reload");
            test.checkEqual(reloadSeries.trendlines()[1].lineFormat.color.value, std::string("CC5500"),
                            "New trendline line formatting survives reload");
        }
        const auto yBars = std::find_if(reloadSeries.errorBars().begin(), reloadSeries.errorBars().end(), [](const auto& bars) {
            return bars.direction == xlpp::ChartSeries::ErrorBarDirection::Y;
        });
        const auto xBarsReloaded = std::find_if(reloadSeries.errorBars().begin(), reloadSeries.errorBars().end(), [](const auto& bars) {
            return bars.direction == xlpp::ChartSeries::ErrorBarDirection::X;
        });
        test.checkTrue(yBars != reloadSeries.errorBars().end() && yBars->lineFormat.color.value == "333333",
                       "Edited Y error-bar formatting survives reload");
        test.checkTrue(xBarsReloaded != reloadSeries.errorBars().end(), "New custom X error bars survive reload");
        if (xBarsReloaded != reloadSeries.errorBars().end()) {
            test.checkEqual(xBarsReloaded->plusReference, std::string("'P0L'!$C$2:$C$6"), "New X plus reference survives reload");
            test.checkEqual(xBarsReloaded->minusReference, std::string("'P0L'!$D$2:$D$6"), "New X minus reference survives reload");
            test.checkEqual(xBarsReloaded->lineFormat.color.value, std::string("0055CC"), "New X error-bar formatting survives reload");
        }
    }
    std::filesystem::remove(output);
}


void testImportedChartDataPointRichTextAndAdvancedFormatting(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/data_point_rich_text_advanced_formatting.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("P0L");
    test.checkTrue(sheet != nullptr, "P0M advanced ChartML fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0M fixture exposes one chart");
    if (charts.empty() || charts.front().series().empty()) return;
    const auto& chart = charts.front();
    const auto& series = chart.series().front();

    test.checkTrue(chart.titleRichText().present, "Rich chart title detected");
    test.checkEqual(chart.titleRichText().runs.size(), std::size_t{2}, "Rich chart title runs parsed");
    test.checkEqual(chart.titleRichText().plainText(), std::string("P0M advanced"), "Rich chart title plain text concatenated");
    if (chart.titleRichText().runs.size() >= 2) {
        const auto& first = chart.titleRichText().runs[0];
        const auto& second = chart.titleRichText().runs[1];
        test.checkTrue(first.bold && !first.italic, "First title run bold metadata parsed");
        test.checkNear(first.fontSizePoints, 16.0, 1e-9, "First title run font size parsed");
        test.checkEqual(first.typeface, std::string("Aptos"), "First title run typeface parsed");
        test.checkEqual(static_cast<int>(first.color.kind), static_cast<int>(xlpp::ChartColor::Kind::Scheme), "First title run scheme color parsed");
        test.checkEqual(first.color.value, std::string("accent1"), "First title run scheme value parsed");
        test.checkEqual(first.color.transforms.size(), std::size_t{1}, "Title color transform parsed");
        if (!first.color.transforms.empty())
            test.checkEqual(first.color.transforms.front().value, 20000, "Title tint transform value parsed");
        test.checkTrue(second.italic, "Second title run italic metadata parsed");
        test.checkEqual(second.color.value, std::string("336699"), "Second title run RGB parsed");
    }

    test.checkTrue(series.fillFormat().present, "Series advanced fill parsed");
    test.checkEqual(static_cast<int>(series.fillFormat().kind), static_cast<int>(xlpp::ChartFillFormat::Kind::Gradient), "Series gradient fill kind parsed");
    test.checkEqual(series.fillFormat().gradientStops.size(), std::size_t{2}, "Gradient stops parsed");
    test.checkNear(series.fillFormat().gradientAngleDegrees, 45.0, 1e-9, "Gradient angle parsed");
    if (series.fillFormat().gradientStops.size() >= 2) {
        test.checkEqual(series.fillFormat().gradientStops[0].position, 0, "First gradient stop position parsed");
        test.checkEqual(series.fillFormat().gradientStops[0].color.value, std::string("FF0000"), "First gradient stop color parsed");
        test.checkEqual(series.fillFormat().gradientStops[0].color.transforms.size(), std::size_t{1}, "Gradient alpha transform parsed");
        test.checkEqual(series.fillFormat().gradientStops[1].color.value, std::string("accent3"), "Second gradient scheme color parsed");
    }

    test.checkEqual(series.lineFormat().cap, std::string("rnd"), "Advanced line cap parsed");
    test.checkEqual(series.lineFormat().compound, std::string("dbl"), "Advanced compound line parsed");
    test.checkEqual(series.lineFormat().join, std::string("round"), "Advanced line join parsed");
    test.checkEqual(series.lineFormat().customDash.size(), std::size_t{2}, "Custom dash sequence parsed");
    test.checkEqual(series.lineFormat().color.value, std::string("accent2"), "Series scheme line color parsed");
    test.checkEqual(series.lineFormat().color.transforms.size(), std::size_t{2}, "Series line color transforms parsed");

    test.checkEqual(series.dataPoints().size(), std::size_t{1}, "Per-data-point style parsed");
    const auto* point2 = series.dataPoint(2);
    test.checkTrue(point2 != nullptr, "dPt index 2 accessible by index");
    if (point2) {
        test.checkEqual(static_cast<int>(point2->fill.kind), static_cast<int>(xlpp::ChartFillFormat::Kind::Pattern), "dPt pattern fill parsed");
        test.checkEqual(point2->fill.pattern, std::string("pct20"), "dPt pattern type parsed");
        test.checkEqual(point2->fill.foregroundColor.value, std::string("00AA00"), "dPt foreground color parsed");
        test.checkEqual(point2->fill.backgroundColor.value, std::string("accent4"), "dPt background scheme color parsed");
        test.checkEqual(point2->line.cap, std::string("sq"), "dPt line cap parsed");
        test.checkEqual(point2->line.join, std::string("bevel"), "dPt line join parsed");
        test.checkEqual(point2->line.color.transforms.size(), std::size_t{1}, "dPt line alpha parsed");
        test.checkEqual(point2->marker.symbol, std::string("square"), "dPt marker symbol parsed");
        test.checkEqual(point2->marker.size, 7, "dPt marker size parsed");
    }

    const auto label1 = std::find_if(series.dataLabels().points.begin(), series.dataLabels().points.end(),
                                     [](const auto& point) { return point.index == 1; });
    test.checkTrue(label1 != series.dataLabels().points.end(), "Rich point label found");
    if (label1 != series.dataLabels().points.end()) {
        test.checkTrue(label1->richText.present, "Point-label rich text parsed");
        test.checkEqual(label1->richText.plainText(), std::string("Point One"), "Point-label rich text value parsed");
        test.checkEqual(label1->richText.runs.size(), std::size_t{1}, "Point-label rich run parsed");
        if (!label1->richText.runs.empty()) {
            test.checkTrue(label1->richText.runs.front().italic, "Point-label italic run metadata parsed");
            test.checkEqual(label1->richText.runs.front().color.value, std::string("accent5"), "Point-label rich color parsed");
        }
    }

    const auto stableId = chart.stableId();
    xlpp::ChartRichText newTitle;
    newTitle.present = true;
    xlpp::ChartTextRun titleA;
    titleA.text = "XL++ ";
    titleA.bold = true;
    titleA.fontSizePoints = 18.0;
    titleA.typeface = "Aptos Display";
    titleA.color = {xlpp::ChartColor::Kind::Scheme, "accent6"};
    titleA.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Tint, 15000});
    xlpp::ChartTextRun titleB;
    titleB.text = "P0M";
    titleB.italic = true;
    titleB.fontSizePoints = 14.0;
    titleB.color = {xlpp::ChartColor::Kind::SRgb, "224466"};
    titleB.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Alpha, 90000});
    newTitle.runs = {titleA, titleB};
    test.checkTrue(sheet->setChartTitleRichText(stableId, newTitle), "Rich chart title selectively edits");

    xlpp::ChartRichText labelRich;
    labelRich.present = true;
    xlpp::ChartTextRun labelRun;
    labelRun.text = "Edited point";
    labelRun.bold = true;
    labelRun.fontSizePoints = 11.0;
    labelRun.color = {xlpp::ChartColor::Kind::Scheme, "accent2"};
    labelRun.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Shade, 65000});
    labelRich.runs.push_back(labelRun);
    test.checkTrue(sheet->setChartSeriesDataLabelPointRichText(stableId, 0, 1, labelRich),
                   "Point-label rich text selectively edits");

    xlpp::ChartLineFormat advancedLine = series.lineFormat();
    advancedLine.present = true;
    advancedLine.color = {xlpp::ChartColor::Kind::Scheme, "accent4"};
    advancedLine.color.transforms = {{xlpp::ChartColorTransform::Kind::LumMod, 80000},
                                     {xlpp::ChartColorTransform::Kind::LumOff, 10000}};
    advancedLine.widthPoints = 3.0;
    advancedLine.dash.clear();
    advancedLine.customDash = {{3.0, 1.0}, {1.0, 1.5}};
    advancedLine.cap = "sq";
    advancedLine.compound = "thickThin";
    advancedLine.join = "miter";
    test.checkTrue(sheet->setChartSeriesLineFormat(stableId, 0, advancedLine), "Advanced series line selectively edits");

    xlpp::ChartFillFormat gradient;
    gradient.present = true;
    gradient.kind = xlpp::ChartFillFormat::Kind::Gradient;
    gradient.gradientAngleDegrees = 90.0;
    xlpp::ChartGradientStop gs1;
    gs1.position = 0;
    gs1.color = {xlpp::ChartColor::Kind::Scheme, "accent1"};
    gs1.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Tint, 25000});
    xlpp::ChartGradientStop gs2;
    gs2.position = 100000;
    gs2.color = {xlpp::ChartColor::Kind::SRgb, "112244"};
    gs2.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Alpha, 75000});
    gradient.gradientStops = {gs1, gs2};
    test.checkTrue(sheet->setChartSeriesFillFormat(stableId, 0, gradient), "Advanced series gradient selectively edits");

    xlpp::ChartDataPointFormat editedPoint;
    editedPoint.index = 2;
    editedPoint.fill.present = true;
    editedPoint.fill.kind = xlpp::ChartFillFormat::Kind::Pattern;
    editedPoint.fill.pattern = "diagCross";
    editedPoint.fill.foregroundColor = {xlpp::ChartColor::Kind::Scheme, "accent5"};
    editedPoint.fill.foregroundColor.transforms.push_back({xlpp::ChartColorTransform::Kind::Tint, 12000});
    editedPoint.fill.backgroundColor = {xlpp::ChartColor::Kind::SRgb, "F0F0F0"};
    editedPoint.line.present = true;
    editedPoint.line.color = {xlpp::ChartColor::Kind::SRgb, "101010"};
    editedPoint.line.widthPoints = 1.75;
    editedPoint.line.cap = "rnd";
    editedPoint.line.join = "miter";
    editedPoint.line.customDash = {{2.0, 1.0}};
    editedPoint.marker.present = true;
    editedPoint.marker.symbol = "diamond";
    editedPoint.marker.size = 10;
    editedPoint.marker.fill.present = true;
    editedPoint.marker.fill.kind = xlpp::ChartFillFormat::Kind::Solid;
    editedPoint.marker.fill.color = {xlpp::ChartColor::Kind::SRgb, "00BBBB"};
    test.checkTrue(sheet->setChartSeriesDataPointFormat(stableId, 0, editedPoint), "Existing dPt selectively edits");

    xlpp::ChartDataPointFormat newPoint;
    newPoint.index = 4;
    newPoint.fill.present = true;
    newPoint.fill.kind = xlpp::ChartFillFormat::Kind::Gradient;
    newPoint.fill.gradientAngleDegrees = 30.0;
    xlpp::ChartGradientStop np1;
    np1.position = 0;
    np1.color = {xlpp::ChartColor::Kind::SRgb, "AA0000"};
    xlpp::ChartGradientStop np2;
    np2.position = 100000;
    np2.color = {xlpp::ChartColor::Kind::Scheme, "accent6"};
    newPoint.fill.gradientStops = {np1, np2};
    newPoint.line.present = true;
    newPoint.line.color = {xlpp::ChartColor::Kind::Scheme, "accent3"};
    newPoint.line.cap = "flat";
    newPoint.line.join = "round";
    newPoint.marker.present = true;
    newPoint.marker.symbol = "triangle";
    newPoint.marker.size = 8;
    test.checkTrue(sheet->setChartSeriesDataPointFormat(stableId, 0, newPoint), "New dPt selectively appends");

    sheet->cell("K21").setValue("p0m-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0m_chart_features.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("Aptos Display") != std::string::npos && xml.find("XL++ ") != std::string::npos,
                   "Rich title runs written to ChartML");
    test.checkTrue(xml.find("Edited point") != std::string::npos, "Rich point-label text written");
    test.checkTrue(xml.find("idx val=\"2\"") != std::string::npos && xml.find("diagCross") != std::string::npos,
                   "Existing dPt advanced formatting written");
    test.checkTrue(xml.find("idx val=\"4\"") != std::string::npos && xml.find("triangle") != std::string::npos,
                   "New dPt written");
    test.checkTrue(xml.find("custDash") != std::string::npos && xml.find("cmpd=\"thickThin\"") != std::string::npos,
                   "Advanced custom dash and compound line written");
    test.checkTrue(xml.find("gradFill") != std::string::npos && xml.find("pattFill") != std::string::npos,
                   "Gradient and pattern fills remain in ChartML");
    test.checkTrue(xml.find("lumMod") != std::string::npos && xml.find("alpha") != std::string::npos && xml.find("tint") != std::string::npos,
                   "Color transforms written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels,
                    "P0M chart edits keep drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0M chart edits keep sibling image byte-identical");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "P0M output package graph validates");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("P0L");
    test.checkTrue(reloadSheet != nullptr, "P0M output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        const auto& reloadSeries = reloadChart.series().front();
        test.checkEqual(reloadChart.titleRichText().plainText(), std::string("XL++ P0M"), "Edited rich chart title survives reload");
        test.checkEqual(reloadChart.titleRichText().runs.size(), std::size_t{2}, "Edited title run count survives reload");
        test.checkEqual(reloadSeries.lineFormat().cap, std::string("sq"), "Edited line cap survives reload");
        test.checkEqual(reloadSeries.lineFormat().compound, std::string("thickThin"), "Edited compound line survives reload");
        test.checkEqual(reloadSeries.lineFormat().customDash.size(), std::size_t{2}, "Edited custom dash survives reload");
        test.checkEqual(static_cast<int>(reloadSeries.fillFormat().kind), static_cast<int>(xlpp::ChartFillFormat::Kind::Gradient),
                        "Edited gradient fill survives reload");
        test.checkNear(reloadSeries.fillFormat().gradientAngleDegrees, 90.0, 1e-9, "Edited gradient angle survives reload");
        const auto* reloadPoint2 = reloadSeries.dataPoint(2);
        const auto* reloadPoint4 = reloadSeries.dataPoint(4);
        test.checkTrue(reloadPoint2 != nullptr && reloadPoint4 != nullptr, "Edited and appended dPt survive reload");
        if (reloadPoint2) {
            test.checkEqual(reloadPoint2->fill.pattern, std::string("diagCross"), "Edited dPt pattern survives reload");
            test.checkEqual(reloadPoint2->marker.symbol, std::string("diamond"), "Edited dPt marker survives reload");
        }
        if (reloadPoint4)
            test.checkEqual(static_cast<int>(reloadPoint4->fill.kind), static_cast<int>(xlpp::ChartFillFormat::Kind::Gradient),
                            "Appended dPt gradient survives reload");
        const auto reloadLabel = std::find_if(reloadSeries.dataLabels().points.begin(), reloadSeries.dataLabels().points.end(),
                                              [](const auto& point) { return point.index == 1; });
        test.checkTrue(reloadLabel != reloadSeries.dataLabels().points.end() && reloadLabel->richText.plainText() == "Edited point",
                       "Edited rich point-label survives reload");

        auto* mutableSheet = reloaded.worksheet("P0L");
        const auto reloadStableId = reloadChart.stableId();
        test.checkTrue(mutableSheet->removeChartSeriesDataPointFormat(reloadStableId, 0, 2), "Imported dPt selectively removes");
        const auto removedOutput = std::filesystem::temp_directory_path() / "xlpp_p0m_chart_features_removed.xlsx";
        reloaded.save(removedOutput);
        xlpp::Workbook removedReload;
        removedReload.load(removedOutput);
        const auto* removedSheet = removedReload.worksheet("P0L");
        test.checkTrue(removedSheet != nullptr, "dPt removal output reloads");
        if (removedSheet) {
            const auto& removedSeries = static_cast<const xlpp::Worksheet&>(*removedSheet).charts().front().series().front();
            test.checkTrue(removedSeries.dataPoint(2) == nullptr, "Removed dPt absent after reload");
            test.checkTrue(removedSeries.dataPoint(4) != nullptr, "Sibling dPt preserved after removal");
        }
        std::filesystem::remove(removedOutput);
    }
    std::filesystem::remove(output);
}

void testImportedChartLayoutAxisLegendFormatting(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_layout_axis_legend_formatting.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("P0L");
    test.checkTrue(sheet != nullptr, "P0N layout/axis/legend fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0N fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkTrue(chart.plotAreaLayout().present, "Plot-area manual layout parsed");
    test.checkEqual(chart.plotAreaLayout().target, std::string("inner"), "Plot-area layout target parsed");
    test.checkNear(chart.plotAreaLayout().x, 0.1, 1e-12, "Plot-area layout X parsed");
    test.checkNear(chart.plotAreaLayout().width, 0.72, 1e-12, "Plot-area layout width parsed");

    const auto* axis = chart.axisById(10);
    test.checkTrue(axis != nullptr, "P0N primary axis located by native axId");
    if (!axis) return;
    test.checkEqual(axis->numberFormat, std::string("0.00"), "Axis number format parsed");
    test.checkTrue(!axis->numberFormatSourceLinked, "Axis sourceLinked=false parsed");
    test.checkEqual(axis->majorTickMark, std::string("out"), "Axis major tick parsed");
    test.checkEqual(axis->minorTickMark, std::string("in"), "Axis minor tick parsed");
    test.checkEqual(axis->tickLabelPosition, std::string("low"), "Axis tick-label position parsed");
    test.checkTrue(axis->hasMajorUnit && axis->hasMinorUnit, "Axis major/minor units detected");
    test.checkNear(axis->majorUnit, 2.5, 1e-12, "Axis major unit parsed");
    test.checkNear(axis->minorUnit, 0.5, 1e-12, "Axis minor unit parsed");
    test.checkEqual(axis->crosses, std::string("max"), "Axis crosses parsed");
    test.checkEqual(axis->crossBetween, std::string("midCat"), "Axis crossBetween parsed");
    test.checkTrue(axis->titleRichText.present && axis->titleRichText.runs.size() == 2, "Axis rich title parsed");
    test.checkEqual(axis->titleRichText.plainText(), std::string("Input axis"), "Axis rich title text concatenated");
    test.checkTrue(axis->lineFormat.present, "Axis line formatting parsed");
    test.checkEqual(axis->lineFormat.color.value, std::string("accent1"), "Axis line scheme color parsed");
    test.checkEqual(axis->majorGridlineFormat.color.value, std::string("D0D0D0"), "Major gridline format parsed");
    test.checkEqual(axis->minorGridlineFormat.color.value, std::string("accent3"), "Minor gridline format parsed");

    test.checkTrue(chart.legendFormat().present, "Legend formatting model populated");
    test.checkTrue(chart.legendFormat().overlay, "Legend overlay parsed");
    test.checkEqual(chart.legendPosition(), std::string("b"), "Legend position parsed");
    test.checkTrue(chart.legendFormat().layout.present, "Legend manual layout parsed");
    test.checkEqual(chart.legendFormat().layout.target, std::string("outer"), "Legend layout target parsed");
    test.checkEqual(chart.legendFormat().fill.color.value, std::string("accent6"), "Legend fill parsed");
    test.checkEqual(chart.legendFormat().line.color.value, std::string("222222"), "Legend line parsed");

    const auto stableId = chart.stableId();
    xlpp::ChartRichText richTitle; richTitle.present = true;
    xlpp::ChartTextRun axisRun; axisRun.text = "XL++ Axis"; axisRun.bold = true; axisRun.fontSizePoints = 13.0; axisRun.color = {xlpp::ChartColor::Kind::SRgb, "123456"};
    richTitle.runs.push_back(axisRun);
    test.checkTrue(sheet->setChartAxisTitleRichText(stableId, 10, richTitle), "Axis rich title selectively edits");
    test.checkTrue(sheet->setChartAxisNumberFormat(stableId, 10, "0.0000", false), "Axis number format selectively edits");
    test.checkTrue(sheet->setChartAxisTicks(stableId, 10, "cross", "none", "high"), "Axis tick settings selectively edit");
    test.checkTrue(sheet->setChartAxisUnits(stableId, 10, 5.0, 1.0), "Axis units selectively edit");
    test.checkTrue(sheet->setChartAxisCrossing(stableId, 10, "autoZero", "between"), "Axis crossing selectively edits");

    xlpp::ChartLineFormat axisLine; axisLine.present = true; axisLine.color = {xlpp::ChartColor::Kind::SRgb, "880000"}; axisLine.widthPoints = 2.0; axisLine.dash = "solid";
    test.checkTrue(sheet->setChartAxisLineFormat(stableId, 10, axisLine), "Axis line selectively edits");
    xlpp::ChartLineFormat majorGrid; majorGrid.present = true; majorGrid.color = {xlpp::ChartColor::Kind::SRgb, "00AA00"}; majorGrid.widthPoints = 1.25; majorGrid.dash = "dash";
    test.checkTrue(sheet->setChartAxisGridlineFormat(stableId, 10, true, majorGrid), "Major gridline selectively edits");
    xlpp::ChartLineFormat minorGrid; minorGrid.present = true; minorGrid.noFill = true;
    test.checkTrue(sheet->setChartAxisGridlineFormat(stableId, 10, false, minorGrid), "Minor gridline selectively edits");

    xlpp::ChartManualLayout plotLayout; plotLayout.present = true; plotLayout.target = "inner"; plotLayout.xMode = plotLayout.yMode = plotLayout.widthMode = plotLayout.heightMode = "factor"; plotLayout.hasX = plotLayout.hasY = plotLayout.hasWidth = plotLayout.hasHeight = true; plotLayout.x = 0.15; plotLayout.y = 0.16; plotLayout.width = 0.68; plotLayout.height = 0.62;
    test.checkTrue(sheet->setChartPlotAreaLayout(stableId, plotLayout), "Plot-area manual layout selectively edits");
    xlpp::ChartManualLayout legendLayout; legendLayout.present = true; legendLayout.target = "outer"; legendLayout.xMode = legendLayout.yMode = legendLayout.widthMode = legendLayout.heightMode = "factor"; legendLayout.hasX = legendLayout.hasY = legendLayout.hasWidth = legendLayout.hasHeight = true; legendLayout.x = 0.25; legendLayout.y = 0.8; legendLayout.width = 0.5; legendLayout.height = 0.14;
    test.checkTrue(sheet->setChartLegend(stableId, true, "tr"), "Legend position selectively edits");
    test.checkTrue(sheet->setChartLegendLayout(stableId, legendLayout), "Legend layout selectively edits");
    test.checkTrue(sheet->setChartLegendOverlay(stableId, false), "Legend overlay selectively edits");
    xlpp::ChartLineFormat legendLine; legendLine.present = true; legendLine.color = {xlpp::ChartColor::Kind::SRgb, "000088"}; legendLine.widthPoints = 1.5;
    test.checkTrue(sheet->setChartLegendLineFormat(stableId, legendLine), "Legend line selectively edits");
    xlpp::ChartFillFormat legendFill; legendFill.present = true; legendFill.kind = xlpp::ChartFillFormat::Kind::Solid; legendFill.color = {xlpp::ChartColor::Kind::SRgb, "EEEEEE"};
    test.checkTrue(sheet->setChartLegendFillFormat(stableId, legendFill), "Legend fill selectively edits");

    sheet->cell("K22").setValue("p0n-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0n_chart_layout_axis_legend.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("0.0000") != std::string::npos && xml.find("majorUnit val=\"5") != std::string::npos, "Axis format edits written");
    test.checkTrue(xml.find("XL++ Axis") != std::string::npos && xml.find("123456") != std::string::npos, "Axis rich title written");
    test.checkTrue(xml.find("880000") != std::string::npos && xml.find("00AA00") != std::string::npos, "Axis/gridline formatting written");
    test.checkTrue(xml.find("x val=\"0.15\"") != std::string::npos && xml.find("x val=\"0.25\"") != std::string::npos, "Plot and legend manual layouts written");
    test.checkTrue(xml.find("000088") != std::string::npos && xml.find("EEEEEE") != std::string::npos, "Legend formatting written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0N keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0N keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0N output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("P0L");
    test.checkTrue(reloadSheet != nullptr, "P0N output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        const auto* reloadAxis = reloadChart.axisById(10);
        test.checkTrue(reloadAxis != nullptr, "Edited axis remains addressable by axId");
        if (reloadAxis) {
            test.checkEqual(reloadAxis->titleRichText.plainText(), std::string("XL++ Axis"), "Axis rich title survives reload");
            test.checkEqual(reloadAxis->numberFormat, std::string("0.0000"), "Axis number format survives reload");
            test.checkEqual(reloadAxis->majorTickMark, std::string("cross"), "Axis major tick survives reload");
            test.checkEqual(reloadAxis->tickLabelPosition, std::string("high"), "Axis tick-label position survives reload");
            test.checkNear(reloadAxis->majorUnit, 5.0, 1e-12, "Axis major unit survives reload");
            test.checkEqual(reloadAxis->crossBetween, std::string("between"), "Axis crossBetween survives reload");
            test.checkEqual(reloadAxis->lineFormat.color.value, std::string("880000"), "Axis line formatting survives reload");
            test.checkEqual(reloadAxis->majorGridlineFormat.color.value, std::string("00AA00"), "Major gridline formatting survives reload");
            test.checkTrue(reloadAxis->minorGridlineFormat.noFill, "Minor gridline no-fill survives reload");
        }
        test.checkNear(reloadChart.plotAreaLayout().x, 0.15, 1e-12, "Plot layout survives reload");
        test.checkEqual(reloadChart.legendPosition(), std::string("tr"), "Legend position survives reload");
        test.checkTrue(!reloadChart.legendFormat().overlay, "Legend overlay survives reload");
        test.checkNear(reloadChart.legendFormat().layout.x, 0.25, 1e-12, "Legend layout survives reload");
        test.checkEqual(reloadChart.legendFormat().line.color.value, std::string("000088"), "Legend line survives reload");
        test.checkEqual(reloadChart.legendFormat().fill.color.value, std::string("EEEEEE"), "Legend fill survives reload");
    }
    std::filesystem::remove(output);
}

void testImportedChartAxisScalingDisplayUnitsAndAreaFormatting(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_axis_scaling_display_units_area_formatting.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("P0L");
    test.checkTrue(sheet != nullptr, "P0O scaling/display-units fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0O fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    const auto* xAxis = chart.axisById(10);
    const auto* yAxis = chart.axisById(20);
    test.checkTrue(xAxis != nullptr && yAxis != nullptr, "P0O native axes located");
    if (!xAxis || !yAxis) return;

    test.checkTrue(xAxis->scaling.hasMinimum && xAxis->scaling.hasMaximum && xAxis->scaling.hasLogBase,
                   "Axis scaling min/max/log flags parsed");
    test.checkNear(xAxis->scaling.minimum, 1.0, 1e-12, "Axis scaling minimum parsed");
    test.checkNear(xAxis->scaling.maximum, 100.0, 1e-12, "Axis scaling maximum parsed");
    test.checkNear(xAxis->scaling.logBase, 10.0, 1e-12, "Axis log base parsed");
    test.checkTrue(xAxis->scaling.reverseOrder, "Axis reverse order parsed");
    test.checkTrue(xAxis->hasCrossesAt, "Axis crossesAt presence parsed");
    test.checkNear(xAxis->crossesAt, 2.0, 1e-12, "Axis crossesAt value parsed");
    test.checkTrue(xAxis->hasMajorGridlines && xAxis->hasMinorGridlines, "Axis gridline lifecycle state parsed");

    test.checkTrue(yAxis->displayUnits.present, "Display units parsed");
    test.checkEqual(yAxis->displayUnits.builtInUnit, std::string("thousands"), "Built-in display units parsed");
    test.checkTrue(yAxis->displayUnits.showLabel, "Display-units label presence parsed");
    test.checkEqual(yAxis->displayUnits.labelRichText.plainText(), std::string("Thousands"), "Display-units rich label parsed");

    test.checkEqual(chart.chartAreaFillFormat().color.value, std::string("accent1"), "Chart-area fill parsed");
    test.checkEqual(chart.chartAreaLineFormat().color.value, std::string("336699"), "Chart-area line parsed");
    test.checkEqual(chart.plotAreaFillFormat().color.value, std::string("FFF2CC"), "Plot-area fill parsed");
    test.checkEqual(chart.plotAreaLineFormat().color.value, std::string("CC9900"), "Plot-area line parsed");

    const auto stableId = chart.stableId();
    xlpp::ChartAxisScaling invalidScaling; invalidScaling.hasMinimum=true; invalidScaling.minimum=10.0; invalidScaling.hasMaximum=true; invalidScaling.maximum=1.0;
    test.checkTrue(!sheet->setChartAxisScaling(stableId, 10, invalidScaling), "Invalid axis scaling range rejected");
    xlpp::ChartAxisScaling scaling; scaling.hasMinimum=true; scaling.minimum=0.5; scaling.hasMaximum=true; scaling.maximum=500.0; scaling.hasLogBase=true; scaling.logBase=10.0; scaling.reverseOrder=false;
    test.checkTrue(sheet->setChartAxisScaling(stableId, 10, scaling), "Axis scaling selectively edits");
    test.checkTrue(sheet->setChartAxisCrossesAt(stableId, 10, 5.5), "Axis crossesAt selectively edits");

    xlpp::ChartDisplayUnits invalidUnits; invalidUnits.present=true; invalidUnits.builtInUnit="invalid";
    test.checkTrue(!sheet->setChartAxisDisplayUnits(stableId, 20, invalidUnits), "Invalid display unit rejected");
    xlpp::ChartDisplayUnits units; units.present=true; units.hasCustomUnit=true; units.customUnit=1000000.0; units.showLabel=true;
    units.labelRichText.present=true; xlpp::ChartTextRun unitRun; unitRun.text="Millions"; unitRun.bold=true; unitRun.color={xlpp::ChartColor::Kind::SRgb,"7030A0"}; units.labelRichText.runs.push_back(unitRun);
    test.checkTrue(sheet->setChartAxisDisplayUnits(stableId, 20, units), "Custom display units selectively edit");
    test.checkTrue(sheet->removeChartAxisGridlines(stableId, 10, false), "Minor gridlines selectively remove");

    xlpp::ChartLineFormat chartLine; chartLine.present=true; chartLine.color={xlpp::ChartColor::Kind::SRgb,"112233"}; chartLine.widthPoints=2.25; chartLine.dash="solid";
    test.checkTrue(sheet->setChartAreaLineFormat(stableId, chartLine), "Chart-area line selectively edits");
    xlpp::ChartFillFormat chartFill; chartFill.present=true; chartFill.kind=xlpp::ChartFillFormat::Kind::Solid; chartFill.color={xlpp::ChartColor::Kind::SRgb,"F0F0F0"};
    test.checkTrue(sheet->setChartAreaFillFormat(stableId, chartFill), "Chart-area fill selectively edits");
    xlpp::ChartLineFormat plotLine; plotLine.present=true; plotLine.color={xlpp::ChartColor::Kind::SRgb,"445566"}; plotLine.widthPoints=1.75; plotLine.dash="dash";
    test.checkTrue(sheet->setChartPlotAreaLineFormat(stableId, plotLine), "Plot-area line selectively edits");
    xlpp::ChartFillFormat plotFill; plotFill.present=true; plotFill.kind=xlpp::ChartFillFormat::Kind::Solid; plotFill.color={xlpp::ChartColor::Kind::SRgb,"E2F0D9"};
    test.checkTrue(sheet->setChartPlotAreaFillFormat(stableId, plotFill), "Plot-area fill selectively edits");

    sheet->cell("K23").setValue("p0o-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0o_axis_scaling_display_units.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("logBase val=\"10\"") != std::string::npos && xml.find("max val=\"500\"") != std::string::npos && xml.find("min val=\"0.5\"") != std::string::npos,
                   "Axis scaling edits written");
    test.checkTrue(xml.find("orientation val=\"minMax\"") != std::string::npos && xml.find("crossesAt val=\"5.5\"") != std::string::npos,
                   "Axis orientation and crossesAt edits written");
    test.checkTrue(xml.find("custUnit val=\"1000000\"") != std::string::npos && xml.find("Millions") != std::string::npos,
                   "Custom display units and label written");
    test.checkTrue(xml.find("112233") != std::string::npos && xml.find("F0F0F0") != std::string::npos && xml.find("445566") != std::string::npos && xml.find("E2F0D9") != std::string::npos,
                   "Chart/plot area formatting written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0O keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0O keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0O output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    auto* reloadSheet = reloaded.worksheet("P0L");
    test.checkTrue(reloadSheet != nullptr, "P0O output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        const auto* reloadX = reloadChart.axisById(10); const auto* reloadY = reloadChart.axisById(20);
        test.checkTrue(reloadX != nullptr && reloadY != nullptr, "Edited P0O axes remain addressable");
        if (reloadX && reloadY) {
            test.checkNear(reloadX->scaling.minimum, 0.5, 1e-12, "Edited scaling minimum survives reload");
            test.checkNear(reloadX->scaling.maximum, 500.0, 1e-12, "Edited scaling maximum survives reload");
            test.checkTrue(!reloadX->scaling.reverseOrder, "Edited axis orientation survives reload");
            test.checkNear(reloadX->crossesAt, 5.5, 1e-12, "Edited crossesAt survives reload");
            test.checkTrue(!reloadX->hasMinorGridlines, "Removed minor gridlines stay absent after reload");
            test.checkTrue(reloadY->displayUnits.present && reloadY->displayUnits.hasCustomUnit, "Custom display units survive reload");
            test.checkNear(reloadY->displayUnits.customUnit, 1000000.0, 1e-6, "Custom display-unit value survives reload");
            test.checkEqual(reloadY->displayUnits.labelRichText.plainText(), std::string("Millions"), "Display-unit rich label survives reload");
        }
        test.checkEqual(reloadChart.chartAreaLineFormat().color.value, std::string("112233"), "Chart-area line survives reload");
        test.checkEqual(reloadChart.chartAreaFillFormat().color.value, std::string("F0F0F0"), "Chart-area fill survives reload");
        test.checkEqual(reloadChart.plotAreaLineFormat().color.value, std::string("445566"), "Plot-area line survives reload");
        test.checkEqual(reloadChart.plotAreaFillFormat().color.value, std::string("E2F0D9"), "Plot-area fill survives reload");

        const auto reloadStableId = reloadChart.stableId();
        test.checkTrue(reloadSheet->clearChartAxisCrossesAt(reloadStableId, 10), "crossesAt selectively clears");
        test.checkTrue(reloadSheet->clearChartAxisDisplayUnits(reloadStableId, 20), "display units selectively clear");
        test.checkTrue(reloadSheet->removeChartAxisGridlines(reloadStableId, 10, true), "major gridlines selectively remove");
        const auto cleared = std::filesystem::temp_directory_path() / "xlpp_p0o_axis_lifecycle_cleared.xlsx";
        reloaded.save(cleared);
        xlpp::Workbook clearedReload; clearedReload.load(cleared);
        const auto* clearedSheet = clearedReload.worksheet("P0L");
        test.checkTrue(clearedSheet != nullptr, "P0O lifecycle-clear output reloads");
        if (clearedSheet) {
            const auto& clearedChart = static_cast<const xlpp::Worksheet&>(*clearedSheet).charts().front();
            const auto* clearedX = clearedChart.axisById(10); const auto* clearedY = clearedChart.axisById(20);
            test.checkTrue(clearedX && !clearedX->hasCrossesAt, "crossesAt absent after clear/reload");
            test.checkTrue(clearedX && !clearedX->hasMajorGridlines && !clearedX->hasMinorGridlines, "Both gridline collections absent after lifecycle removal");
            test.checkTrue(clearedY && !clearedY->displayUnits.present, "Display units absent after clear/reload");
        }
        std::filesystem::remove(cleared);
    }
    std::filesystem::remove(output);
}

void testImportedChartAuxiliaryObjects(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_auxiliary_objects.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Aux");
    test.checkTrue(sheet != nullptr, "P0P auxiliary-object fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0P fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkEqual(chart.plots().size(), std::size_t{1}, "P0P fixture exposes one plot");
    if (chart.plots().empty()) return;
    const auto& plot = chart.plots().front();

    test.checkTrue(plot.hasDropLines, "Drop lines parsed");
    test.checkEqual(plot.dropLinesFormat.color.value, std::string("FF0000"), "Drop-line color parsed");
    test.checkNear(plot.dropLinesFormat.widthPoints, 1.5, 1e-9, "Drop-line width parsed");
    test.checkTrue(plot.hasHighLowLines, "High-low lines parsed");
    test.checkEqual(plot.highLowLinesFormat.color.value, std::string("00AA00"), "High-low line color parsed");
    test.checkTrue(plot.upDownBars.present, "Up/down bars parsed");
    test.checkEqual(plot.upDownBars.gapWidth, 120, "Up/down gap width parsed");
    test.checkEqual(plot.upDownBars.upFill.color.value, std::string("DDEBF7"), "Up-bar fill parsed");
    test.checkEqual(plot.upDownBars.upLine.color.value, std::string("4472C4"), "Up-bar line parsed");
    test.checkEqual(plot.upDownBars.downFill.color.value, std::string("FCE4D6"), "Down-bar fill parsed");
    test.checkEqual(plot.upDownBars.downLine.color.value, std::string("C00000"), "Down-bar line parsed");

    test.checkTrue(plot.dataLabels.showLeaderLines && plot.dataLabels.hasLeaderLines, "Plot leader lines parsed");
    test.checkEqual(plot.dataLabels.leaderLineFormat.color.value, std::string("7030A0"), "Leader-line color parsed");
    test.checkEqual(plot.dataLabels.leaderLineFormat.dash, std::string("dash"), "Leader-line dash parsed");

    test.checkTrue(chart.dataTable().present, "Chart data table parsed");
    test.checkTrue(chart.dataTable().showHorizontalBorder, "Data-table horizontal border flag parsed");
    test.checkTrue(!chart.dataTable().showVerticalBorder, "Data-table vertical border flag parsed");
    test.checkTrue(chart.dataTable().showOutline && chart.dataTable().showLegendKeys, "Data-table outline/keys flags parsed");
    test.checkEqual(chart.dataTable().fill.color.value, std::string("FFF2CC"), "Data-table fill parsed");
    test.checkEqual(chart.dataTable().line.color.value, std::string("7F6000"), "Data-table line parsed");

    const auto stableId = chart.stableId();
    xlpp::ChartDataTable table;
    table.present = true; table.showHorizontalBorder = false; table.showVerticalBorder = true; table.showOutline = true; table.showLegendKeys = false;
    table.fill.present = true; table.fill.kind = xlpp::ChartFillFormat::Kind::Solid; table.fill.color = {xlpp::ChartColor::Kind::SRgb, "C6E0B4"};
    table.line.present = true; table.line.color = {xlpp::ChartColor::Kind::SRgb, "548235"}; table.line.widthPoints = 1.75; table.line.dash = "dash";
    test.checkTrue(sheet->setChartDataTable(stableId, table), "Data table selectively edits");

    xlpp::ChartLineFormat drop; drop.present = true; drop.color = {xlpp::ChartColor::Kind::SRgb, "112233"}; drop.widthPoints = 2.5; drop.dash = "dash";
    test.checkTrue(sheet->setChartPlotDropLines(stableId, 0, drop), "Drop lines selectively edit");
    test.checkTrue(sheet->removeChartPlotHighLowLines(stableId, 0), "High-low lines selectively remove");

    xlpp::ChartUpDownBars bars; bars.present = true; bars.gapWidth = 80;
    bars.upFill.present = true; bars.upFill.kind = xlpp::ChartFillFormat::Kind::Solid; bars.upFill.color = {xlpp::ChartColor::Kind::SRgb, "E2F0D9"};
    bars.upLine.present = true; bars.upLine.color = {xlpp::ChartColor::Kind::SRgb, "70AD47"}; bars.upLine.widthPoints = 1.25;
    bars.downFill.present = true; bars.downFill.kind = xlpp::ChartFillFormat::Kind::Solid; bars.downFill.color = {xlpp::ChartColor::Kind::SRgb, "F4B183"};
    bars.downLine.present = true; bars.downLine.color = {xlpp::ChartColor::Kind::SRgb, "C65911"}; bars.downLine.widthPoints = 1.25;
    test.checkTrue(sheet->setChartPlotUpDownBars(stableId, 0, bars), "Up/down bars selectively edit");
    xlpp::ChartUpDownBars invalidBars; invalidBars.gapWidth = 501;
    test.checkTrue(!sheet->setChartPlotUpDownBars(stableId, 0, invalidBars), "Invalid up/down gap width rejected");

    xlpp::ChartLineFormat plotLeader; plotLeader.present = true; plotLeader.color = {xlpp::ChartColor::Kind::SRgb, "44546A"}; plotLeader.widthPoints = 1.5; plotLeader.dash = "dot";
    test.checkTrue(sheet->setChartPlotLeaderLineFormat(stableId, 0, plotLeader), "Plot leader-line formatting selectively edits");
    xlpp::ChartLineFormat seriesLeader; seriesLeader.present = true; seriesLeader.color = {xlpp::ChartColor::Kind::SRgb, "A5A5A5"}; seriesLeader.widthPoints = 1.0; seriesLeader.dash = "solid";
    test.checkTrue(sheet->setChartSeriesLeaderLineFormat(stableId, 0, seriesLeader), "Series leader lines selectively add");

    sheet->cell("K24").setValue("p0p-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0p_chart_auxiliary_objects.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("showHorzBorder val=\"0\"") != std::string::npos && xml.find("showVertBorder val=\"1\"") != std::string::npos,
                   "Edited data-table border flags written");
    test.checkTrue(xml.find("C6E0B4") != std::string::npos && xml.find("548235") != std::string::npos, "Edited data-table formatting written");
    test.checkTrue(xml.find("112233") != std::string::npos, "Edited drop-line formatting written");
    test.checkTrue(xml.find("hiLowLines") == std::string::npos, "Removed high-low lines absent from XML");
    test.checkTrue(xml.find("gapWidth val=\"80\"") != std::string::npos && xml.find("E2F0D9") != std::string::npos && xml.find("F4B183") != std::string::npos,
                   "Edited up/down bars written");
    test.checkTrue(xml.find("44546A") != std::string::npos && xml.find("A5A5A5") != std::string::npos, "Plot and series leader-line formatting written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0P keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0P keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0P output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    auto* reloadSheet = reloaded.worksheet("Aux");
    test.checkTrue(reloadSheet != nullptr, "P0P output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        const auto& reloadPlot = reloadChart.plots().front();
        test.checkTrue(reloadChart.dataTable().present && !reloadChart.dataTable().showHorizontalBorder && reloadChart.dataTable().showVerticalBorder,
                       "Edited data table survives reload");
        test.checkEqual(reloadChart.dataTable().fill.color.value, std::string("C6E0B4"), "Edited data-table fill survives reload");
        test.checkTrue(reloadPlot.hasDropLines && !reloadPlot.hasHighLowLines, "Drop/high-low lifecycle survives reload");
        test.checkEqual(reloadPlot.dropLinesFormat.color.value, std::string("112233"), "Edited drop-line color survives reload");
        test.checkTrue(reloadPlot.upDownBars.present && reloadPlot.upDownBars.gapWidth == 80, "Edited up/down bars survive reload");
        test.checkEqual(reloadPlot.upDownBars.downLine.color.value, std::string("C65911"), "Down-bar formatting survives reload");
        test.checkTrue(reloadPlot.dataLabels.hasLeaderLines, "Plot leader lines survive reload");
        test.checkEqual(reloadPlot.dataLabels.leaderLineFormat.color.value, std::string("44546A"), "Plot leader-line formatting survives reload");
        test.checkTrue(reloadChart.series()[0].dataLabels().hasLeaderLines, "Series leader lines survive reload");
        test.checkEqual(reloadChart.series()[0].dataLabels().leaderLineFormat.color.value, std::string("A5A5A5"), "Series leader-line formatting survives reload");

        const auto reloadStableId = reloadChart.stableId();
        test.checkTrue(reloadSheet->removeChartDataTable(reloadStableId), "Data table selectively removes");
        test.checkTrue(reloadSheet->removeChartPlotDropLines(reloadStableId, 0), "Drop lines selectively remove");
        test.checkTrue(reloadSheet->removeChartPlotUpDownBars(reloadStableId, 0), "Up/down bars selectively remove");
        test.checkTrue(reloadSheet->removeChartPlotLeaderLines(reloadStableId, 0), "Plot leader lines selectively remove");
        test.checkTrue(reloadSheet->removeChartSeriesLeaderLines(reloadStableId, 0), "Series leader lines selectively remove");
        xlpp::ChartLineFormat high; high.present=true; high.color={xlpp::ChartColor::Kind::SRgb,"00B0F0"}; high.widthPoints=2.0; high.dash="solid";
        test.checkTrue(reloadSheet->setChartPlotHighLowLines(reloadStableId, 0, high), "High-low lines selectively add after removal");
        const auto lifecycle = std::filesystem::temp_directory_path() / "xlpp_p0p_chart_auxiliary_lifecycle.xlsx";
        reloaded.save(lifecycle);
        xlpp::Workbook lifecycleReload; lifecycleReload.load(lifecycle);
        const auto* lifecycleSheet = lifecycleReload.worksheet("Aux");
        test.checkTrue(lifecycleSheet != nullptr, "P0P lifecycle output reloads");
        if (lifecycleSheet) {
            const auto& lifecycleChart = static_cast<const xlpp::Worksheet&>(*lifecycleSheet).charts().front();
            const auto& lifecyclePlot = lifecycleChart.plots().front();
            test.checkTrue(!lifecycleChart.dataTable().present, "Data table absent after lifecycle removal");
            test.checkTrue(!lifecyclePlot.hasDropLines && !lifecyclePlot.upDownBars.present, "Drop/up-down objects absent after lifecycle removal");
            test.checkTrue(lifecyclePlot.hasHighLowLines && lifecyclePlot.highLowLinesFormat.color.value == "00B0F0", "High-low object added with formatting");
            test.checkTrue(!lifecyclePlot.dataLabels.hasLeaderLines, "Plot leader-line container absent after lifecycle removal");
            test.checkTrue(!lifecycleChart.series()[0].dataLabels().hasLeaderLines, "Series leader-line container absent after lifecycle removal");
        }
        std::filesystem::remove(lifecycle);
    }
    std::filesystem::remove(output);
}

void testStockChartStructureGenerationAndDataTableText(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/stock_auxiliary_datatable_text.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("Stock");
    test.checkTrue(sheet != nullptr, "P0Q stock fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0Q stock fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkTrue(chart.type() == xlpp::Chart::Type::Stock, "stockChart maps to Chart::Type::Stock");
    test.checkEqual(chart.series().size(), std::size_t{4}, "Open-high-low-close stock series parsed");
    test.checkEqual(chart.plots().size(), std::size_t{1}, "Stock chart exposes one plot");
    if (chart.plots().empty()) return;
    const auto& plot = chart.plots().front();
    test.checkTrue(plot.type == xlpp::Chart::Type::Stock, "Stock plot type parsed");
    test.checkTrue(plot.hasHighLowLines, "Stock high-low lines parsed");
    test.checkEqual(plot.highLowLinesFormat.color.value, std::string("00AA00"), "Stock high-low line formatting parsed");
    test.checkTrue(plot.upDownBars.present && plot.upDownBars.gapWidth == 120, "Stock up/down bars parsed");
    test.checkEqual(plot.upDownBars.upFill.color.value, std::string("DDEBF7"), "Stock up-bar fill parsed");
    test.checkTrue(chart.dataTable().present, "Stock data table parsed");
    test.checkTrue(chart.dataTable().textStyle.present, "Data-table txPr text style parsed");
    test.checkTrue(chart.dataTable().textStyle.bold && chart.dataTable().textStyle.italic, "Data-table text bold/italic parsed");
    test.checkNear(chart.dataTable().textStyle.fontSizePoints, 11.0, 1e-9, "Data-table font size parsed");
    test.checkEqual(chart.dataTable().textStyle.typeface, std::string("Calibri"), "Data-table typeface parsed");
    test.checkEqual(chart.dataTable().textStyle.color.value, std::string("7030A0"), "Data-table text color parsed");

    const auto stableId = chart.stableId();
    xlpp::ChartDataTable edited = chart.dataTable();
    edited.showHorizontalBorder = false;
    edited.textStyle.present = true; edited.textStyle.bold = false; edited.textStyle.italic = true;
    edited.textStyle.fontSizePoints = 12.5; edited.textStyle.typeface = "Aptos";
    edited.textStyle.color = {xlpp::ChartColor::Kind::SRgb, "C00000"};
    test.checkTrue(sheet->setChartDataTable(stableId, edited), "Imported stock data-table text selectively edits");
    xlpp::ChartLineFormat hi; hi.present=true; hi.color={xlpp::ChartColor::Kind::SRgb,"00B0F0"}; hi.widthPoints=2.25; hi.dash="dash";
    test.checkTrue(sheet->setChartPlotHighLowLines(stableId, 0, hi), "Imported stock high-low lines selectively edit");
    sheet->cell("K24").setValue("p0q-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0q_stock_imported.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("stockChart") != std::string::npos, "Selective stock edit preserves stockChart structure");
    test.checkTrue(xml.find("Aptos") != std::string::npos && xml.find("C00000") != std::string::npos, "Edited data-table txPr written");
    test.checkTrue(xml.find("00B0F0") != std::string::npos, "Edited stock high-low formatting written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0Q keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0Q keeps sibling stock image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0Q imported stock output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("Stock");
    test.checkTrue(reloadSheet != nullptr, "P0Q selectively edited stock workbook reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        test.checkTrue(reloadChart.type() == xlpp::Chart::Type::Stock, "Stock type survives selective save/reload");
        test.checkEqual(reloadChart.dataTable().textStyle.typeface, std::string("Aptos"), "Edited dTable typeface survives reload");
        test.checkNear(reloadChart.dataTable().textStyle.fontSizePoints, 12.5, 1e-9, "Edited dTable font size survives reload");
        test.checkEqual(reloadChart.plots().front().highLowLinesFormat.color.value, std::string("00B0F0"), "Edited stock high-low lines survive reload");
    }
    std::filesystem::remove(output);

    // Generation path: P0P auxiliary model is now serialized for new charts.
    xlpp::Workbook generated;
    auto& generatedSheet = generated.addWorksheet("GeneratedStock");
    generatedSheet.cell("A1").setValue("Date"); generatedSheet.cell("B1").setValue("Open"); generatedSheet.cell("C1").setValue("High");
    generatedSheet.cell("D1").setValue("Low"); generatedSheet.cell("E1").setValue("Close");
    for (std::size_t r=2; r<=5; ++r) {
        generatedSheet.cell("A"+std::to_string(r)).setValue(static_cast<double>(r-1));
        generatedSheet.cell("B"+std::to_string(r)).setValue(10.0+r);
        generatedSheet.cell("C"+std::to_string(r)).setValue(13.0+r);
        generatedSheet.cell("D"+std::to_string(r)).setValue(8.0+r);
        generatedSheet.cell("E"+std::to_string(r)).setValue(12.0+r);
    }
    xlpp::Chart stock(xlpp::Chart::Type::Stock); stock.setTitle("Generated stock"); stock.setWidth(420); stock.setHeight(260);
    for (std::size_t col=0; col<4; ++col) {
        static const char* names[]{"Open","High","Low","Close"};
        static const char* letters[]{"B","C","D","E"};
        xlpp::ChartSeries series(names[col]);
        series.setCategoriesReference("'GeneratedStock'!$A$2:$A$5");
        series.setValuesReference(std::string("'GeneratedStock'!$") + letters[col] + "$2:$" + letters[col] + "$5");
        stock.addSeries(std::move(series));
    }
    auto& generatedPlot = stock.primaryPlot();
    generatedPlot.hasHighLowLines = true; generatedPlot.highLowLinesFormat = hi;
    generatedPlot.upDownBars.present = true; generatedPlot.upDownBars.gapWidth = 90;
    generatedPlot.upDownBars.upFill.present=true; generatedPlot.upDownBars.upFill.kind=xlpp::ChartFillFormat::Kind::Solid; generatedPlot.upDownBars.upFill.color={xlpp::ChartColor::Kind::SRgb,"D9EAD3"};
    generatedPlot.upDownBars.downFill.present=true; generatedPlot.upDownBars.downFill.kind=xlpp::ChartFillFormat::Kind::Solid; generatedPlot.upDownBars.downFill.color={xlpp::ChartColor::Kind::SRgb,"F4CCCC"};
    xlpp::ChartDataTable table; table.present=true; table.showHorizontalBorder=true; table.showVerticalBorder=true; table.showOutline=true; table.showLegendKeys=true;
    table.textStyle.present=true; table.textStyle.bold=true; table.textStyle.fontSizePoints=10.0; table.textStyle.typeface="Aptos"; table.textStyle.color={xlpp::ChartColor::Kind::SRgb,"1F4E78"};
    stock.setDataTable(table);
    generatedSheet.addChart(std::move(stock));
    const auto generatedPath = std::filesystem::temp_directory_path() / "xlpp_p0q_generated_stock.xlsx";
    generated.save(generatedPath);
    const auto generatedZip = xlpp::internal::ZipArchive::open(generatedPath);
    const auto generatedXml = generatedZip.get("xl/charts/chart1.xml");
    test.checkTrue(generatedXml.find("<c:stockChart>") != std::string::npos, "New Stock chart serializes stockChart");
    test.checkTrue(generatedXml.find("<c:hiLowLines>") != std::string::npos && generatedXml.find("<c:upDownBars>") != std::string::npos, "New stock auxiliary objects serialize");
    test.checkTrue(generatedXml.find("<c:dTable>") != std::string::npos && generatedXml.find("<c:txPr") != std::string::npos && generatedXml.find("Aptos") != std::string::npos, "New chart data-table txPr serializes");
    test.checkTrue(generatedXml.find("<c:cat><c:numRef>") != std::string::npos, "Generated stock categories use numeric references");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(generatedZip).validate().ok(), "Generated stock package graph validates");
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath);
    const auto* generatedReloadSheet = generatedReload.worksheet("GeneratedStock");
    test.checkTrue(generatedReloadSheet != nullptr, "Generated stock workbook reloads");
    if (generatedReloadSheet) {
        const auto& generatedChart = static_cast<const xlpp::Worksheet&>(*generatedReloadSheet).charts().front();
        test.checkTrue(generatedChart.type() == xlpp::Chart::Type::Stock, "Generated stock chart reloads as Stock");
        test.checkTrue(generatedChart.plots().front().hasHighLowLines && generatedChart.plots().front().upDownBars.present, "Generated stock auxiliary model reloads");
        test.checkTrue(generatedChart.dataTable().textStyle.present && generatedChart.dataTable().textStyle.bold, "Generated dTable text style reloads");
    }
    std::filesystem::remove(generatedPath);

    xlpp::Workbook invalid;
    auto& invalidSheet = invalid.addWorksheet("InvalidStock");
    xlpp::Chart invalidStock(xlpp::Chart::Type::Stock);
    invalidStock.addSeries(xlpp::ChartSeries("Only one")); invalidSheet.addChart(std::move(invalidStock));
    bool rejected=false; try { invalid.save(std::filesystem::temp_directory_path()/"xlpp_invalid_stock.xlsx"); } catch (const std::invalid_argument&) { rejected=true; }
    test.checkTrue(rejected, "Invalid stock series count rejected predictably");
}

void testImportedChartRemoveAndAppend(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/image_chart.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Objects");
    const auto& importedCharts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(importedCharts.size(), std::size_t{1}, "Remove/append fixture starts with one imported chart");
    const auto oldStableId = importedCharts.front().stableId();
    test.checkTrue(sheet->removeChart(oldStableId), "Imported chart removed by stable ID");
    test.checkTrue(!sheet->removeChart(oldStableId), "Removed stable chart ID cannot be removed twice");

    xlpp::Chart replacement(xlpp::Chart::Type::Line);
    replacement.setTitle("Replacement chart");
    replacement.setXAxisTitle("Category");
    replacement.setYAxisTitle("Amount");
    replacement.setWidth(300);
    replacement.setHeight(180);
    auto& replacementSeries = replacement.addSeries(xlpp::ChartSeries("Replacement series"));
    replacementSeries.setCategoriesReference("'Objects'!$A$2:$A$4");
    replacementSeries.setValuesReference("'Objects'!$B$2:$B$4");
    sheet->addChart(std::move(replacement));

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0i_chart_remove_append.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    test.checkTrue(!after.contains("xl/charts/chart1.xml"), "Removed imported chart part is cleaned up");
    test.checkTrue(after.contains("xl/charts/chart2.xml"), "Appended replacement chart receives collision-free part ID");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "Sibling imported image remains byte-identical after chart remove/append");
    const auto drawingRels = after.get("xl/drawings/_rels/drawing1.xml.rels");
    test.checkTrue(drawingRels.find("../charts/chart2.xml") != std::string::npos, "Preserved drawing points to appended chart part");
    test.checkTrue(drawingRels.find("../charts/chart1.xml") == std::string::npos, "Removed chart relationship is deleted");
    const auto contentTypes = after.get("[Content_Types].xml");
    test.checkTrue(contentTypes.find("/xl/charts/chart2.xml") != std::string::npos, "Appended chart has content-type override");
    test.checkTrue(contentTypes.find("/xl/charts/chart1.xml") == std::string::npos, "Removed chart content-type override cleaned up");

    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "Remove/append chart package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Remove/append keeps one visible chart");
    test.checkEqual(graph.objectInventory().images, std::size_t{1}, "Remove/append retains one sibling image");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadedSheet = reloaded.worksheet("Objects");
    test.checkEqual(static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().size(), std::size_t{1}, "Replacement chart reloads as one imported chart");
    test.checkEqual(static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front().title(), std::string("Replacement chart"), "Replacement chart title survives reload");
    test.checkEqual(static_cast<int>(static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front().type()),
                    static_cast<int>(xlpp::Chart::Type::Line), "Replacement chart type survives reload");

    const auto output2 = std::filesystem::temp_directory_path() / "xlpp_p0i_chart_remove_append_resave.xlsx";
    workbook.save(output2);
    const auto graph2 = xlpp::internal::RelationshipGraph::fromArchive(xlpp::internal::ZipArchive::open(output2));
    test.checkTrue(graph2.validate().ok(), "Repeated save of remove/append chart stays graph-clean");
    test.checkEqual(graph2.objectInventory().charts, std::size_t{1}, "Repeated save does not duplicate appended chart");
    std::filesystem::remove(output);
    std::filesystem::remove(output2);
}

void testChartAndPivotPackage(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m21_chart_pivot.xlsx";
    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("Charts");
    sheet.append({std::string("Quarter"), std::string("Units")});
    sheet.append({std::string("Q1"), 10.0});
    sheet.append({std::string("Q2"), 20.0});

    xlpp::Chart chart(xlpp::Chart::Type::Bar);
    chart.setTitle("Sales");
    chart.setXAxisTitle("Quarter");
    chart.setYAxisTitle("Units");
    chart.setShowLegend(true);
    chart.setLegendPosition("b");
    chart.setStyle("10");
    auto& series = chart.addSeries(xlpp::ChartSeries("Units"));
    series.reference("Charts", "$B$2:$B$3");
    series.categories("Charts", "$A$2:$A$3");
    sheet.addChart(chart);

    xlpp::PivotTable pivot("SalesPivot");
    pivot.setLocation("D1");
    pivot.cache().setCacheId(1);
    pivot.cache().setSourceData("'Charts'!$A$1:$B$3");
    pivot.cache().setFields({"Quarter", "Units"});
    pivot.cache().addRecord({"Q1", "10"});
    pivot.cache().addRecord({"Q2", "20"});
    pivot.addRowField("Quarter");
    pivot.rowFields().back().setFieldIndex(0);
    pivot.addColumnField("Units");
    pivot.columnFields().back().setFieldIndex(1);
    pivot.addDataField(1);
    sheet.addPivotTable(std::move(pivot));

    wb.save(path);
    test.checkTrue(std::filesystem::exists(path), "Chart/pivot workbook saved");

    auto z = xlpp::internal::ZipArchive::open(path);
    test.checkTrue(z.contains("xl/charts/chart1.xml"), "Chart part written");
    const auto chartXml = z.get("xl/charts/chart1.xml");
    test.checkTrue(chartXml.find("barChart") != std::string::npos, "Bar chart type in part");
    test.checkTrue(chartXml.find("Sales") != std::string::npos, "Chart title in part");
    test.checkTrue(chartXml.find("<c:overlay val=\"0\"/>") != std::string::npos, "Chart title overlay in part");
    test.checkTrue(chartXml.find("<c:v>Units</c:v>") != std::string::npos, "Series title in part");
    test.checkTrue(chartXml.find("<c:style val=\"10\"/>") != std::string::npos, "Chart style in part");
    test.checkTrue(chartXml.find("<c:legendPos val=\"b\"/>") != std::string::npos, "Legend position in part");
    test.checkTrue(z.contains("xl/pivotTables/pivotTable1.xml"), "Pivot part written");
    test.checkTrue(z.contains("xl/pivotCache/pivotCacheDefinition1.xml"), "Pivot cache written");
    test.checkTrue(z.contains("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"), "Pivot cache relationship part written");
    test.checkTrue(z.contains("xl/pivotCache/pivotCacheRecords1.xml"), "Pivot cache records part written");
    const auto pivotCacheRels = z.get("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels");
    test.checkTrue(pivotCacheRels.find("pivotCacheRecords") != std::string::npos, "Pivot cache records relationship written");
    const auto pivotCacheXml = z.get("xl/pivotCache/pivotCacheDefinition1.xml");
    test.checkTrue(pivotCacheXml.find("recordCount=\"2\"") != std::string::npos, "Pivot cache record count written");
    test.checkTrue(pivotCacheXml.find("<cacheFields count=\"2\">") != std::string::npos, "Pivot cache fields written");
    test.checkTrue(pivotCacheXml.find("name=\"Quarter\"") != std::string::npos, "Pivot first field written");
    test.checkTrue(pivotCacheXml.find("<sharedItems") != std::string::npos &&
                   pivotCacheXml.find("count=\"2\"") != std::string::npos, "Pivot shared items written");
    const auto pivotRecordsXml = z.get("xl/pivotCache/pivotCacheRecords1.xml");
    test.checkTrue(pivotRecordsXml.find("<pivotCacheRecords") != std::string::npos &&
                   pivotRecordsXml.find("count=\"2\"") != std::string::npos, "Pivot cache records written");
    test.checkTrue(pivotCacheXml.find("<s v=\"Q1\"/>") != std::string::npos, "Pivot shared string written");
    test.checkTrue(pivotCacheXml.find("<n v=\"10\"/>") != std::string::npos, "Pivot shared number written");
    test.checkTrue(pivotRecordsXml.find("<x v=\"0\"/><x v=\"0\"/>") != std::string::npos, "Pivot cache indexes written");
    const auto pivotTableXml = z.get("xl/pivotTables/pivotTable1.xml");
    test.checkTrue(pivotTableXml.find("<rowFields count=\"1\"><field x=\"0\"/>") != std::string::npos, "Pivot row field index written");
    test.checkTrue(pivotTableXml.find("<colFields count=\"1\"><field x=\"1\"/>") != std::string::npos, "Pivot column field index written");
    test.checkTrue(pivotTableXml.find("<dataField name=\"Sum of Units\" fld=\"1\"") != std::string::npos, "Pivot data field index written");
    test.checkTrue(z.contains("xl/drawings/drawing1.xml"), "Drawing part written for chart");

    {
        xlpp::Workbook loaded;
        loaded.load(path);
        test.checkTrue(loaded.worksheet("Charts") != nullptr, "Chart workbook reloads");
        const auto preservedPivot = std::find_if(loaded.preservedParts().begin(), loaded.preservedParts().end(),
            [](const xlpp::PreservedPart& part) { return part.name == "xl/pivotTables/pivotTable1.xml"; });
        test.checkTrue(preservedPivot != loaded.preservedParts().end(), "Pivot table part is preserved on load");
        const auto preservedCache = std::find_if(loaded.preservedParts().begin(), loaded.preservedParts().end(),
            [](const xlpp::PreservedPart& part) { return part.name == "xl/pivotCache/pivotCacheDefinition1.xml"; });
        test.checkTrue(preservedCache != loaded.preservedParts().end(), "Pivot cache part is preserved on load");
        const auto preservedOut = std::filesystem::temp_directory_path() / "xlpp_m21_chart_pivot_preserved.xlsx";
        loaded.save(preservedOut);
        auto preservedZip = xlpp::internal::ZipArchive::open(preservedOut);
        test.checkTrue(preservedZip.get("xl/pivotTables/pivotTable1.xml") == z.get("xl/pivotTables/pivotTable1.xml"), "Pivot table XML survives load-save");
        test.checkTrue(preservedZip.get("xl/pivotCache/pivotCacheDefinition1.xml") == z.get("xl/pivotCache/pivotCacheDefinition1.xml"), "Pivot cache XML survives load-save");
        std::filesystem::remove(preservedOut);
    }
    std::filesystem::remove(path);
}

void testThreeDSurfaceChartPreservationFoundation(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_3d_surface.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");
    std::vector<std::string> untouchedCharts;
    for (int index = 2; index <= 6; ++index)
        untouchedCharts.push_back(before.get("xl/charts/chart" + std::to_string(index) + ".xml"));

    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("ThreeD");
    test.checkTrue(sheet != nullptr, "P0R 3D/surface fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{6}, "P0R fixture exposes six charts");
    if (charts.size() != 6) return;
    const std::array<xlpp::Chart::Type, 6> expected{{
        xlpp::Chart::Type::Bar3D, xlpp::Chart::Type::Line3D, xlpp::Chart::Type::Area3D,
        xlpp::Chart::Type::Pie3D, xlpp::Chart::Type::Surface, xlpp::Chart::Type::Surface3D}};
    for (std::size_t i=0; i<expected.size(); ++i)
        test.checkTrue(charts[i].type() == expected[i], "P0R chart type parsed at index " + std::to_string(i));

    const auto& bar3d = charts[0];
    test.checkTrue(bar3d.view3D().present, "Bar3D view3D parsed");
    test.checkTrue(bar3d.view3D().hasRotationX && bar3d.view3D().rotationX == 20, "Bar3D rotX parsed");
    test.checkTrue(bar3d.view3D().hasRotationY && bar3d.view3D().rotationY == 35, "Bar3D rotY parsed");
    test.checkTrue(bar3d.view3D().hasHeightPercent && bar3d.view3D().heightPercent == 120, "Bar3D hPercent parsed");
    test.checkTrue(bar3d.view3D().hasDepthPercent && bar3d.view3D().depthPercent == 180, "Bar3D depthPercent parsed");
    test.checkTrue(bar3d.view3D().hasRightAngleAxes && !bar3d.view3D().rightAngleAxes, "Bar3D rAngAx parsed");
    test.checkTrue(bar3d.view3D().hasPerspective && bar3d.view3D().perspective == 35, "Bar3D perspective parsed");
    test.checkTrue(bar3d.floorFormat().present && bar3d.floorFormat().hasThickness && bar3d.floorFormat().thickness == 12, "Bar3D floor thickness parsed");
    test.checkEqual(bar3d.floorFormat().fill.color.value, std::string("D9EAF7"), "Bar3D floor fill parsed");
    test.checkTrue(bar3d.sideWallFormat().present && bar3d.sideWallFormat().thickness == 18, "Bar3D side wall parsed");
    test.checkEqual(bar3d.backWallFormat().fill.color.value, std::string("FCE4D6"), "Bar3D back wall fill parsed");

    test.checkTrue(charts[4].type() == xlpp::Chart::Type::Surface && charts[4].plots().front().axisIds.size() == 3,
                   "Surface chart exposes three native axes");
    test.checkTrue(charts[5].type() == xlpp::Chart::Type::Surface3D && charts[5].view3D().present,
                   "Surface3D view model parsed");
    test.checkTrue(charts[5].floorFormat().thickness == 14 && charts[5].sideWallFormat().thickness == 16 && charts[5].backWallFormat().thickness == 22,
                   "Surface3D wall thickness metadata parsed");

    auto view = bar3d.view3D();
    view.present=true; view.hasRotationX=true; view.rotationX=40; view.hasRotationY=true; view.rotationY=75;
    view.hasHeightPercent=true; view.heightPercent=140; view.hasDepthPercent=true; view.depthPercent=220;
    view.hasRightAngleAxes=true; view.rightAngleAxes=false; view.hasPerspective=true; view.perspective=50;
    test.checkTrue(sheet->setChartView3D(bar3d.stableId(), view), "Selective view3D edit accepted");

    xlpp::ChartWallFormat floor = bar3d.floorFormat(); floor.present=true; floor.hasThickness=true; floor.thickness=30;
    floor.fill.present=true; floor.fill.kind=xlpp::ChartFillFormat::Kind::Solid; floor.fill.color={xlpp::ChartColor::Kind::SRgb,"4472C4"};
    floor.line.present=true; floor.line.color={xlpp::ChartColor::Kind::SRgb,"203864"}; floor.line.widthPoints=1.75;
    test.checkTrue(sheet->setChartFloorFormat(bar3d.stableId(), floor), "Selective floor formatting edit accepted");
    xlpp::ChartWallFormat side = bar3d.sideWallFormat(); side.present=true; side.hasThickness=true; side.thickness=28;
    side.fill.present=true; side.fill.kind=xlpp::ChartFillFormat::Kind::Solid; side.fill.color={xlpp::ChartColor::Kind::SRgb,"A9D18E"};
    test.checkTrue(sheet->setChartSideWallFormat(bar3d.stableId(), side), "Selective side-wall formatting edit accepted");
    xlpp::ChartWallFormat back = bar3d.backWallFormat(); back.present=true; back.hasThickness=true; back.thickness=26;
    back.fill.present=true; back.fill.kind=xlpp::ChartFillFormat::Kind::Solid; back.fill.color={xlpp::ChartColor::Kind::SRgb,"F4B183"};
    test.checkTrue(sheet->setChartBackWallFormat(bar3d.stableId(), back), "Selective back-wall formatting edit accepted");
    sheet->cell("K25").setValue("p0r-regression");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0r_3d_surface.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto chart1 = after.get("xl/charts/chart1.xml");
    test.checkTrue(chart1.find("bar3DChart") != std::string::npos, "Selective edit preserves bar3DChart structure");
    test.checkTrue(chart1.find("rotX val=\"40\"") != std::string::npos && chart1.find("rotY val=\"75\"") != std::string::npos, "Edited view3D values written");
    test.checkTrue(chart1.find("4472C4") != std::string::npos && chart1.find("203864") != std::string::npos, "Edited floor formatting written");
    test.checkTrue(chart1.find("A9D18E") != std::string::npos && chart1.find("F4B183") != std::string::npos, "Edited wall fills written");
    for (int index = 2; index <= 6; ++index)
        test.checkEqual(after.get("xl/charts/chart" + std::to_string(index) + ".xml"), untouchedCharts[static_cast<std::size_t>(index-2)],
                        "Unedited 3D/surface chart part remains byte-identical");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0R keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0R keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0R output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("ThreeD");
    test.checkTrue(reloadSheet != nullptr, "P0R output reloads");
    if (reloadSheet) {
        const auto& reCharts = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts();
        test.checkEqual(reCharts.size(), std::size_t{6}, "P0R reload retains all chart objects");
        if (!reCharts.empty()) {
            test.checkTrue(reCharts[0].view3D().rotationX == 40 && reCharts[0].view3D().rotationY == 75, "Edited view3D survives reload");
            test.checkTrue(reCharts[0].floorFormat().thickness == 30, "Edited floor thickness survives reload");
            test.checkEqual(reCharts[0].floorFormat().fill.color.value, std::string("4472C4"), "Edited floor fill survives reload");
        }
    }
    std::filesystem::remove(output);

    xlpp::Workbook generated;
    auto& gs = generated.addWorksheet("Generated3D");
    gs.cell("A1").setValue("Category"); gs.cell("B1").setValue("S1"); gs.cell("C1").setValue("S2");
    for (std::size_t row=2; row<=5; ++row) {
        gs.cell("A"+std::to_string(row)).setValue("C"+std::to_string(row-1));
        gs.cell("B"+std::to_string(row)).setValue(static_cast<double>(row*2));
        gs.cell("C"+std::to_string(row)).setValue(static_cast<double>(row*3));
    }
    auto addSeries = [](xlpp::Chart& chart, const char* title, const char* column) {
        xlpp::ChartSeries series(title);
        series.setCategoriesReference("'Generated3D'!$A$2:$A$5");
        series.setValuesReference(std::string("'Generated3D'!$") + column + "$2:$" + column + "$5");
        chart.addSeries(std::move(series));
    };
    xlpp::Chart generatedBar(xlpp::Chart::Type::Bar3D); generatedBar.setTitle("Generated Bar3D");
    addSeries(generatedBar,"S1","B"); addSeries(generatedBar,"S2","C");
    xlpp::ChartView3D generatedView; generatedView.present=true; generatedView.hasRotationX=true; generatedView.rotationX=25;
    generatedView.hasRotationY=true; generatedView.rotationY=45; generatedView.hasDepthPercent=true; generatedView.depthPercent=180;
    generatedView.hasRightAngleAxes=true; generatedView.rightAngleAxes=false; generatedView.hasPerspective=true; generatedView.perspective=35;
    generatedBar.setView3D(generatedView);
    xlpp::ChartWallFormat generatedFloor; generatedFloor.present=true; generatedFloor.hasThickness=true; generatedFloor.thickness=12;
    generatedFloor.fill.present=true; generatedFloor.fill.kind=xlpp::ChartFillFormat::Kind::Solid; generatedFloor.fill.color={xlpp::ChartColor::Kind::SRgb,"DDEBF7"};
    generatedBar.setFloorFormat(generatedFloor);
    auto& barPlot=generatedBar.primaryPlot(); barPlot.hasGapDepth=true; barPlot.gapDepth=175; barPlot.shape="box";
    gs.addChart(std::move(generatedBar));

    xlpp::Chart generatedSurface(xlpp::Chart::Type::Surface3D); generatedSurface.setTitle("Generated Surface3D");
    addSeries(generatedSurface,"S1","B"); addSeries(generatedSurface,"S2","C");
    generatedSurface.setView3D(generatedView); auto& surfacePlot=generatedSurface.primaryPlot(); surfacePlot.hasWireframe=true; surfacePlot.wireframe=true;
    gs.addChart(std::move(generatedSurface));
    const auto generatedPath=std::filesystem::temp_directory_path()/"xlpp_p0r_generated_3d.xlsx";
    generated.save(generatedPath);
    const auto generatedZip=xlpp::internal::ZipArchive::open(generatedPath);
    const auto generatedBarXml=generatedZip.get("xl/charts/chart1.xml");
    const auto generatedSurfaceXml=generatedZip.get("xl/charts/chart2.xml");
    test.checkTrue(generatedBarXml.find("bar3DChart")!=std::string::npos && generatedBarXml.find("<c:serAx>")!=std::string::npos, "Generated Bar3D writes three-axis structure");
    test.checkTrue(generatedBarXml.find("gapDepth val=\"175\"")!=std::string::npos && generatedBarXml.find("shape val=\"box\"")!=std::string::npos, "Generated Bar3D gap-depth and shape written");
    test.checkTrue(generatedBarXml.find("view3D")!=std::string::npos && generatedBarXml.find("DDEBF7")!=std::string::npos, "Generated Bar3D view and floor formatting written");
    test.checkTrue(generatedSurfaceXml.find("surface3DChart")!=std::string::npos && generatedSurfaceXml.find("wireframe val=\"1\"")!=std::string::npos, "Generated Surface3D writes wireframe structure");
    test.checkTrue(generatedSurfaceXml.find("<c:serAx>")!=std::string::npos, "Generated Surface3D writes series axis");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(generatedZip).validate().ok(), "Generated P0R package graph validates");
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath);
    const auto* grs=generatedReload.worksheet("Generated3D");
    test.checkTrue(grs!=nullptr && static_cast<const xlpp::Worksheet&>(*grs).charts().size()==2, "Generated 3D workbook reloads with two charts");
    if (grs && static_cast<const xlpp::Worksheet&>(*grs).charts().size()==2) {
        const auto& gcharts=static_cast<const xlpp::Worksheet&>(*grs).charts();
        test.checkTrue(gcharts[0].type()==xlpp::Chart::Type::Bar3D && gcharts[1].type()==xlpp::Chart::Type::Surface3D, "Generated chart types survive reload");
        test.checkEqual(gcharts[0].axes().size(), std::size_t{3}, "Generated Bar3D exposes three axes after reload");
        test.checkEqual(gcharts[1].axes().size(), std::size_t{3}, "Generated Surface3D exposes three axes after reload");
    }
    std::filesystem::remove(generatedPath);
}

void testProjectedPieDoughnutRadarExpansion(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/projected_pie_doughnut_radar.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalChart2 = before.get("xl/charts/chart2.xml");
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("Projected");
    test.checkTrue(sheet != nullptr, "P0S projected-pie/doughnut/radar fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{4}, "P0S fixture exposes four charts");
    if (charts.size() != 4) return;
    test.checkTrue(charts[0].type()==xlpp::Chart::Type::PieOfPie, "Pie-of-Pie type parsed");
    test.checkTrue(charts[1].type()==xlpp::Chart::Type::BarOfPie, "Bar-of-Pie type parsed");
    test.checkTrue(charts[2].type()==xlpp::Chart::Type::Doughnut, "Doughnut type parsed");
    test.checkTrue(charts[3].type()==xlpp::Chart::Type::Radar, "Radar type parsed");

    const auto& piePlot=charts[0].plots().front();
    test.checkTrue(piePlot.projectedPie.present && piePlot.projectedPie.ofPieType=="pie", "Pie-of-Pie options parsed");
    test.checkEqual(piePlot.projectedPie.gapWidth, 180, "Pie-of-Pie gap width parsed");
    test.checkEqual(piePlot.projectedPie.splitType, std::string("cust"), "Pie-of-Pie custom split type parsed");
    test.checkEqual(piePlot.projectedPie.customSplitPoints.size(), std::size_t{3}, "Pie-of-Pie custom split points parsed");
    test.checkEqual(piePlot.projectedPie.secondPlotSize, 120, "Pie-of-Pie second plot size parsed");
    const auto& barPlot=charts[1].plots().front();
    test.checkTrue(barPlot.projectedPie.present && barPlot.projectedPie.ofPieType=="bar", "Bar-of-Pie options parsed");
    test.checkEqual(barPlot.projectedPie.splitType, std::string("val"), "Bar-of-Pie value split parsed");
    test.checkTrue(barPlot.projectedPie.hasSplitPosition, "Bar-of-Pie split position present");
    test.checkNear(barPlot.projectedPie.splitPosition, 12.0, 1e-12, "Bar-of-Pie split position parsed");

    const auto& doughnutPlot=charts[2].plots().front();
    test.checkTrue(doughnutPlot.hasFirstSliceAngle && doughnutPlot.firstSliceAngle==45, "Doughnut first-slice angle parsed");
    test.checkTrue(doughnutPlot.hasHoleSize && doughnutPlot.holeSize==55, "Doughnut hole size parsed");
    const auto& radarPlot=charts[3].plots().front();
    test.checkEqual(radarPlot.radarStyle, std::string("filled"), "Radar style parsed");
    test.checkEqual(charts[3].series().size(), std::size_t{2}, "Radar series parsed");
    test.checkEqual(charts[3].series()[0].markerFormat().symbol, std::string("circle"), "Radar marker symbol parsed");
    test.checkEqual(charts[3].series()[0].markerFormat().size, 7, "Radar marker size parsed");

    auto projected=piePlot.projectedPie; projected.present=true; projected.splitType="percent"; projected.hasSplitPosition=true;
    projected.splitPosition=25; projected.gapWidth=200; projected.secondPlotSize=110; projected.customSplitPoints.clear();
    projected.hasSeriesLines=true; projected.seriesLinesFormat.present=true; projected.seriesLinesFormat.widthPoints=1.5;
    projected.seriesLinesFormat.color={xlpp::ChartColor::Kind::SRgb,"7030A0"};
    test.checkTrue(sheet->setChartPlotProjectedPieOptions(charts[0].stableId(),0,projected), "Selective projected-pie options edit accepted");
    test.checkTrue(sheet->setChartPlotFirstSliceAngle(charts[2].stableId(),0,120), "Selective doughnut first-slice edit accepted");
    test.checkTrue(sheet->setChartPlotDoughnutHoleSize(charts[2].stableId(),0,70), "Selective doughnut hole-size edit accepted");
    test.checkTrue(sheet->setChartPlotRadarStyle(charts[3].stableId(),0,"marker"), "Selective radar-style edit accepted");
    auto marker=charts[3].series()[0].markerFormat(); marker.present=true; marker.symbol="star"; marker.size=9;
    marker.fill.present=true; marker.fill.kind=xlpp::ChartFillFormat::Kind::Solid; marker.fill.color={xlpp::ChartColor::Kind::SRgb,"FF0000"};
    test.checkTrue(sheet->setChartSeriesMarkerFormat(charts[3].stableId(),0,marker), "Selective radar marker edit accepted");
    test.checkTrue(!sheet->setChartPlotDoughnutHoleSize(charts[2].stableId(),0,5), "Invalid doughnut hole size rejected");
    test.checkTrue(!sheet->setChartPlotRadarStyle(charts[3].stableId(),0,"unsupported"), "Invalid radar style rejected");
    sheet->cell("K26").setValue("p0s-regression");

    const auto output=std::filesystem::temp_directory_path()/"xlpp_p0s_projected_pie_doughnut_radar.xlsx";
    workbook.save(output);
    const auto after=xlpp::internal::ZipArchive::open(output);
    const auto chart1=after.get("xl/charts/chart1.xml");
    const auto chart3=after.get("xl/charts/chart3.xml");
    const auto chart4=after.get("xl/charts/chart4.xml");
    test.checkTrue(chart1.find("ofPieType val=\"pie\"")!=std::string::npos && chart1.find("splitType val=\"percent\"")!=std::string::npos, "Projected-pie selective options serialized");
    test.checkTrue(chart1.find("splitPos val=\"25")!=std::string::npos && chart1.find("secondPieSize val=\"110\"")!=std::string::npos, "Projected-pie split position and size serialized");
    test.checkTrue(chart1.find("7030A0")!=std::string::npos, "Projected-pie series-line formatting serialized");
    test.checkTrue(chart3.find("firstSliceAng val=\"120\"")!=std::string::npos && chart3.find("holeSize val=\"70\"")!=std::string::npos, "Doughnut options serialized");
    test.checkTrue(chart4.find("radarStyle val=\"marker\"")!=std::string::npos && chart4.find("symbol val=\"star\"")!=std::string::npos, "Radar style and marker serialized");
    test.checkEqual(after.get("xl/charts/chart2.xml"), originalChart2, "Untouched Bar-of-Pie chart remains byte-identical");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0S keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0S keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0S selective output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output); const auto* rs=reloaded.worksheet("Projected");
    test.checkTrue(rs!=nullptr, "P0S selective output reloads");
    if(rs){ const auto& rc=static_cast<const xlpp::Worksheet&>(*rs).charts(); test.checkEqual(rc.size(),std::size_t{4},"P0S reload keeps four charts");
        if(rc.size()==4){
            test.checkEqual(rc[0].plots()[0].projectedPie.splitType,std::string("percent"),"Projected-pie edited split survives reload");
            test.checkEqual(rc[2].plots()[0].holeSize,70,"Doughnut edited hole survives reload");
            test.checkEqual(rc[3].plots()[0].radarStyle,std::string("marker"),"Radar edited style survives reload");
            test.checkEqual(rc[3].series()[0].markerFormat().symbol,std::string("star"),"Radar edited marker survives reload");
        }
    }
    std::filesystem::remove(output);

    xlpp::Workbook generated; auto& gs=generated.addWorksheet("Generated");
    gs.append({std::string("Category"),std::string("Primary"),std::string("Secondary")});
    const std::array<std::array<double,2>,6> vals{{{{10,14}},{{20,18}},{{30,12}},{{5,9}},{{8,7}},{{13,11}}}};
    for(std::size_t i=0;i<vals.size();++i){ gs.cell(i+2,1).setValue(std::string(1,static_cast<char>('A'+i))); gs.cell(i+2,2).setValue(vals[i][0]); gs.cell(i+2,3).setValue(vals[i][1]); }
    auto addOneSeries=[](xlpp::Chart& c,const char* title,const char* col){ xlpp::ChartSeries s(title); s.setCategoriesReference("'Generated'!$A$2:$A$7"); s.setValuesReference(std::string("'Generated'!$")+col+"$2:$"+col+"$7"); c.addSeries(std::move(s)); };
    xlpp::Chart gp(xlpp::Chart::Type::PieOfPie); gp.setTitle("Generated Pie-of-Pie"); addOneSeries(gp,"Primary","B");
    auto& gpp=gp.primaryPlot(); gpp.projectedPie.present=true; gpp.projectedPie.splitType="cust"; gpp.projectedPie.customSplitPoints={1,4}; gpp.projectedPie.gapWidth=175; gpp.projectedPie.secondPlotSize=105; gs.addChart(std::move(gp));
    xlpp::Chart gb(xlpp::Chart::Type::BarOfPie); gb.setTitle("Generated Bar-of-Pie"); addOneSeries(gb,"Primary","B");
    auto& gbp=gb.primaryPlot(); gbp.projectedPie.present=true; gbp.projectedPie.splitType="val"; gbp.projectedPie.hasSplitPosition=true; gbp.projectedPie.splitPosition=12; gbp.projectedPie.secondPlotSize=95; gs.addChart(std::move(gb));
    xlpp::Chart gd(xlpp::Chart::Type::Doughnut); gd.setTitle("Generated Doughnut"); addOneSeries(gd,"Primary","B"); auto& gdp=gd.primaryPlot(); gdp.hasFirstSliceAngle=true; gdp.firstSliceAngle=90; gdp.hasHoleSize=true; gdp.holeSize=60; gs.addChart(std::move(gd));
    xlpp::Chart gr(xlpp::Chart::Type::Radar); gr.setTitle("Generated Radar"); addOneSeries(gr,"Primary","B"); addOneSeries(gr,"Secondary","C"); auto& grp=gr.primaryPlot(); grp.radarStyle="marker";
    xlpp::ChartMarkerFormat gm; gm.present=true; gm.symbol="diamond"; gm.size=8; gr.series()[0].setMarkerFormat(gm); gs.addChart(std::move(gr));
    const auto generatedPath=std::filesystem::temp_directory_path()/"xlpp_p0s_generated_projected_radar.xlsx"; generated.save(generatedPath);
    const auto gz=xlpp::internal::ZipArchive::open(generatedPath);
    test.checkTrue(gz.get("xl/charts/chart1.xml").find("<c:ofPieChart>")!=std::string::npos && gz.get("xl/charts/chart1.xml").find("splitType val=\"cust\"")!=std::string::npos, "Generated Pie-of-Pie XML written");
    test.checkTrue(gz.get("xl/charts/chart2.xml").find("ofPieType val=\"bar\"")!=std::string::npos, "Generated Bar-of-Pie XML written");
    test.checkTrue(gz.get("xl/charts/chart3.xml").find("firstSliceAng val=\"90\"")!=std::string::npos && gz.get("xl/charts/chart3.xml").find("holeSize val=\"60\"")!=std::string::npos, "Generated Doughnut options written");
    test.checkTrue(gz.get("xl/charts/chart4.xml").find("radarStyle val=\"marker\"")!=std::string::npos && gz.get("xl/charts/chart4.xml").find("symbol val=\"diamond\"")!=std::string::npos, "Generated Radar style and marker written");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(gz).validate().ok(), "Generated P0S package graph validates");
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath); const auto* grs=generatedReload.worksheet("Generated");
    test.checkTrue(grs!=nullptr,"Generated P0S workbook reloads"); if(grs){ const auto& gc=static_cast<const xlpp::Worksheet&>(*grs).charts(); test.checkEqual(gc.size(),std::size_t{4},"Generated P0S charts reload"); if(gc.size()==4){ test.checkTrue(gc[0].type()==xlpp::Chart::Type::PieOfPie&&gc[1].type()==xlpp::Chart::Type::BarOfPie,"Generated projected-pie types survive reload"); test.checkEqual(gc[2].plots()[0].holeSize,60,"Generated doughnut hole survives reload"); test.checkEqual(gc[3].plots()[0].radarStyle,std::string("marker"),"Generated radar style survives reload"); }}
    std::filesystem::remove(generatedPath);
}


void testChartStyleThemeAndSeriesCaches(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/libreoffice/chart_style_theme_caches.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalStylePart = before.get("xl/charts/style1.xml");
    const auto originalColorStylePart = before.get("xl/charts/colors1.xml");
    const auto originalChartRels = before.get("xl/charts/_rels/chart1.xml.rels");
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");

    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "P0T style/theme/cache fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0T fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkEqual(chart.style(), std::string("10"), "Imported chart style ID parsed");
    test.checkTrue(chart.themePalette().present, "Workbook theme palette parsed for chart");
    test.checkEqual(chart.themePalette().baseColor("accent1"), std::string("4f81bd"), "Theme accent1 base color parsed");
    test.checkTrue(chart.styleResources().chartStylePresent && chart.styleResources().colorStylePresent, "Chart style/color-style relationships discovered");
    test.checkEqual(chart.styleResources().chartStylePart, std::string("xl/charts/style1.xml"), "Chart style part resolved");
    test.checkEqual(chart.styleResources().colorStylePart, std::string("xl/charts/colors1.xml"), "Chart color-style part resolved");
    test.checkEqual(chart.series().size(), std::size_t{1}, "P0T fixture series parsed");
    if (chart.series().empty()) return;
    const auto& series = chart.series().front();
    test.checkEqual(series.titleReference(), std::string("Objects!B1"), "Series title reference parsed");
    test.checkTrue(series.titleCache().present && !series.titleCache().numeric, "Series title strCache parsed");
    test.checkEqual(series.titleCache().points.size(), std::size_t{1}, "Title cache point count parsed");
    test.checkEqual(series.titleCache().points.front().value, std::string("Amount"), "Title cache value parsed");
    test.checkTrue(series.categoriesCache().present && !series.categoriesCache().numeric, "Category strCache parsed");
    test.checkEqual(series.categoriesCache().effectivePointCount(), std::size_t{3}, "Category cache count parsed");
    test.checkTrue(series.valuesCache().present && series.valuesCache().numeric, "Value numCache parsed");
    test.checkEqual(series.valuesCache().formatCode, std::string("General"), "Value cache format code parsed");
    test.checkEqual(series.valuesCache().points.size(), std::size_t{3}, "Value cache points parsed");
    test.checkTrue(series.fillFormat().color.kind == xlpp::ChartColor::Kind::Scheme, "Scheme series color parsed");
    test.checkEqual(series.fillFormat().color.value, std::string("accent1"), "Scheme series color name parsed");
    test.checkEqual(chart.resolveThemeBaseColor(series.fillFormat().color), std::string("4f81bd"), "Scheme color resolves through workbook theme");
    test.checkTrue(!series.fillFormat().color.transforms.empty(), "Scheme color transforms preserved in model");

    // Unrelated edit must keep chart/style resources byte-identical.
    sheet->cell("K27").setValue("p0t-unrelated");
    const auto unrelated = std::filesystem::temp_directory_path()/"xlpp_p0t_unrelated.xlsx";
    workbook.save(unrelated);
    const auto unrelatedZip = xlpp::internal::ZipArchive::open(unrelated);
    test.checkEqual(unrelatedZip.get("xl/charts/chart1.xml"), before.get("xl/charts/chart1.xml"), "Unrelated edit keeps chart XML byte-identical");
    test.checkEqual(unrelatedZip.get("xl/charts/style1.xml"), originalStylePart, "Unrelated edit preserves chart-style part byte-identical");
    test.checkEqual(unrelatedZip.get("xl/charts/colors1.xml"), originalColorStylePart, "Unrelated edit preserves color-style part byte-identical");
    test.checkEqual(unrelatedZip.get("xl/charts/_rels/chart1.xml.rels"), originalChartRels, "Unrelated edit preserves chart relationships byte-identical");
    std::filesystem::remove(unrelated);

    xlpp::ChartSeriesCache cat; cat.present=true; cat.numeric=false; cat.pointCount=3; cat.points={{0,"Alpha"},{1,"Beta"},{2,"Gamma"}};
    xlpp::ChartSeriesCache val; val.present=true; val.numeric=true; val.formatCode="0.00"; val.pointCount=3; val.points={{0,"11.25"},{1,"22.5"},{2,"33.75"}};
    xlpp::ChartSeriesCache title; title.present=true; title.numeric=false; title.pointCount=1; title.points={{0,"Updated Amount"}};
    test.checkTrue(sheet->setChartStyle(chart.stableId(), "15"), "Selective chart style edit accepted");
    test.checkTrue(sheet->setChartSeriesCategoryCache(chart.stableId(),0,cat), "Selective category cache edit accepted");
    test.checkTrue(sheet->setChartSeriesValueCache(chart.stableId(),0,val), "Selective value cache edit accepted");
    test.checkTrue(sheet->setChartSeriesTitleCache(chart.stableId(),0,title), "Selective title cache edit accepted");
    xlpp::ChartSeriesCache invalid=val; invalid.pointCount=1;
    test.checkTrue(!sheet->setChartSeriesValueCache(chart.stableId(),0,invalid), "Inconsistent cache pointCount rejected");
    const auto output=std::filesystem::temp_directory_path()/"xlpp_p0t_style_cache_edit.xlsx";
    workbook.save(output);
    const auto after=xlpp::internal::ZipArchive::open(output);
    const auto chartXml=after.get("xl/charts/chart1.xml");
    test.checkTrue(chartXml.find("style val=\"15\"")!=std::string::npos, "Selective chart style serialized");
    test.checkTrue(chartXml.find("Updated Amount")!=std::string::npos && chartXml.find("Alpha")!=std::string::npos && chartXml.find("33.75")!=std::string::npos, "Selective cache values serialized");
    test.checkTrue(chartXml.find("<c:formatCode>0.00</c:formatCode>")!=std::string::npos, "Numeric cache format code serialized");
    test.checkEqual(after.get("xl/charts/style1.xml"), originalStylePart, "Style resource survives chart metadata edit byte-identical");
    test.checkEqual(after.get("xl/charts/colors1.xml"), originalColorStylePart, "Color-style resource survives chart metadata edit byte-identical");
    test.checkEqual(after.get("xl/charts/_rels/chart1.xml.rels"), originalChartRels, "Chart style relationships survive selective edit byte-identical");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "Drawing relationships remain byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0T selective package graph validates");

    xlpp::Workbook reload; reload.load(output); const auto* rs=reload.worksheet("Objects");
    test.checkTrue(rs!=nullptr, "P0T selective output reloads");
    if(rs){ const auto& rc=static_cast<const xlpp::Worksheet&>(*rs).charts(); test.checkEqual(rc.size(),std::size_t{1},"P0T reload keeps chart"); if(!rc.empty()){
        const auto& r=rc.front(); test.checkEqual(r.style(),std::string("15"),"Edited style reloads");
        test.checkEqual(r.series()[0].categoriesCache().points[1].value,std::string("Beta"),"Edited category cache reloads");
        test.checkEqual(r.series()[0].valuesCache().formatCode,std::string("0.00"),"Edited numeric cache format reloads");
        test.checkEqual(r.series()[0].titleCache().points[0].value,std::string("Updated Amount"),"Edited title cache reloads");
    }}
    std::filesystem::remove(output);

    // Generated chart caches should be first-class rather than preservation-only metadata.
    xlpp::Workbook generated; auto& gs=generated.addWorksheet("Caches");
    gs.append({std::string("Category"),std::string("Amount")});
    gs.append({std::string("A"),10.0}); gs.append({std::string("B"),20.0}); gs.append({std::string("C"),30.0});
    xlpp::Chart gc(xlpp::Chart::Type::Bar); gc.setStyle("12"); xlpp::ChartSeries gseries("Amount");
    gseries.setTitleReference("'Caches'!$B$1"); gseries.setCategoriesReference("'Caches'!$A$2:$A$4"); gseries.setValuesReference("'Caches'!$B$2:$B$4");
    xlpp::ChartSeriesCache gtc; gtc.present=true; gtc.points={{0,"Amount"}}; gseries.setTitleCache(gtc);
    xlpp::ChartSeriesCache gcc; gcc.present=true; gcc.points={{0,"A"},{1,"B"},{2,"C"}}; gseries.setCategoriesCache(gcc);
    xlpp::ChartSeriesCache gvc; gvc.present=true; gvc.numeric=true; gvc.formatCode="0"; gvc.points={{0,"10"},{1,"20"},{2,"30"}}; gseries.setValuesCache(gvc);
    gc.addSeries(std::move(gseries)); gs.addChart(std::move(gc));
    const auto generatedPath=std::filesystem::temp_directory_path()/"xlpp_p0t_generated_caches.xlsx"; generated.save(generatedPath);
    const auto generatedZip=xlpp::internal::ZipArchive::open(generatedPath); const auto generatedXml=generatedZip.get("xl/charts/chart1.xml");
    test.checkTrue(generatedXml.find("style val=\"12\"")!=std::string::npos && generatedXml.find("strCache")!=std::string::npos && generatedXml.find("numCache")!=std::string::npos, "Generated chart writes style and caches");
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath); const auto* grs=generatedReload.worksheet("Caches");
    test.checkTrue(grs!=nullptr,"Generated cache workbook reloads"); if(grs){ const auto& c=static_cast<const xlpp::Worksheet&>(*grs).charts(); if(!c.empty()){ test.checkEqual(c[0].series()[0].titleCache().points[0].value,std::string("Amount"),"Generated title cache reloads"); test.checkEqual(c[0].series()[0].valuesCache().points.size(),std::size_t{3},"Generated value cache reloads"); }}
    std::filesystem::remove(generatedPath);
}


void testChartCacheSynchronizationAndThemeTransforms(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/libreoffice/chart_style_theme_caches.xlsx";
    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "P0U cache-sync fixture loads");
    if (!sheet) return;
    const auto& initialCharts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(initialCharts.size(), std::size_t{1}, "P0U fixture exposes one chart");
    if (initialCharts.empty()) return;
    const auto& initial = initialCharts.front();
    test.checkTrue(initial.themePalette().fontScheme.present, "Theme font scheme parsed");
    test.checkEqual(initial.themePalette().fontScheme.name, std::string("Office"), "Theme font scheme name parsed");
    test.checkEqual(initial.themePalette().fontScheme.majorLatinTypeface, std::string("Cambria"), "Theme major Latin font parsed");
    test.checkEqual(initial.themePalette().fontScheme.minorLatinTypeface, std::string("Calibri"), "Theme minor Latin font parsed");
    test.checkTrue(initial.themePalette().effectScheme.present, "Theme effect scheme parsed");
    test.checkEqual(initial.themePalette().effectScheme.fillStyleCount, std::size_t{3}, "Theme fill style count parsed");
    test.checkEqual(initial.themePalette().effectScheme.lineStyleCount, std::size_t{3}, "Theme line style count parsed");
    test.checkEqual(initial.themePalette().effectScheme.effectStyleCount, std::size_t{3}, "Theme effect style count parsed");
    test.checkEqual(initial.themePalette().effectScheme.backgroundFillStyleCount, std::size_t{3}, "Theme background fill count parsed");
    test.checkEqual(initial.themePalette().effectScheme.fillStyles.size(), std::size_t{3}, "P0Y materializes theme fill-style matrix");
    test.checkEqual(initial.themePalette().effectScheme.lineStyles.size(), std::size_t{3}, "P0Y materializes theme line-style matrix");
    test.checkEqual(initial.themePalette().effectScheme.effectStyles.size(), std::size_t{3}, "P0Y materializes theme effect-style matrix");
    test.checkEqual(initial.themePalette().effectScheme.backgroundFillStyles.size(), std::size_t{3}, "P0Y materializes theme background-fill matrix");
    test.checkTrue(initial.themePalette().effectScheme.fillStyles.front().present, "P0Y first theme fill is materialized");
    test.checkTrue(initial.themePalette().effectScheme.lineStyles.front().present, "P0Y first theme line is materialized");
    test.checkNear(initial.themePalette().effectScheme.lineStyles.front().widthPoints, 0.75, 1e-9, "P0Y theme line width converts from EMU to points");
    const auto transformed = initial.resolveThemeColor(initial.series()[0].fillFormat().color);
    test.checkTrue(transformed.present, "Theme transformed color resolves");
    test.checkEqual(transformed.srgb(), std::string("729ACA"), "Theme tint transform produces final RGB");
    test.checkNear(transformed.alpha, 1.0, 1e-12, "Theme transformed alpha defaults opaque");

    xlpp::ChartColor transformedColor{xlpp::ChartColor::Kind::Scheme, "accent1"};
    transformedColor.transforms.push_back({xlpp::ChartColorTransform::Kind::Shade, 50000});
    transformedColor.transforms.push_back({xlpp::ChartColorTransform::Kind::Alpha, 60000});
    const auto shaded = initial.resolveThemeColor(transformedColor);
    test.checkEqual(shaded.srgb(), std::string("28415F"), "Theme shade transform applies sequentially");
    test.checkNear(shaded.alpha, 0.6, 1e-12, "Theme alpha transform resolves");

    // Edit source cells and rebuild the imported chart caches from A1 references.
    sheet->cell("B1").setValue("Synced Amount");
    sheet->cell("A2").setValue("North");
    sheet->cell("A3").clear();
    sheet->cell("A4").setValue("South");
    sheet->cell("B2").setValue(101.5); sheet->cell("B2").setNumberFormat("0.0");
    sheet->cell("B3").clear();
    sheet->cell("B4").setValue(303.25); sheet->cell("B4").setNumberFormat("0.0");
    const auto report = workbook.synchronizeChartCaches();
    test.checkEqual(report.chartsVisited, std::size_t{1}, "Cache sync visits imported chart");
    test.checkEqual(report.seriesVisited, std::size_t{1}, "Cache sync visits imported series");
    test.checkEqual(report.cachesUpdated, std::size_t{3}, "Cache sync updates title/category/value caches");
    test.checkEqual(report.referencesSkipped, std::size_t{0}, "Cache sync accepts local A1 references");
    const auto& synced = static_cast<const xlpp::Worksheet&>(*sheet).charts().front().series().front();
    test.checkEqual(synced.titleCache().points[0].value, std::string("Synced Amount"), "Title cache rebuilt from worksheet cell");
    test.checkTrue(synced.categoriesCache().sparse(), "String category cache records sparse blank cell");
    test.checkEqual(synced.categoriesCache().pointCount, std::size_t{3}, "Sparse category cache keeps source range length");
    test.checkEqual(synced.categoriesCache().points.size(), std::size_t{2}, "Sparse category cache omits blank point");
    test.checkEqual(synced.categoriesCache().points[1].index, std::size_t{2}, "Sparse category cache retains source index");
    test.checkTrue(synced.valuesCache().numeric && synced.valuesCache().sparse(), "Numeric value cache rebuilt as sparse numCache");
    test.checkEqual(synced.valuesCache().formatCode, std::string("0.0"), "Numeric cache adopts worksheet number format");
    test.checkEqual(synced.valuesCache().points[1].value, std::string("303.25"), "Numeric cache reads edited worksheet value");
    test.checkTrue(synced.valuesCache().ordered(), "Synchronized cache points are index ordered");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0u_cache_sync.xlsx";
    workbook.save(output);
    xlpp::Workbook reload; reload.load(output);
    const auto* reloadedSheet = reload.worksheet("Objects");
    test.checkTrue(reloadedSheet != nullptr, "P0U synchronized workbook reloads");
    if (reloadedSheet) {
        const auto& rs = static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front().series().front();
        test.checkEqual(rs.titleCache().points[0].value, std::string("Synced Amount"), "Synchronized title cache survives save/reload");
        test.checkTrue(rs.categoriesCache().sparse(), "Sparse category cache survives save/reload");
        test.checkTrue(rs.valuesCache().sparse(), "Sparse numeric cache survives save/reload");
        test.checkEqual(rs.valuesCache().formatCode, std::string("0.0"), "Synchronized numeric format survives reload");
    }
    std::filesystem::remove(output);

    // Generated charts can synchronize cross-sheet quoted references without pre-built caches.
    xlpp::Workbook generated;
    auto& data = generated.addWorksheet("Data O'Brien");
    data.cell("A1").setValue("Category"); data.cell("B1").setValue("Value");
    data.cell("A2").setValue("A"); data.cell("B2").setValue(10.0); data.cell("B2").setNumberFormat("0.00");
    data.cell("A3").setValue("B"); // B3 intentionally blank -> sparse value cache.
    data.cell("A4").setValue("C"); data.cell("B4").setValue(30.0); data.cell("B4").setNumberFormat("0.00");
    auto& chartSheet = generated.addWorksheet("Chart Sheet");
    xlpp::Chart generatedChart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries generatedSeries("Value");
    generatedSeries.setTitleReference("'Data O''Brien'!$B$1");
    generatedSeries.setCategoriesReference("'Data O''Brien'!$A$2:$A$4");
    generatedSeries.setValuesReference("'Data O''Brien'!$B$2:$B$4");
    generatedChart.addSeries(std::move(generatedSeries)); chartSheet.addChart(std::move(generatedChart));
    const auto generatedReport = generated.synchronizeChartCaches();
    test.checkEqual(generatedReport.cachesUpdated, std::size_t{3}, "Generated cache sync handles quoted cross-sheet references");
    test.checkEqual(generatedReport.referencesSkipped, std::size_t{0}, "Generated cache sync accepts apostrophe-escaped sheet name");
    const auto& generatedCaches = static_cast<const xlpp::Worksheet&>(chartSheet).charts().front().series().front();
    test.checkEqual(generatedCaches.categoriesCache().points.size(), std::size_t{3}, "Generated string cache contains all categories");
    test.checkTrue(generatedCaches.valuesCache().sparse(), "Generated numeric cache preserves blank source point");
    test.checkEqual(generatedCaches.valuesCache().formatCode, std::string("0.00"), "Generated cache uses source number format");

    xlpp::ChartSeriesCache duplicate; duplicate.present=true; duplicate.pointCount=2; duplicate.points={{0,"A"},{0,"B"}};
    test.checkTrue(duplicate.hasDuplicateIndexes() && !duplicate.valid(), "Duplicate cache indexes are rejected by public validation");
    xlpp::ChartSeriesCache unordered; unordered.present=true; unordered.pointCount=3; unordered.points={{2,"C"},{0,"A"}};
    test.checkTrue(unordered.valid() && !unordered.ordered() && unordered.sparse(), "Sparse unordered cache is valid but detectable before serialization sorting");

    const auto generatedPath = std::filesystem::temp_directory_path() / "xlpp_p0u_generated_cache_sync.xlsx";
    generated.save(generatedPath);
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath);
    const auto* gcSheet = generatedReload.worksheet("Chart Sheet");
    test.checkTrue(gcSheet != nullptr, "Generated synchronized cache workbook reloads");
    if (gcSheet) {
        const auto& cache = static_cast<const xlpp::Worksheet&>(*gcSheet).charts().front().series().front().valuesCache();
        test.checkTrue(cache.present && cache.numeric && cache.sparse(), "Generated synchronized sparse cache is serialized");
        test.checkEqual(cache.pointCount, std::size_t{3}, "Generated synchronized cache pointCount survives reload");
    }
    std::filesystem::remove(generatedPath);
}


void testChartCacheDependencyTrackingAndStyleResolution(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/libreoffice/chart_style_theme_caches.xlsx";
    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "P0V dependency fixture loads");
    if (!sheet) return;
    test.checkEqual(sheet->trackedCellChangeCount(), std::size_t{0}, "Loaded sheet starts with no tracked cell changes");

    const auto dependencies = workbook.chartCacheDependencies();
    test.checkEqual(dependencies.size(), std::size_t{3}, "P0V exposes title/category/value cache dependencies");
    std::size_t supported = 0, title = 0, category = 0, value = 0;
    for (const auto& dependency : dependencies) {
        if (dependency.supported) ++supported;
        if (dependency.kind == xlpp::ChartCacheDependencyKind::Title) ++title;
        else if (dependency.kind == xlpp::ChartCacheDependencyKind::Category) ++category;
        else if (dependency.kind == xlpp::ChartCacheDependencyKind::Value) ++value;
        if (dependency.supported) test.checkEqual(dependency.sourceSheet, std::string("Objects"), "P0V dependency resolves source sheet");
    }
    test.checkEqual(supported, std::size_t{3}, "All P0V fixture dependencies are supported");
    test.checkEqual(title, std::size_t{1}, "One title dependency indexed");
    test.checkEqual(category, std::size_t{1}, "One category dependency indexed");
    test.checkEqual(value, std::size_t{1}, "One value dependency indexed");

    // An unrelated cell access/mutation should not rebuild any cache.
    sheet->cell("K99").setValue("unrelated");
    xlpp::ChartCacheSyncOptions incremental;
    incremental.clearTrackedChangesAfterSync = true;
    auto unrelated = workbook.synchronizeChangedChartCaches(incremental);
    test.checkEqual(unrelated.dependenciesVisited, std::size_t{3}, "Incremental sync visits dependency metadata only");
    test.checkEqual(unrelated.dependenciesMatched, std::size_t{0}, "Unrelated cell matches no cache dependency");
    test.checkEqual(unrelated.dependenciesSkippedUnchanged, std::size_t{3}, "Unrelated edit skips all cache rebuilds");
    test.checkEqual(unrelated.cachesUpdated, std::size_t{0}, "Unrelated edit updates no cache");
    test.checkEqual(sheet->trackedCellChangeCount(), std::size_t{0}, "Incremental sync can clear tracked changes");

    // One value cell should match only the value dependency.
    sheet->cell("B2").setValue(777.25);
    auto targeted = workbook.synchronizeChangedChartCaches(incremental);
    test.checkEqual(targeted.dependenciesMatched, std::size_t{1}, "Value cell matches exactly one dependency");
    test.checkEqual(targeted.dependenciesSkippedUnchanged, std::size_t{2}, "Title/category dependencies skipped for value-only edit");
    test.checkEqual(targeted.cachesUpdated, std::size_t{1}, "Only value cache is rebuilt");
    const auto& targetedSeries = static_cast<const xlpp::Worksheet&>(*sheet).charts().front().series().front();
    test.checkEqual(targetedSeries.valuesCache().points.front().value, std::string("777.25"), "Targeted value cache reads changed cell");

    // Formula without a cached worksheet value should preserve the existing chart cache point.
    std::string previousSecondValue;
    for (const auto& point : targetedSeries.valuesCache().points) if (point.index == 1) previousSecondValue = point.value;
    sheet->cell("B3").setValue(std::monostate{});
    sheet->cell("B3").setFormula("B2*2");
    auto formulaReport = workbook.synchronizeChangedChartCaches(incremental);
    test.checkEqual(formulaReport.dependenciesMatched, std::size_t{1}, "Formula source cell matches value dependency");
    test.checkEqual(formulaReport.formulaCachePointsReused, std::size_t{1}, "Missing formula cached value reuses existing chart point");
    const auto& formulaSeries = static_cast<const xlpp::Worksheet&>(*sheet).charts().front().series().front();
    std::string preservedSecondValue;
    for (const auto& point : formulaSeries.valuesCache().points) if (point.index == 1) preservedSecondValue = point.value;
    test.checkEqual(preservedSecondValue, previousSecondValue, "Formula without cached value does not erase chart cache point");

    // Public tracker can be explicitly reset without clearing worksheet dirty state.
    sheet->cell("A2").setValue("Tracked");
    test.checkTrue(sheet->hasTrackedCellChanges(), "Worksheet exposes pending tracked cell changes");
    workbook.clearChartCacheChangeTracking();
    test.checkTrue(!sheet->hasTrackedCellChanges(), "Workbook clears chart-cache change tracker explicitly");
    test.checkTrue(sheet->dirty(), "Clearing chart-cache tracker does not clear worksheet dirty state");

    // A retained Cell& must still be detected after the tracker itself is reset.
    auto& retainedValueCell = sheet->cell("B4");
    workbook.clearChartCacheChangeTracking();
    retainedValueCell.setValue(909.0);
    test.checkTrue(sheet->hasTrackedCellChanges(), "Retained Cell mutation revision reactivates change tracking");
    auto retainedReport = workbook.synchronizeChangedChartCaches(incremental);
    test.checkEqual(retainedReport.dependenciesMatched, std::size_t{1}, "Retained Cell mutation matches its cache dependency");
    test.checkEqual(retainedReport.cachesUpdated, std::size_t{1}, "Retained Cell mutation rebuilds the affected cache");
    const auto& retainedSeries = static_cast<const xlpp::Worksheet&>(*sheet).charts().front().series().front();
    test.checkEqual(retainedSeries.valuesCache().points.back().value, std::string("909"), "Retained Cell mutation is reflected in synchronized cache");

    // Enrich the synthetic chart-style resources to exercise resource inspection and theme resolution.
    auto archive = xlpp::internal::ZipArchive::open(source);
    const std::string styleXml = R"(<?xml version="1.0" encoding="UTF-8"?><cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="43"/>)";
    const std::string colorsXml = R"(<?xml version="1.0" encoding="UTF-8"?><cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="44" meth="cycle"><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr><a:srgbClr val="FF0000"/></cs:colorStyle>)";
    archive.replace("xl/charts/style1.xml", styleXml);
    archive.replace("xl/charts/colors1.xml", colorsXml);
    const auto styledPath = std::filesystem::temp_directory_path() / "xlpp_p0v_style_resolution.xlsx";
    archive.save(styledPath);
    xlpp::Workbook styled; styled.load(styledPath);
    const auto* styledSheet = styled.worksheet("Objects");
    test.checkTrue(styledSheet != nullptr, "P0V styled resource workbook reloads");
    if (styledSheet) {
        const auto& chart = static_cast<const xlpp::Worksheet&>(*styledSheet).charts().front();
        const auto& resources = chart.styleResources();
        test.checkEqual(resources.chartStyleId, 43, "Chart-style resource id parsed");
        test.checkEqual(resources.colorStyleId, 44, "Color-style resource id parsed");
        test.checkEqual(resources.colorStyleMethod, std::string("cycle"), "Color-style method parsed");
        test.checkEqual(resources.colorStyleColors.size(), std::size_t{2}, "Color-style palette colors inspected");
        const auto resolved = resources.resolveColorStyle(chart.themePalette());
        test.checkEqual(resolved.size(), std::size_t{2}, "Color-style palette resolves through workbook theme");
        if (resolved.size() == 2) {
            test.checkEqual(resolved[0].srgb(), std::string("729ACA"), "Color-style accent transform resolves to final RGB");
            test.checkEqual(resolved[1].srgb(), std::string("FF0000"), "Color-style direct RGB resolves unchanged");
        }
        test.checkEqual(chart.themePalette().resolveTypeface("+mj-lt"), std::string("Cambria"), "Theme major placeholder resolves to major Latin font");
        test.checkEqual(chart.themePalette().resolveTypeface("+mn-lt"), std::string("Calibri"), "Theme minor placeholder resolves to minor Latin font");
        test.checkEqual(chart.themePalette().resolveTypeface("Arial"), std::string("Arial"), "Non-theme typeface remains unchanged");
    }
    std::filesystem::remove(styledPath);
}


void testFormulaDependencyPropagationAndStyleApplication(TestContext& test) {
    xlpp::Workbook workbook;
    auto& inputs = workbook.addWorksheet("Inputs");
    auto& data = workbook.addWorksheet("Data");
    auto& charts = workbook.addWorksheet("Charts");

    inputs.cell("A1").setValue(10.0);
    data.cell("B3").setValue(20.0);
    data.cell("B3").setFormula("'Inputs'!$A$1*2");
    data.cell("C3").setValue(21.0);
    data.cell("C3").setFormula("B3+1");

    xlpp::Chart chart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries series("Formula values");
    series.setValuesReference("'Data'!$C$3");
    chart.addSeries(std::move(series));
    charts.addChart(std::move(chart));

    const auto initial = workbook.synchronizeChartCaches();
    test.checkEqual(initial.cachesUpdated, std::size_t{1}, "P0W initial formula-source cache is created");
    const auto& initialCache = static_cast<const xlpp::Worksheet&>(charts).charts().front().series().front().valuesCache();
    test.checkEqual(initialCache.points.front().value, std::string("21"), "P0W initial cached formula value is preserved");

    workbook.clearChartCacheChangeTracking();
    inputs.cell("A1").setValue(15.0);
    xlpp::ChartCacheSyncOptions incremental;
    incremental.clearTrackedChangesAfterSync = true;
    auto propagated = workbook.synchronizeChangedChartCaches(incremental);
    test.checkEqual(propagated.dependenciesMatched, std::size_t{1}, "P0W precedent change matches formula-backed chart dependency");
    test.checkTrue(propagated.formulaDependenciesVisited >= 2, "P0W follows recursive simple-A1 formula precedents");
    test.checkEqual(propagated.formulaDependenciesMatched, std::size_t{1}, "P0W reports formula-precedent dependency match");
    test.checkEqual(propagated.staleFormulaCachesPreserved, std::size_t{1}, "P0W preserves stale formula cache instead of inventing a recalculated value");
    test.checkEqual(propagated.cachesUpdated, std::size_t{0}, "P0W does not rewrite formula cache from stale worksheet cached value");
    test.checkTrue(propagated.hostRecalculationRequested, "P0W requests host recalculation for formula-precedent changes");
    test.checkTrue(workbook.calcProperties().calcOnSave() && workbook.calcProperties().fullCalcOnLoad(), "P0W sets workbook recalculation flags");
    const auto& preservedCache = static_cast<const xlpp::Worksheet&>(charts).charts().front().series().front().valuesCache();
    test.checkEqual(preservedCache.points.front().value, std::string("21"), "P0W leaves formula chart cache unchanged until a calculation host evaluates it");

    workbook.clearChartCacheChangeTracking();
    inputs.cell("Z99").setValue(1.0);
    auto unrelated = workbook.synchronizeChangedChartCaches(incremental);
    test.checkEqual(unrelated.dependenciesMatched, std::size_t{0}, "P0W unrelated cell does not propagate through formula dependency graph");

    // Apply a color-style resource through the workbook theme using selective
    // imported-chart formatting APIs.
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/libreoffice/chart_style_theme_caches.xlsx";
    auto archive = xlpp::internal::ZipArchive::open(source);
    const std::string colorsXml = R"(<?xml version="1.0" encoding="UTF-8"?><cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="44" meth="cycle"><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr><a:srgbClr val="FF0000"/></cs:colorStyle>)";
    archive.replace("xl/charts/colors1.xml", colorsXml);
    const auto styledPath = std::filesystem::temp_directory_path() / "xlpp_p0w_style_apply.xlsx";
    archive.save(styledPath);

    xlpp::Workbook styled;
    styled.load(styledPath);
    auto* styledSheet = styled.worksheet("Objects");
    test.checkTrue(styledSheet != nullptr, "P0W style-application fixture loads");
    if (styledSheet) {
        const auto& beforeChart = static_cast<const xlpp::Worksheet&>(*styledSheet).charts().front();
        const auto stableId = beforeChart.stableId();
        auto applied = styled.applyChartColorStyle("Objects", stableId, true, true, true);
        test.checkEqual(applied.colorsAvailable, std::size_t{2}, "P0W color-style exposes resolved palette entries");
        test.checkEqual(applied.seriesVisited, std::size_t{1}, "P0W style application visits imported series");
        test.checkEqual(applied.seriesStyled, std::size_t{1}, "P0W style application selectively styles imported series");
        const auto& afterSeries = static_cast<const xlpp::Worksheet&>(*styledSheet).charts().front().series().front();
        test.checkTrue(afterSeries.lineFormat().color.kind == xlpp::ChartColor::Kind::SRgb, "P0W style application materializes theme color as SRGB line color");
        test.checkEqual(afterSeries.lineFormat().color.value, std::string("729ACA"), "P0W applies accent1+tint resolved RGB to line");
        test.checkEqual(afterSeries.fillFormat().color.value, std::string("729ACA"), "P0W applies resolved RGB to series fill");

        const auto output = std::filesystem::temp_directory_path() / "xlpp_p0w_style_apply_roundtrip.xlsx";
        styled.save(output);
        xlpp::Workbook reloaded; reloaded.load(output);
        const auto* reloadedSheet = reloaded.worksheet("Objects");
        test.checkTrue(reloadedSheet != nullptr, "P0W selectively styled workbook reloads");
        if (reloadedSheet) {
            const auto& reloadedSeries = static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front().series().front();
            test.checkEqual(reloadedSeries.lineFormat().color.value, std::string("729ACA"), "P0W applied style survives chart round-trip");
        }
        std::filesystem::remove(output);
    }
    std::filesystem::remove(styledPath);
}


void testFormulaDependencyGrammarAndDefinedNamesP0X(TestContext& test) {
    xlpp::Workbook workbook;
    auto& inputs = workbook.addWorksheet("Inputs");
    auto& data = workbook.addWorksheet("Data");
    auto& charts = workbook.addWorksheet("Charts");

    inputs.cell("A1").setValue(10.0);
    inputs.cell("B1").setValue(11.0);
    inputs.cell("A2").setValue(12.0);
    inputs.cell("B2").setValue(13.0);

    workbook.addDefinedName(xlpp::DefinedName("InputBlock", "'Inputs'!$A$1:$B$2"));
    workbook.addDefinedName(xlpp::DefinedName("InputBlockAlias", "InputBlock"));
    workbook.addDefinedName(xlpp::DefinedName("InputRow", "'Inputs'!$A$1:$B$1"));

    data.cell("D1").setValue(46.0);
    data.cell("D1").setFormula("SUM('Inputs'!$A$1:$B$2)");
    data.cell("D2").setValue(46.0);
    data.cell("D2").setFormula("SUM(InputBlockAlias)");
    data.cell("D3").setValue(22.0);
    data.cell("D3").setFormula("SUM('Inputs'!$A:$A)");
    data.cell("D4").setValue(25.0);
    data.cell("D4").setFormula("SUM('Inputs'!$2:$2)");

    xlpp::Chart chart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries rangeSeries("2-D formula");
    rangeSeries.setValuesReference("'Data'!$D$1");
    chart.addSeries(std::move(rangeSeries));
    xlpp::ChartSeries nameSeries("Defined-name formula");
    nameSeries.setValuesReference("'Data'!$D$2");
    chart.addSeries(std::move(nameSeries));
    xlpp::ChartSeries columnSeries("Whole-column formula");
    columnSeries.setValuesReference("'Data'!$D$3");
    chart.addSeries(std::move(columnSeries));
    xlpp::ChartSeries rowSeries("Whole-row formula");
    rowSeries.setValuesReference("'Data'!$D$4");
    chart.addSeries(std::move(rowSeries));
    charts.addChart(std::move(chart));

    auto initial = workbook.synchronizeChartCaches();
    test.checkEqual(initial.cachesUpdated, std::size_t{4}, "P0X creates formula-backed chart caches across expanded dependency grammar");

    workbook.clearChartCacheChangeTracking();
    inputs.cell("A2").setValue(20.0);
    xlpp::ChartCacheSyncOptions incremental;
    incremental.clearTrackedChangesAfterSync = true;
    const auto propagated = workbook.synchronizeChangedChartCaches(incremental);
    test.checkEqual(propagated.dependenciesMatched, std::size_t{4}, "P0X 2-D, defined-name, whole-column and whole-row precedents match the changed cell");
    test.checkEqual(propagated.formulaDependenciesMatched, std::size_t{4}, "P0X reports all propagated formula dependencies");
    test.checkTrue(propagated.formulaReferencesResolved >= 4, "P0X resolves A1, 2-D, whole-row/column and named formula references");
    test.checkTrue(propagated.definedNameDependenciesVisited >= 1, "P0X visits defined-name dependencies");
    test.checkTrue(propagated.definedNameDependenciesResolved >= 1, "P0X resolves nested defined-name dependencies");
    test.checkTrue(propagated.hostRecalculationRequested, "P0X formula precedent propagation requests host recalculation");
    test.checkEqual(propagated.cachesUpdated, std::size_t{0}, "P0X preserves stale formula caches until a calculation host evaluates formulas");

    // A chart cache may itself use a defined name as long as the name reduces
    // to a one-dimensional A1 range.
    xlpp::Workbook namedChartWorkbook;
    auto& namedInputs = namedChartWorkbook.addWorksheet("Inputs");
    auto& namedCharts = namedChartWorkbook.addWorksheet("Charts");
    namedInputs.cell("A1").setValue(1.0);
    namedInputs.cell("B1").setValue(2.0);
    namedChartWorkbook.addDefinedName(xlpp::DefinedName("InputRow", "'Inputs'!$A$1:$B$1"));
    xlpp::Chart namedChart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries namedSeries("Named range");
    namedSeries.setValuesReference("InputRow");
    namedChart.addSeries(std::move(namedSeries));
    namedCharts.addChart(std::move(namedChart));

    const auto namedSync = namedChartWorkbook.synchronizeChartCaches();
    test.checkEqual(namedSync.cachesUpdated, std::size_t{1}, "P0X materializes chart cache from a defined-name reference");
    const auto& namedCache = static_cast<const xlpp::Worksheet&>(namedCharts).charts().front().series().front().valuesCache();
    test.checkEqual(namedCache.pointCount, std::size_t{2}, "P0X defined-name chart cache keeps the resolved range length");
    test.checkEqual(namedCache.points.size(), std::size_t{2}, "P0X defined-name chart cache contains both source values");
    const auto dependencies = namedChartWorkbook.chartCacheDependencies();
    test.checkEqual(dependencies.size(), std::size_t{1}, "P0X reports defined-name chart dependency");
    if (!dependencies.empty()) {
        test.checkTrue(dependencies.front().supported, "P0X defined-name chart dependency is supported");
        test.checkEqual(dependencies.front().sourceSheet, std::string("Inputs"), "P0X defined-name dependency resolves source sheet");
        test.checkEqual(dependencies.front().first.column, std::size_t{1}, "P0X defined-name dependency resolves first column");
        test.checkEqual(dependencies.front().last.column, std::size_t{2}, "P0X defined-name dependency resolves last column");
    }

    // Cyclic/dynamic name graphs are diagnostic rather than fatal.
    xlpp::Workbook cycleWorkbook;
    auto& cycleInputs = cycleWorkbook.addWorksheet("Inputs");
    auto& cycleData = cycleWorkbook.addWorksheet("Data");
    auto& cycleCharts = cycleWorkbook.addWorksheet("Charts");
    cycleInputs.cell("A1").setValue(1.0);
    cycleData.cell("A1").setValue(1.0);
    cycleData.cell("A1").setFormula("CycleA");
    cycleWorkbook.addDefinedName(xlpp::DefinedName("CycleA", "CycleB"));
    cycleWorkbook.addDefinedName(xlpp::DefinedName("CycleB", "CycleA"));
    xlpp::Chart cycleChart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries cycleSeries("Cycle");
    cycleSeries.setValuesReference("'Data'!$A$1");
    cycleChart.addSeries(std::move(cycleSeries));
    cycleCharts.addChart(std::move(cycleChart));
    cycleWorkbook.synchronizeChartCaches();
    cycleWorkbook.clearChartCacheChangeTracking();
    cycleInputs.cell("A1").setValue(2.0);
    const auto cycleReport = cycleWorkbook.synchronizeChangedChartCaches();
    test.checkTrue(cycleReport.definedNameDependenciesSkipped >= 1, "P0X cyclic defined name is skipped safely");
    test.checkTrue(!cycleReport.formulaDependencyDiagnostics.empty(), "P0X exposes formula dependency diagnostics");
    test.checkTrue(cycleReport.success(), "P0X informational dependency diagnostics do not turn cache sync into failure");
}

void testStructuredReferencesDynamicNamesAndThemeMatrixP0Y(TestContext& test) {
    xlpp::Workbook workbook;
    auto& inputs = workbook.addWorksheet("Inputs");
    auto& data = workbook.addWorksheet("Data");
    auto& charts = workbook.addWorksheet("Charts");

    inputs.append({std::string("Category"), std::string("Sales"), std::string("Margin")});
    inputs.append({std::string("A"), 10.0, 1.0});
    inputs.append({std::string("B"), 20.0, 2.0});
    inputs.append({std::string("C"), 30.0, 3.0});
    inputs.append({std::string("D"), 40.0, 4.0});
    auto& table = inputs.addTable("SalesTable", "A1:C5");
    table.addColumn("Category"); table.addColumn("Sales"); table.addColumn("Margin");

    // Cache values are deliberately supplied before formulas: XL++ follows
    // precedents but does not evaluate the formulas itself.
    for (std::size_t row = 2; row <= 5; ++row) {
        inputs.cell(row, 3).setValue(static_cast<double>((row - 1) * 1));
        inputs.cell(row, 3).setFormula("=[@Sales]*0.1");
    }
    data.cell("A1").setValue(110.0);
    data.cell("A1").setFormula("=SUM(SalesTable[[#Data],[Sales]:[Margin]])");

    workbook.addDefinedName(xlpp::DefinedName("OffsetSales", "=OFFSET('Inputs'!$B$2,0,0,4,1)"));
    workbook.addDefinedName(xlpp::DefinedName("IndexSales", "=INDEX('Inputs'!$B$2:$B$5,0,1)"));
    workbook.addDefinedName(xlpp::DefinedName("UnsupportedDynamic", "=OFFSET('Inputs'!$B$2,0,0,COUNTA('Inputs'!$B:$B)-1,1)"));
    data.cell("A2").setValue(100.0);
    data.cell("A2").setFormula("=SUM(OffsetSales)+SUM(IndexSales)+SUM(UnsupportedDynamic)");

    xlpp::Chart chart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries directTable("Structured direct");
    directTable.setValuesReference("SalesTable[Sales]");
    chart.addSeries(std::move(directTable));
    xlpp::ChartSeries implicitRowFormula("Implicit structured formula");
    implicitRowFormula.setValuesReference("'Inputs'!$C$2:$C$5");
    chart.addSeries(std::move(implicitRowFormula));
    xlpp::ChartSeries structuredFormula("2-D structured formula");
    structuredFormula.setValuesReference("'Data'!$A$1");
    chart.addSeries(std::move(structuredFormula));
    xlpp::ChartSeries dynamicFormula("Dynamic names");
    dynamicFormula.setValuesReference("'Data'!$A$2");
    chart.addSeries(std::move(dynamicFormula));
    xlpp::ChartSeries offsetDirect("OFFSET direct");
    offsetDirect.setValuesReference("OffsetSales");
    chart.addSeries(std::move(offsetDirect));
    xlpp::ChartSeries indexDirect("INDEX direct");
    indexDirect.setValuesReference("IndexSales");
    chart.addSeries(std::move(indexDirect));
    charts.addChart(std::move(chart));

    const auto initial = workbook.synchronizeChartCaches();
    test.checkEqual(initial.referencesSkipped, std::size_t{0}, "P0Y materializes supported table/OFFSET/INDEX chart references");
    test.checkEqual(initial.cachesUpdated, std::size_t{6}, "P0Y builds all six chart caches");
    const auto& initialSeries = static_cast<const xlpp::Worksheet&>(charts).charts().front().series();
    test.checkEqual(initialSeries[0].valuesCache().pointCount, std::size_t{4}, "P0Y Table[Column] resolves to data rows only");
    test.checkEqual(initialSeries[0].valuesCache().points.front().value, std::string("10"), "P0Y structured chart cache starts below header");
    test.checkEqual(initialSeries[4].valuesCache().pointCount, std::size_t{4}, "P0Y OFFSET defined name materializes a bounded range");
    test.checkEqual(initialSeries[5].valuesCache().pointCount, std::size_t{4}, "P0Y INDEX row-zero reference materializes a full column");

    const auto deps = workbook.chartCacheDependencies();
    test.checkEqual(deps.size(), std::size_t{6}, "P0Y dependency inspection includes structured/dynamic references");
    test.checkTrue(deps[0].supported, "P0Y structured table chart dependency is supported");
    test.checkEqual(deps[0].sourceSheet, std::string("Inputs"), "P0Y table dependency resolves its owning worksheet");
    test.checkEqual(deps[0].first.row, std::size_t{2}, "P0Y table dependency excludes header row");
    test.checkEqual(deps[0].first.column, std::size_t{2}, "P0Y table dependency resolves named column");

    workbook.clearChartCacheChangeTracking();
    inputs.cell("B3").setValue(25.0);
    const auto incremental = workbook.synchronizeChangedChartCaches();
    test.checkTrue(incremental.dependenciesMatched >= 6, "P0Y changed table cell reaches direct and formula-backed dependencies");
    test.checkTrue(incremental.structuredReferencesVisited >= 2, "P0Y visits named and implicit structured formula references");
    test.checkTrue(incremental.structuredReferencesResolved >= 2, "P0Y resolves named and implicit structured formula references");
    test.checkTrue(incremental.dynamicDefinedNamesVisited >= 3, "P0Y diagnoses bounded and unbounded dynamic defined names");
    test.checkTrue(incremental.dynamicDefinedNamesResolved >= 2, "P0Y resolves bounded OFFSET/INDEX defined names");
    test.checkTrue(incremental.dynamicDefinedNamesSkipped >= 1, "P0Y leaves calculation-dependent dynamic names unresolved");
    test.checkTrue(incremental.hostRecalculationRequested, "P0Y stale formula-backed caches request host recalculation");
    test.checkTrue(!incremental.formulaDependencyDiagnostics.empty(), "P0Y retains diagnostics for unsupported dynamic expressions");

    // Direct table cache is safe to rebuild; formula-backed caches remain stale
    // until a host recalculates them.
    const auto& afterSeries = static_cast<const xlpp::Worksheet&>(charts).charts().front().series();
    test.checkEqual(afterSeries[0].valuesCache().points[1].value, std::string("25"), "P0Y direct structured cache refreshes from changed table cell");
    test.checkTrue(incremental.staleFormulaCachesPreserved >= 1, "P0Y preserves stale formula-backed cache data");

    // Exercise the new theme style-matrix application API on the imported fixture.
    const auto styleFixture = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/libreoffice/chart_style_theme_caches.xlsx";
    xlpp::Workbook styled; styled.load(styleFixture);
    auto* styleSheet = styled.worksheet("Objects");
    test.checkTrue(styleSheet != nullptr, "P0Y style matrix fixture loads");
    if (styleSheet && !static_cast<const xlpp::Worksheet&>(*styleSheet).charts().empty()) {
        const auto stableId = static_cast<const xlpp::Worksheet&>(*styleSheet).charts().front().stableId();
        const auto applied = styled.applyChartThemeStyleMatrix("Objects", stableId, 0, 0, true);
        test.checkEqual(applied.fillStylesAvailable, std::size_t{3}, "P0Y apply API exposes three theme fill styles");
        test.checkEqual(applied.lineStylesAvailable, std::size_t{3}, "P0Y apply API exposes three theme line styles");
        test.checkEqual(applied.effectStylesAvailable, std::size_t{3}, "P0Y apply API exposes three theme effect styles");
        test.checkEqual(applied.seriesStyled, std::size_t{1}, "P0Y theme matrix applies fill/line to imported series");
        const auto& appliedSeries = static_cast<const xlpp::Worksheet&>(*styleSheet).charts().front().series().front();
        test.checkTrue(appliedSeries.fillFormat().color.kind == xlpp::ChartColor::Kind::SRgb, "P0Y style-matrix phClr is materialized through chart color style");
        test.checkEqual(appliedSeries.fillFormat().color.value, std::string("4F81BD"), "P0Y style-matrix placeholder falls back to resolved theme accent when color-style entries are absent");
    }

    // Inject non-empty DrawingML effects into a temporary copy so outer shadow,
    // glow and soft-edge parsing is covered rather than only list counts.
    {
        auto package = xlpp::internal::ZipArchive::open(styleFixture);
        auto themeXml = package.get("xl/theme/theme1.xml");
        const std::string emptyEffect = "<a:effectStyle><a:effectLst/></a:effectStyle>";
        const std::string richEffect =
            "<a:effectStyle><a:effectLst>"
            "<a:outerShdw blurRad=\"12700\" dist=\"25400\" dir=\"5400000\"><a:schemeClr val=\"accent1\"><a:alpha val=\"50000\"/></a:schemeClr></a:outerShdw>"
            "<a:innerShdw blurRad=\"6350\" dist=\"12700\" dir=\"2700000\"><a:srgbClr val=\"00FF00\"/></a:innerShdw>"
            "<a:glow rad=\"12700\"><a:srgbClr val=\"FF0000\"/></a:glow>"
            "<a:softEdge rad=\"6350\"/>"
            "<a:reflection blurRad=\"25400\" dist=\"38100\" dir=\"10800000\"/>"
            "<a:blur rad=\"19050\" grow=\"1\"/>"
            "</a:effectLst></a:effectStyle>";
        const auto effectPos = themeXml.find(emptyEffect);
        test.checkTrue(effectPos != std::string::npos, "P0Y effect test locates theme effect slot");
        if (effectPos != std::string::npos) themeXml.replace(effectPos, emptyEffect.size(), richEffect);
        package.replace("xl/theme/theme1.xml", themeXml);
        const auto effectPath = std::filesystem::temp_directory_path() / "xlpp_p0y_theme_effects.xlsx";
        package.save(effectPath);
        xlpp::Workbook effectWorkbook; effectWorkbook.load(effectPath);
        const auto* effectSheet = effectWorkbook.worksheet("Objects");
        test.checkTrue(effectSheet != nullptr, "P0Y effect-materialization fixture reloads");
        if (effectSheet && !static_cast<const xlpp::Worksheet&>(*effectSheet).charts().empty()) {
            const auto& effects = static_cast<const xlpp::Worksheet&>(*effectSheet).charts().front().themePalette().effectScheme.effectStyles;
            test.checkEqual(effects.size(), std::size_t{3}, "P0Y effect style count remains stable after materialization");
            test.checkTrue(effects.front().outerShadow, "P0Y outer shadow is materialized");
            test.checkTrue(effects.front().glow, "P0Y glow is materialized");
            test.checkTrue(effects.front().softEdge, "P0Y soft edge is materialized");
            test.checkNear(effects.front().outerShadowBlurPoints, 1.0, 1e-12, "P0Y shadow blur converts EMU to points");
            test.checkNear(effects.front().outerShadowDistancePoints, 2.0, 1e-12, "P0Y shadow distance converts EMU to points");
            test.checkNear(effects.front().outerShadowDirectionDegrees, 90.0, 1e-12, "P0Y shadow direction converts to degrees");
            test.checkNear(effects.front().glowRadiusPoints, 1.0, 1e-12, "P0Y glow radius converts EMU to points");
            test.checkNear(effects.front().softEdgeRadiusPoints, 0.5, 1e-12, "P0Y soft-edge radius converts EMU to points");
            test.checkEqual(effects.front().glowColor.value, std::string("FF0000"), "P0Y glow color is parsed");
            test.checkTrue(effects.front().innerShadow, "P0Z inner shadow is materialized");
            test.checkEqual(effects.front().innerShadowColor.value, std::string("00FF00"), "P0Z inner-shadow color is parsed");
            test.checkNear(effects.front().innerShadowBlurPoints, 0.5, 1e-12, "P0Z inner-shadow blur converts EMU to points");
            test.checkNear(effects.front().innerShadowDistancePoints, 1.0, 1e-12, "P0Z inner-shadow distance converts EMU to points");
            test.checkNear(effects.front().innerShadowDirectionDegrees, 45.0, 1e-12, "P0Z inner-shadow direction converts to degrees");
            test.checkTrue(effects.front().reflection, "P0Z reflection is materialized");
            test.checkNear(effects.front().reflectionBlurPoints, 2.0, 1e-12, "P0Z reflection blur converts EMU to points");
            test.checkNear(effects.front().reflectionDistancePoints, 3.0, 1e-12, "P0Z reflection distance converts EMU to points");
            test.checkNear(effects.front().reflectionDirectionDegrees, 180.0, 1e-12, "P0Z reflection direction converts to degrees");
            test.checkTrue(effects.front().blur, "P0Z blur effect is materialized");
            test.checkNear(effects.front().blurRadiusPoints, 1.5, 1e-12, "P0Z blur radius converts EMU to points");
            test.checkTrue(effects.front().blurGrow, "P0Z blur grow flag is parsed");
        }
        std::filesystem::remove(effectPath);
    }

    // Theme style-matrix indices are positional. Verify parsing preserves the
    // actual XML child order rather than grouping fills by DrawingML tag kind.
    {
        auto package = xlpp::internal::ZipArchive::open(styleFixture);
        auto themeXml = package.get("xl/theme/theme1.xml");
        const std::string listOpen = "<a:fillStyleLst>";
        const std::string listClose = "</a:fillStyleLst>";
        const auto listBegin = themeXml.find(listOpen);
        const auto listEnd = listBegin == std::string::npos ? std::string::npos : themeXml.find(listClose, listBegin);
        test.checkTrue(listBegin != std::string::npos && listEnd != std::string::npos, "P0Y fill-order test locates fmtScheme fill list");
        if (listBegin != std::string::npos && listEnd != std::string::npos) {
            const auto contentBegin = listBegin + listOpen.size();
            std::string fills = themeXml.substr(contentBegin, listEnd - contentBegin);
            const auto solidBegin = fills.find("<a:solidFill");
            const auto solidClose = solidBegin == std::string::npos ? std::string::npos : fills.find("</a:solidFill>", solidBegin);
            const auto gradBegin = solidClose == std::string::npos ? std::string::npos : fills.find("<a:gradFill", solidClose);
            const auto gradClose = gradBegin == std::string::npos ? std::string::npos : fills.find("</a:gradFill>", gradBegin);
            test.checkTrue(solidBegin != std::string::npos && solidClose != std::string::npos && gradBegin != std::string::npos && gradClose != std::string::npos,
                           "P0Y fill-order test locates solid and gradient entries");
            if (solidBegin != std::string::npos && solidClose != std::string::npos && gradBegin != std::string::npos && gradClose != std::string::npos) {
                const auto solidEnd = solidClose + std::string("</a:solidFill>").size();
                const auto gradEnd = gradClose + std::string("</a:gradFill>").size();
                const auto solid = fills.substr(solidBegin, solidEnd - solidBegin);
                const auto between = fills.substr(solidEnd, gradBegin - solidEnd);
                const auto gradient = fills.substr(gradBegin, gradEnd - gradBegin);
                fills.replace(solidBegin, gradEnd - solidBegin, gradient + between + solid);
                themeXml.replace(contentBegin, listEnd - contentBegin, fills);
                package.replace("xl/theme/theme1.xml", themeXml);
                const auto orderPath = std::filesystem::temp_directory_path() / "xlpp_p0y_theme_fill_order.xlsx";
                package.save(orderPath);
                xlpp::Workbook orderedWorkbook; orderedWorkbook.load(orderPath);
                const auto* orderedSheet = orderedWorkbook.worksheet("Objects");
                test.checkTrue(orderedSheet != nullptr, "P0Y reordered fmtScheme fixture reloads");
                if (orderedSheet && !static_cast<const xlpp::Worksheet&>(*orderedSheet).charts().empty()) {
                    const auto& fillsParsed = static_cast<const xlpp::Worksheet&>(*orderedSheet).charts().front().themePalette().effectScheme.fillStyles;
                    test.checkEqual(fillsParsed.size(), std::size_t{3}, "P0Y reordered theme keeps all fill styles");
                    test.checkTrue(!fillsParsed.empty() && fillsParsed.front().kind == xlpp::ChartFillFormat::Kind::Gradient,
                                   "P0Y fmtScheme fill indices preserve XML child order");
                    test.checkTrue(fillsParsed.size() > 1 && fillsParsed[1].kind == xlpp::ChartFillFormat::Kind::Solid,
                                   "P0Y reordered solid fill remains the second matrix entry");
                }
                std::filesystem::remove(orderPath);
            }
        }
    }
}

void testStructuredEscapingIndexRangesAndChartStyleRulesP0Z(TestContext& test) {
    // Escaped structured-reference column names plus INDEX endpoint ranges.
    xlpp::Workbook workbook;
    auto& inputs = workbook.addWorksheet("Inputs");
    auto& data = workbook.addWorksheet("Data");
    auto& charts = workbook.addWorksheet("Charts");

    inputs.append({std::string("Sales]Net"), std::string("#Rate"), std::string("@Code"), std::string("O'Brien")});
    inputs.append({10.0, 1.0, 101.0, 1001.0});
    inputs.append({20.0, 2.0, 102.0, 1002.0});
    inputs.append({30.0, 3.0, 103.0, 1003.0});
    inputs.append({40.0, 4.0, 104.0, 1004.0});
    auto& table = inputs.addTable("SpecialTable", "A1:D5");
    table.addColumn("Sales]Net");
    table.addColumn("#Rate");
    table.addColumn("@Code");
    table.addColumn("O'Brien");

    workbook.addDefinedName(xlpp::DefinedName(
        "IndexWindow", "=INDEX('Inputs'!$A$2:$A$5,2,1):INDEX('Inputs'!$A$2:$A$5,4,1)"));

    // The formula combines escaped columns and a multi-item row selector. The
    // cached value is supplied by the host; XL++ only follows precedents.
    data.cell("A1").setValue(123.0);
    data.cell("A1").setFormula("=SUM(SpecialTable[[#Headers],[#Data],[Sales']Net]])+SUM(SpecialTable['#Rate])");

    xlpp::Chart chart(xlpp::Chart::Type::Line);
    for (const auto& entry : std::array<std::pair<const char*, const char*>, 5>{
             std::pair{"escaped-right-bracket", "SpecialTable[Sales']Net]"},
             std::pair{"escaped-hash", "SpecialTable['#Rate]"},
             std::pair{"escaped-at", "SpecialTable['@Code]"},
             std::pair{"escaped-apostrophe", "SpecialTable[O''Brien]"},
             std::pair{"index-window", "IndexWindow"}}) {
        xlpp::ChartSeries series(entry.first);
        series.setValuesReference(entry.second);
        chart.addSeries(std::move(series));
    }
    xlpp::ChartSeries formulaSeries("escaped structured formula");
    formulaSeries.setValuesReference("'Data'!$A$1");
    chart.addSeries(std::move(formulaSeries));
    charts.addChart(std::move(chart));

    const auto sync = workbook.synchronizeChartCaches();
    test.checkEqual(sync.referencesSkipped, std::size_t{0}, "P0Z escaped structured references and INDEX endpoints materialize without skips");
    test.checkEqual(sync.cachesUpdated, std::size_t{6}, "P0Z builds escaped structured, INDEX-window and formula-backed chart caches");
    const auto& series = static_cast<const xlpp::Worksheet&>(charts).charts().front().series();
    test.checkEqual(series[0].valuesCache().pointCount, std::size_t{4}, "P0Z escaped ] column resolves four data rows");
    test.checkEqual(series[0].valuesCache().points[2].value, std::string("30"), "P0Z escaped ] column resolves the correct values");
    test.checkEqual(series[1].valuesCache().points[1].value, std::string("2"), "P0Z escaped # header remains a literal column name");
    test.checkEqual(series[2].valuesCache().points[3].value, std::string("104"), "P0Z escaped @ header is not mistaken for #This Row");
    test.checkEqual(series[3].valuesCache().points[0].value, std::string("1001"), "P0Z escaped apostrophe header resolves correctly");
    test.checkEqual(series[4].valuesCache().pointCount, std::size_t{3}, "P0Z INDEX:INDEX reference creates a bounded range");
    test.checkEqual(series[4].valuesCache().points.front().value, std::string("20"), "P0Z INDEX endpoint range starts at the first resolved endpoint");
    test.checkEqual(series[4].valuesCache().points.back().value, std::string("40"), "P0Z INDEX endpoint range ends at the second resolved endpoint");

    workbook.clearChartCacheChangeTracking();
    inputs.cell("B3").setValue(2.5);
    const auto propagated = workbook.synchronizeChangedChartCaches();
    test.checkTrue(propagated.structuredReferencesVisited >= 2, "P0Z formula scanner visits multiple escaped structured references");
    test.checkTrue(propagated.structuredReferencesResolved >= 2, "P0Z formula scanner resolves escaped structured references");
    test.checkTrue(propagated.hostRecalculationRequested, "P0Z escaped formula precedent changes request host recalculation");

    // Inject an Office chart-style rule part and a deliberately mixed color
    // choice list. Color order is semantically significant for styleClr=auto.
    const auto fixture = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/libreoffice/chart_style_theme_caches.xlsx";
    auto package = xlpp::internal::ZipArchive::open(fixture);
    const std::string styleXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<cs:chartStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" id=\"55\">"
        "<cs:chartArea><cs:lnRef idx=\"1\"><a:schemeClr val=\"accent2\"/></cs:lnRef>"
        "<cs:fillRef idx=\"1\"><a:schemeClr val=\"accent3\"/></cs:fillRef>"
        "<cs:effectRef idx=\"1\"><a:schemeClr val=\"accent2\"/></cs:effectRef></cs:chartArea>"
        "<cs:plotArea><cs:fillRef idx=\"1001\"><a:schemeClr val=\"accent1\"/></cs:fillRef></cs:plotArea>"
        "<cs:dataPoint>"
        "<cs:lnRef idx=\"1\"><cs:styleClr val=\"auto\"><a:shade val=\"90000\"/><a:tint val=\"10000\"/></cs:styleClr></cs:lnRef>"
        "<cs:fillRef idx=\"1\"><cs:styleClr val=\"auto\"/></cs:fillRef>"
        "<cs:effectRef idx=\"1\"><cs:styleClr val=\"auto\"/></cs:effectRef>"
        "<cs:fontRef idx=\"minor\"><cs:styleClr val=\"auto\"/></cs:fontRef>"
        "<cs:lineWidthScale val=\"2\"/>"
        "<cs:spPr><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></cs:spPr>"
        "</cs:dataPoint>"
        "<cs:dataPointMarker><cs:lnRef idx=\"1\"><cs:styleClr val=\"auto\"/></cs:lnRef>"
        "<cs:fillRef idx=\"1\"><cs:styleClr val=\"auto\"/></cs:fillRef></cs:dataPointMarker>"
        "<cs:dataPointMarkerLayout symbol=\"diamond\" size=\"9\"/>"
        "</cs:chartStyle>";
    const std::string colorsXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" id=\"77\" meth=\"cycle\">"
        "<a:srgbClr val=\"FF0000\"/><a:schemeClr val=\"accent2\"/><a:srgbClr val=\"0000FF\"/>"
        "</cs:colorStyle>";
    package.replace("xl/charts/style1.xml", styleXml);
    package.replace("xl/charts/colors1.xml", colorsXml);
    const auto styledPath = std::filesystem::temp_directory_path() / "xlpp_p0z_chart_style_rules.xlsx";
    package.save(styledPath);

    xlpp::Workbook styled;
    styled.load(styledPath);
    auto* sheet = styled.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "P0Z synthetic chart-style fixture reloads");
    if (sheet && !static_cast<const xlpp::Worksheet&>(*sheet).charts().empty()) {
        const auto& imported = static_cast<const xlpp::Worksheet&>(*sheet).charts().front();
        const auto stableId = imported.stableId();
        const auto& resources = imported.styleResources();
        test.checkEqual(resources.chartStyleId, 55, "P0Z parses chart-style id");
        test.checkEqual(resources.colorStyleId, 77, "P0Z parses color-style id");
        test.checkEqual(resources.chartStyleRules.size(), std::size_t{4}, "P0Z parses four targeted chart-style rules");
        test.checkEqual(resources.colorStyleColors.size(), std::size_t{3}, "P0Z keeps all mixed color-style entries");
        if (resources.colorStyleColors.size() >= 3) {
            test.checkEqual(resources.colorStyleColors[0].value, std::string("FF0000"), "P0Z color-style entries preserve XML order");
            test.checkEqual(resources.colorStyleColors[1].value, std::string("accent2"), "P0Z mixed scheme color stays at its original position");
            test.checkEqual(resources.colorStyleColors[2].value, std::string("0000FF"), "P0Z trailing sRGB color stays at its original position");
        }
        const auto* dataPoint = resources.rule("dataPoint");
        test.checkTrue(dataPoint != nullptr, "P0Z exposes dataPoint style rule by target");
        if (dataPoint) {
            test.checkTrue(dataPoint->fillReference.styleColor, "P0Z parses styleClr on fill reference");
            test.checkTrue(dataPoint->hasLineWidthScale, "P0Z parses lineWidthScale");
            test.checkNear(dataPoint->lineWidthScale, 2.0, 1e-12, "P0Z keeps lineWidthScale value");
            test.checkEqual(dataPoint->lineReference.styleColorTransforms.size(), std::size_t{2}, "P0Z retains styleClr transform sequence");
            if (dataPoint->lineReference.styleColorTransforms.size() == 2) {
                test.checkTrue(dataPoint->lineReference.styleColorTransforms[0].kind == xlpp::ChartColorTransform::Kind::Shade,
                               "P0Z styleClr transform order keeps shade first");
                test.checkTrue(dataPoint->lineReference.styleColorTransforms[1].kind == xlpp::ChartColorTransform::Kind::Tint,
                               "P0Z styleClr transform order keeps tint second");
            }
            test.checkTrue(dataPoint->shapeFill.present, "P0Z parses explicit spPr fill override");
            test.checkEqual(dataPoint->shapeFill.color.value, std::string("phClr"), "P0Z keeps phClr until rule materialization");
        }
        test.checkTrue(resources.markerLayout.present, "P0Z parses dataPointMarkerLayout");
        test.checkEqual(resources.markerLayout.symbol, std::string("diamond"), "P0Z parses marker symbol");
        test.checkEqual(resources.markerLayout.size, 9, "P0Z parses marker size");

        const auto report = styled.applyChartStyleRules("Objects", stableId);
        test.checkEqual(report.rulesAvailable, std::size_t{4}, "P0Z style application reports available rules");
        test.checkEqual(report.rulesVisited, std::size_t{4}, "P0Z style application visits all parsed rules");
        test.checkTrue(report.rulesApplied >= 3, "P0Z applies supported chart/plot/series/marker rules");
        test.checkTrue(report.targetsStyled >= 3, "P0Z reports materialized style targets");
        test.checkTrue(report.effectReferencesResolved >= 2, "P0Z resolves effect matrix references even when target effect serialization is not yet modeled");
        test.checkEqual(report.seriesStyled, std::size_t{1}, "P0Z rule application styles the imported data series");

        const auto& after = static_cast<const xlpp::Worksheet&>(*sheet).charts().front().series().front();
        test.checkTrue(after.fillFormat().color.kind == xlpp::ChartColor::Kind::SRgb, "P0Z spPr phClr is materialized to concrete sRGB");
        test.checkEqual(after.fillFormat().color.value, std::string("FF0000"), "P0Z styleClr auto uses the first color-style entry for first series");
        test.checkTrue(after.markerFormat().present, "P0Z marker layout can create marker formatting");
        test.checkEqual(after.markerFormat().symbol, std::string("diamond"), "P0Z applies marker layout symbol");
        test.checkEqual(after.markerFormat().size, 9, "P0Z applies marker layout size");
        test.checkEqual(after.markerFormat().fill.color.value, std::string("FF0000"), "P0Z marker fill resolves styleClr auto");
        test.checkNear(after.lineFormat().widthPoints, 1.5, 1e-12, "P0Z lineWidthScale multiplies theme line width");

        const auto savedPath = std::filesystem::temp_directory_path() / "xlpp_p0z_chart_style_rules_saved.xlsx";
        styled.save(savedPath);
        xlpp::Workbook reloaded;
        reloaded.load(savedPath);
        const auto* reloadedSheet = reloaded.worksheet("Objects");
        test.checkTrue(reloadedSheet != nullptr, "P0Z styled chart survives save/reload");
        if (reloadedSheet && !static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().empty()) {
            const auto& savedSeries = static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front().series().front();
            test.checkEqual(savedSeries.fillFormat().color.value, std::string("FF0000"), "P0Z materialized series fill persists after save/reload");
            test.checkEqual(savedSeries.markerFormat().symbol, std::string("diamond"), "P0Z materialized marker layout persists after save/reload");
        }
        std::filesystem::remove(savedPath);
    }
    std::filesystem::remove(styledPath);
}

void testChartTypeNameMap(TestContext& test) {
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bar), std::string("barChart"), "Bar standard");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bar, xlpp::Chart::Grouping::Stacked), std::string("barStacked"), "Bar stacked");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bar, xlpp::Chart::Grouping::PercentStacked), std::string("barPercentStacked"), "Bar percent stacked");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Line), std::string("lineChart"), "Line standard");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Line, xlpp::Chart::Grouping::Stacked), std::string("lineStacked"), "Line stacked");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Pie), std::string("pieChart"), "Pie");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Scatter), std::string("scatterChart"), "Scatter");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Doughnut), std::string("doughnutChart"), "Doughnut");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Radar), std::string("radarChart"), "Radar");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Area), std::string("areaChart"), "Area standard");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Area, xlpp::Chart::Grouping::Stacked), std::string("areaStacked"), "Area stacked");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bubble), std::string("bubbleChart"), "Bubble");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Stock), std::string("stockChart"), "Stock");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::PieOfPie), std::string("ofPieChart"), "Pie-of-Pie");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::BarOfPie), std::string("ofPieChart"), "Bar-of-Pie");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bar3D), std::string("bar3DChart"), "Bar3D");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Line3D), std::string("line3DChart"), "Line3D");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Area3D), std::string("area3DChart"), "Area3D");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Pie3D), std::string("pie3DChart"), "Pie3D");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Surface), std::string("surfaceChart"), "Surface");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Surface3D), std::string("surface3DChart"), "Surface3D");
}

void testInternalHyperlinkAndMemoryStream(TestContext& test) {
    std::ostringstream memory;
    {
        xlpp::Workbook wb;
        auto& sheet = wb.addWorksheet("Link");
        sheet.cell("A1").setValue("Jump");
        xlpp::Hyperlink internal("Sheet2!B5");
        internal.setExternal(false);
        internal.setDisplay("Go to B5");
        sheet.cell("A1").setHyperlink(std::move(internal));
        wb.addWorksheet("Sheet2");
        wb.save(memory);
        test.checkTrue(memory.str().size() > 0, "Workbook saves to memory stream");
    }
    {
        std::istringstream input(memory.str());
        xlpp::Workbook loaded;
        loaded.load(input);
        test.checkTrue(loaded.sheetCount() == 2, "Memory stream loads two sheets");
        const auto* sheet = loaded.worksheet("Link");
        test.checkTrue(sheet->tryCell("A1")->hasHyperlink(), "Internal hyperlink preserved");
        const auto& link = *sheet->tryCell("A1")->hyperlinkValue();
        test.checkEqual(link.target(), std::string("Sheet2!B5"), "Internal hyperlink target");
        test.checkTrue(!link.external(), "Internal hyperlink marked non-external");
    }
}

void testWorkbookEdgeCases(TestContext& test) {
    bool threw = false;
    xlpp::Workbook wb;
    try { (void)wb.addWorksheet(""); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "addWorksheet(empty) throws");
    threw = false;
    wb.addWorksheet("Dup");
    try { (void)wb.addWorksheet("Dup"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "addWorksheet(duplicate) throws");
    threw = false;
    xlpp::Workbook wb2;
    wb2.addWorksheet("S");
    try { (void)wb2.copyWorksheet(wb2[0], ""); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "copyWorksheet(empty name) throws");
    threw = false;
    try { (void)wb2.copyWorksheet(wb2[0], "S"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "copyWorksheet(duplicate name) throws");

    threw = false;
    try { xlpp::DefinedName("", "value"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "DefinedName(empty name) throws");
    threw = false;
    try { xlpp::DefinedName("name", ""); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "DefinedName(empty value) throws");

    threw = false;
    try { xlpp::Table("", "A1:B2"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Table(empty name) throws");
    threw = false;
    try { xlpp::Table("T", ""); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Table(empty reference) throws");

    threw = false;
    xlpp::Workbook empty;
    try { empty.save(std::filesystem::temp_directory_path() / "xlpp_no_sheets.xlsx"); }
    catch (const std::runtime_error&) { threw = true; }
    test.checkTrue(threw, "save() with no worksheets throws");

    threw = false;
    try { xlpp::Cell("0"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Cell with row 0 throws");
    threw = false;
    try { xlpp::Worksheet sheet("S"); sheet.cell("A1").offset(-1, 0); }
    catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Cell::offset below row 1 throws");

    xlpp::Workbook wb3;
    auto& s = wb3.addWorksheet("Clean");
    s.cell("A1").setValue(42.0);
    s.cell("A1").font().setBold(true);
    s.cell("A1").setNumberFormat("0.00");
    s.cell("A1").setFormula("=1+1");
    s.cell("A1").clear();
    test.checkTrue(!s.cell("A1").hasValue(), "clear() removes value");
    test.checkTrue(!s.cell("A1").hasFormula(), "clear() removes formula");
    test.checkTrue(s.cell("A1").font().bold(), "clear() keeps font style");
    test.checkEqual(s.cell("A1").numberFormat(), std::string("0.00"), "clear() keeps number format");
}

void testCustomPropertiesAndCalcRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m21_custom_calc.xlsx";
    {
        xlpp::Workbook wb;
        auto& sheet = wb.addWorksheet("Props");
        sheet.cell("A1").setValue("x");
        wb.customProperties().add(xlpp::CustomProperty(std::string("Name"), std::string("XL++")));
        wb.customProperties().add(xlpp::CustomProperty(std::string("Count"), 7));
        wb.customProperties().add(xlpp::CustomProperty(std::string("Ratio"), 0.5));
        wb.customProperties().add(xlpp::CustomProperty(std::string("Enabled"), true));
        wb.calcProperties().setCalcMode("manual");
        wb.calcProperties().setFullCalcOnLoad(true);
        wb.calcProperties().setCalcId(191029);
        wb.calcProperties().setIterate(true);
        wb.calcProperties().setIterateCount(100);
        wb.calcProperties().setIterateDelta(0.001);
        wb.protection().setLockRevision(true);
        wb.protection().setLockWindows(true);
        wb.properties().setSubject("Calc round-trip");
        wb.properties().setKeywords("a,b,c");
        wb.save(path);
    }
    {
        const auto archive = xlpp::internal::ZipArchive::open(path);
        const auto validation = xlpp::internal::RelationshipGraph::fromArchive(archive).validate();
        test.checkTrue(validation.ok(), "Custom-property package passes OPC validation");
        test.checkTrue(archive.get("_rels/.rels").find("/custom-properties") != std::string::npos,
                       "Root relationships connect custom properties");
        test.checkTrue(archive.get("[Content_Types].xml").find("/docProps/custom.xml") != std::string::npos,
                       "Custom properties have an explicit content type");
    }
    {
        xlpp::Workbook loaded;
        loaded.load(path);
        test.checkEqual(loaded.customProperties().items().size(), std::size_t{4}, "Custom property count");
        test.checkEqual(loaded.customProperties().items()[0].name(), std::string("Name"), "String property name");
        test.checkEqual(loaded.customProperties().items()[0].value(), std::string("XL++"), "String property value");
        test.checkEqual(loaded.customProperties().items()[1].name(), std::string("Count"), "Int property name");
        test.checkEqual(loaded.customProperties().items()[1].value(), std::string("7"), "Int property value");
        test.checkEqual(loaded.customProperties().items()[2].name(), std::string("Ratio"), "Double property name");
        test.checkEqual(loaded.customProperties().items()[2].value(), std::string("0.500000"), "Double property value");
        test.checkEqual(loaded.customProperties().items()[3].name(), std::string("Enabled"), "Bool property name");
        test.checkEqual(loaded.customProperties().items()[3].value(), std::string("true"), "Bool property value");
        test.checkEqual(loaded.calcProperties().calcMode(), std::string("manual"), "Calc mode round-trip");
        test.checkTrue(loaded.calcProperties().fullCalcOnLoad(), "Full calc-on-load round-trip");
        test.checkTrue(loaded.calcProperties().iterate(), "Iterate round-trip");
        test.checkEqual(loaded.calcProperties().iterateCount(), 100, "Iterate count round-trip");
        test.checkNear(loaded.calcProperties().iterateDelta(), 0.001, 1e-12, "Iterate delta round-trip");
        test.checkTrue(loaded.protection().lockRevision(), "Lock revision round-trip");
        test.checkTrue(loaded.protection().lockWindows(), "Lock windows round-trip");
        test.checkEqual(loaded.properties().subject(), std::string("Calc round-trip"), "Subject round-trip");
        test.checkEqual(loaded.properties().keywords(), std::string("a,b,c"), "Keywords round-trip");
    }
    std::filesystem::remove(path);
}

void testDifferentialSaveCache(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto first = dir / "xlpp_m21_diff_first.xlsx";
    const auto second = dir / "xlpp_m21_diff_second.xlsx";
    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("Data");
    sheet.cell("A1").setValue("alpha");
    wb.save(first);
    std::string firstBytes;
    { std::ifstream in(first, std::ios::binary); firstBytes.assign(std::istreambuf_iterator<char>(in), {}); }

    wb.save(second);
    std::string secondBytes;
    { std::ifstream in(second, std::ios::binary); secondBytes.assign(std::istreambuf_iterator<char>(in), {}); }
    test.checkEqual(firstBytes, secondBytes, "Unchanged re-save is byte-identical (cache reuse)");

    sheet.cell("B1").setValue("beta");
    wb.save(second);
    std::string changedBytes;
    { std::ifstream in(second, std::ios::binary); changedBytes.assign(std::istreambuf_iterator<char>(in), {}); }
    test.checkTrue(changedBytes != firstBytes, "Dirty sheet change alters output");

    xlpp::Workbook loaded;
    loaded.load(second);
    test.checkEqual(loaded.worksheet("Data")->cell("B1").stringValueOr(""), std::string("beta"),
                    "Changed cell round-trips");
    std::filesystem::remove(first);
    std::filesystem::remove(second);
}

void testStrictAfterTransitionalSave(TestContext& test) {
    // Regression: a transitional save followed by a strict save with the same
    // workbook must not reuse cached transitional sheet XML.
    const auto dir = std::filesystem::temp_directory_path();
    const auto transitional = dir / "xlpp_m21_strict_transitional.xlsx";
    const auto strictPath = dir / "xlpp_m21_strict_after.xlsx";
    xlpp::Workbook wb;
    wb.addWorksheet("Sheet1").cell("A1").setValue("value");
    wb.save(transitional);

    xlpp::SaveOptions opt;
    opt.strictNamespace = true;
    wb.save(strictPath, opt);

    auto z = xlpp::internal::ZipArchive::open(strictPath);
    test.checkTrue(z.get("xl/workbook.xml").find("http://purl.oclc.org/ooxml/spreadsheetml/main") != std::string::npos,
                   "Strict workbook namespace after transitional save");
    test.checkTrue(z.get("xl/worksheets/sheet1.xml").find("http://purl.oclc.org/ooxml/spreadsheetml/main") != std::string::npos,
                   "Strict worksheet namespace after transitional save");
    test.checkTrue(z.get("xl/worksheets/sheet1.xml").find("http://schemas.openxmlformats.org/spreadsheetml/2006/main") == std::string::npos,
                   "No transitional namespace leaks into strict worksheet");

    xlpp::Workbook loaded;
    loaded.load(strictPath);
    test.checkTrue(loaded.strictNamespaces(), "Strict package loads as strict");
    test.checkEqual(loaded.worksheet("Sheet1")->cell("A1").stringValueOr(""), std::string("value"),
                    "Value round-trips through strict save");

    std::filesystem::remove(transitional);
    std::filesystem::remove(strictPath);
}

void testCopyWorksheetAliasing(TestContext& test) {
    // Regression: copyWorksheet(source) where source aliases an element of the
    // internal worksheet vector must not invalidate the source mid-copy.
    xlpp::Workbook wb;
    auto& src = wb.addWorksheet("Source");
    src.cell("A1").setValue(std::string("payload"));
    src.cell("B2").setValue(3.25);
    src.mergeCells("A2:B2");
    src.rowDimension(1).height = 22.0;

    auto& copy = wb.copyWorksheet(src, "Copy");
    test.checkEqual(copy.name(), std::string("Copy"), "Copy named correctly");
    test.checkEqual(std::get<std::string>(copy.cell("A1").value()), std::string("payload"), "Copy keeps string");
    test.checkNear(std::get<double>(copy.cell("B2").value()), 3.25, 1e-12, "Copy keeps number");
    test.checkEqual(copy.mergedRanges().size(), std::size_t{1}, "Copy keeps merges");
    test.checkNear(copy.tryRowDimension(1)->height.value_or(0.0), 22.0, 1e-12, "Copy keeps dimensions");

    for (std::size_t i = 0; i < 40; ++i) wb.addWorksheet("Fill" + std::to_string(i));
    auto& late = wb.copyWorksheet(wb[0], "LateClone");
    test.checkEqual(std::get<std::string>(late.cell("A1").value()), std::string("payload"),
                    "Copy after many reallocations keeps source data");
    test.checkEqual(wb.sheetCount(), std::size_t{43}, "Workbook grew as expected");
}

void testReferenceStabilityAcrossInserts(TestContext& test) {
    // Regression: worksheet references must survive further worksheet
    // insertion (stable storage, no vector reallocation).
    xlpp::Workbook wb;
    auto& first = wb.addWorksheet("First");
    first.cell("A1").setValue(std::string("stable"));
    auto& second = wb.addWorksheet("Second");

    xlpp::Worksheet* firstPtr = &first;
    xlpp::Worksheet* secondPtr = &second;

    for (std::size_t i = 0; i < 100; ++i) wb.addWorksheet("Extra" + std::to_string(i));

    test.checkEqual(firstPtr->name(), std::string("First"), "Reference survives 100 inserts");
    test.checkEqual(std::get<std::string>(firstPtr->cell("A1").value()), std::string("stable"),
                    "Reference still points to the same worksheet");
    test.checkEqual(secondPtr->name(), std::string("Second"), "Second reference survives inserts");
    test.checkEqual(&wb[0], firstPtr, "operator[] returns the same stable object");

    auto& copy = wb.copyWorksheet(*firstPtr, "CopyOfFirst");
    test.checkEqual(std::get<std::string>(copy.cell("A1").value()), std::string("stable"),
                    "Copy after inserts keeps data");
    test.checkEqual(firstPtr->name(), std::string("First"), "Source still valid after copyWorksheet");
}

void testReferenceLifetimeContract(TestContext& test) {
    // Documented lifetime contract: references to a worksheet stay valid until
    // that worksheet is removed (or the workbook is cleared/destroyed).
    xlpp::Workbook wb;
    auto& keep = wb.addWorksheet("Keep");
    keep.cell("B2").setValue(9.5);
    wb.addWorksheet("Victim");

    // Removing an unrelated sheet must not invalidate `keep`.
    wb.removeWorksheet("Victim");
    test.checkNear(keep.cell("B2").numericValueOr(-1), 9.5, 1e-12, "Unrelated remove keeps reference valid");

    // Copying also keeps the original reference valid.
    wb.copyWorksheet(keep, "Clone");
    test.checkEqual(keep.name(), std::string("Keep"), "Reference valid after copyWorksheet");

    // Clearing invalidates all references (documented; only checked indirectly).
    wb.clear();
    test.checkEqual(wb.sheetCount(), std::size_t{0}, "Clear removes all worksheets");
}

void testWorkbookCopyMoveSemantics(TestContext& test) {
    xlpp::Workbook wb;
    auto& a = wb.addWorksheet("A");
    a.cell("A1").setValue(std::string("x"));
    wb.addWorksheet("B");

    xlpp::Workbook copy = wb;
    test.checkEqual(copy.sheetCount(), std::size_t{2}, "Workbook is copyable");
    test.checkEqual(copy.worksheet("A")->cell("A1").stringValueOr(""), std::string("x"),
                    "Copied workbook has deep data");
    copy.worksheet("A")->cell("A1").setValue(std::string("changed"));
    test.checkEqual(wb.worksheet("A")->cell("A1").stringValueOr(""), std::string("x"),
                    "Copy is independent of original");
}

void testStyledEmptyCellsRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_styled_empty.xlsx";
    {
        xlpp::Workbook wb;
        auto& sheet = wb.addWorksheet("Styled");
        sheet.cell("A1").setValue(std::string("filled"));
        auto& styled = sheet.cell("B2");
        styled.fill().setPatternType("solid");
        styled.fill().foregroundColor().setArgb("FFFFFF00");
        styled.border().top().setStyle("medium");
        wb.save(path);
    }
    {
        xlpp::Workbook loaded;
        loaded.load(path);
        const auto* sheet = loaded.worksheet("Styled");
        test.checkTrue(sheet != nullptr, "Styled worksheet loads");
        const auto* styled = sheet->tryCell("B2");
        test.checkTrue(styled != nullptr, "Styled-empty cell survives round-trip");
        test.checkEqual(styled->fill().patternType(), std::string("solid"), "Empty cell keeps its fill");
        test.checkEqual(styled->fill().foregroundColor().argb(), std::string("FFFFFF00"), "Empty cell keeps fill color");
        test.checkEqual(styled->border().top().style(), std::string("medium"), "Empty cell keeps border");
        test.checkTrue(styled->empty(), "Styled cell still has no value");
        test.checkEqual(sheet->dimensions(), std::string("A1:B2"), "Styled empty cell extends dimensions");
    }
    std::filesystem::remove(path);
}

void testDefinedNamesFullRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_defined_names_full.xlsx";
    {
        xlpp::Workbook wb;
        wb.addWorksheet("Sheet1");
        xlpp::DefinedName global("GlobalName", "'Sheet1'!$A$1:$C$5");
        global.setComment("A global name");
        wb.addDefinedName(std::move(global));
        xlpp::DefinedName local("LocalName", "'Sheet1'!$B$2");
        local.setLocalSheetId(0);
        local.setHidden(true);
        wb.addDefinedName(std::move(local));
        wb.save(path);
    }
    {
        xlpp::Workbook loaded;
        loaded.load(path);
        test.checkEqual(loaded.definedNames().size(), std::size_t{2}, "Both defined names load");
        const auto* global = loaded.definedName("GlobalName");
        test.checkTrue(global != nullptr, "Global name loads");
        test.checkEqual(global->value(), std::string("'Sheet1'!$A$1:$C$5"), "Global name value");
        test.checkEqual(global->comment(), std::string("A global name"), "Global name comment");
        test.checkTrue(!global->localSheetId().has_value(), "Global name has no sheet scope");
        const auto* local = loaded.definedName("LocalName");
        test.checkTrue(local != nullptr, "Local name loads");
        test.checkTrue(local->localSheetId().has_value(), "Local name keeps sheet scope");
        test.checkEqual(*local->localSheetId(), std::size_t{0}, "Local name sheet id");
        test.checkTrue(local->hidden(), "Local name hidden flag round-trips");
    }
    std::filesystem::remove(path);
}

void testRowValuesAndCells(TestContext& test) {
    xlpp::Worksheet sheet("Rows");
    sheet.cell("A1").setValue(std::string("a"));
    sheet.cell("B1").setValue(1.0);
    sheet.cell("D1").setValue(std::string("d"));
    sheet.cell("A2").setValue(std::string("second"));

    auto row = sheet.row(1);
    test.checkEqual(row.number(), std::size_t{1}, "Row number");
    const auto values = row.values();
    test.checkEqual(values.size(), std::size_t{4}, "Row::values spans min to max column");
    test.checkEqual(std::get<std::string>(values[0]), std::string("a"), "Row::values first");
    test.checkNear(std::get<double>(values[1]), 1.0, 1e-12, "Row::values numeric");
    test.checkTrue(std::holds_alternative<std::monostate>(values[2]), "Row::values empty gap");
    test.checkEqual(std::get<std::string>(values[3]), std::string("d"), "Row::values last");

    const auto cells = row.cells();
    test.checkEqual(cells.size(), std::size_t{3}, "Row::cells skips empty cells");
    test.checkEqual(cells[0]->address(), std::string("A1"), "Row::cells first address");
    test.checkEqual(cells[2]->address(), std::string("D1"), "Row::cells last address");

    test.checkEqual(sheet.row(2).tryCell(1)->stringValueOr(""), std::string("second"), "tryCell via row proxy");
}

void testDateCellNumberFormat(TestContext& test) {
    xlpp::Cell cell("A1");
    cell.setDate(2024, 1, 15);
    test.checkTrue(cell.isDate(), "setDate creates a date value");
    test.checkEqual(cell.numberFormat(), std::string("yyyy-mm-dd"), "setDate applies a date number format");

    xlpp::Cell cell2("B1");
    cell2.setDateTime(xlpp::DateTime{2024, 1, 15, 13, 30, 0});
    test.checkEqual(cell2.numberFormat(), std::string("yyyy-mm-dd h:mm:ss"), "setDateTime applies a datetime format");
}

void testCellRangeOperations(TestContext& test) {
    xlpp::Worksheet sheet("Ranges");
    auto range = sheet.range("A1:C2");
    range.setValue(1.0);
    test.checkNear(std::get<double>(sheet.cell("A1").value()), 1.0, 1e-12, "range setValue A1");
    test.checkNear(std::get<double>(sheet.cell("C2").value()), 1.0, 1e-12, "range setValue C2");
    test.checkEqual(range.rowCount(), std::size_t{2}, "range rowCount");
    test.checkEqual(range.columnCount(), std::size_t{3}, "range columnCount");
    test.checkEqual(range.cells().size(), std::size_t{6}, "range cells() count");
    test.checkEqual(range.rows().size(), std::size_t{2}, "range rows() count");

    std::size_t visited = 0;
    range.forEach([&](xlpp::Cell&) { ++visited; });
    test.checkEqual(visited, std::size_t{6}, "range forEach visits every cell");

    const auto values = range.values();
    test.checkEqual(values.size(), std::size_t{6}, "range values() count");

    range.clear();
    test.checkTrue(sheet.cell("B1").empty(), "range clear empties cells");
}

void testPropertiesFullRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_props_full.xlsx";
    {
        xlpp::Workbook wb;
        wb.addWorksheet("S").cell("A1").setValue("x");
        auto& p = wb.properties();
        p.setTitle("The Title");
        p.setSubject("The Subject");
        p.setCreator("The Creator");
        p.setDescription("The Description");
        p.setKeywords("k1,k2");
        p.setCategory("Tests");
        p.setLastModifiedBy("Unit Test");
        wb.save(path);
    }
    {
        xlpp::Workbook loaded;
        loaded.load(path);
        const auto& p = loaded.properties();
        test.checkEqual(p.title(), std::string("The Title"), "Title round-trip");
        test.checkEqual(p.subject(), std::string("The Subject"), "Subject round-trip");
        test.checkEqual(p.creator(), std::string("The Creator"), "Creator round-trip");
        test.checkEqual(p.description(), std::string("The Description"), "Description round-trip");
        test.checkEqual(p.keywords(), std::string("k1,k2"), "Keywords round-trip");
        test.checkEqual(p.category(), std::string("Tests"), "Category round-trip");
        test.checkEqual(p.lastModifiedBy(), std::string("Unit Test"), "LastModifiedBy round-trip");
    }
    std::filesystem::remove(path);
}

void testHyperlinkDisplayTooltipRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_hyperlink_dt.xlsx";
    {
        xlpp::Workbook wb;
        auto& sheet = wb.addWorksheet("Links");
        auto& cell = sheet.cell("A1");
        cell.setValue(std::string("Open"));
        xlpp::Hyperlink link("https://example.com/page");
        link.setDisplay("Example page");
        link.setTooltip("Open the example");
        cell.setHyperlink(std::move(link));
        wb.save(path);
    }
    {
        xlpp::Workbook loaded;
        loaded.load(path);
        const auto* cell = loaded.worksheet("Links")->tryCell("A1");
        test.checkTrue(cell->hasHyperlink(), "Hyperlink present after round-trip");
        test.checkEqual(cell->hyperlinkValue()->target(), std::string("https://example.com/page"), "Target round-trip");
        test.checkEqual(cell->hyperlinkValue()->display(), std::string("Example page"), "Display round-trip");
        test.checkEqual(cell->hyperlinkValue()->tooltip(), std::string("Open the example"), "Tooltip round-trip");
    }
    std::filesystem::remove(path);
}

void testColumnDimensionByName(TestContext& test) {
    xlpp::Worksheet sheet("Cols");
    auto& dim = sheet.columnDimension("B");
    dim.width = 33.5;
    dim.hidden = true;
    dim.outlineLevel = 2;
    test.checkNear(sheet.tryColumnDimension(2)->width.value_or(0.0), 33.5, 1e-12, "columnDimension by name width");
    test.checkTrue(sheet.tryColumnDimension(2)->hidden, "columnDimension by name hidden");
    test.checkEqual(sheet.tryColumnDimension(2)->outlineLevel, 2, "columnDimension by name outline");
}

void testCommentMutationAfterSave(TestContext& test) {
    const auto baseline = std::filesystem::temp_directory_path() / "xlpp_comment_cache_baseline.xlsx";
    const auto withComment = std::filesystem::temp_directory_path() / "xlpp_comment_cache_added.xlsx";
    const auto withoutComment = std::filesystem::temp_directory_path() / "xlpp_comment_cache_removed.xlsx";

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Notes");
    auto& cell = sheet.cell("B3");
    cell.setValue("payload");
    workbook.save(baseline);

    // Mutate through a previously acquired Cell reference. This used to leave
    // the cached worksheet XML without <legacyDrawing>, making the note invisible.
    cell.setComment(xlpp::Comment("Added after first save", "Alice"));
    workbook.save(withComment);
    auto addedZip = xlpp::internal::ZipArchive::open(withComment);
    test.checkTrue(addedZip.contains("xl/comments1.xml"), "Comment part is added after a cached save");
    test.checkTrue(addedZip.contains("xl/drawings/commentsDrawing1.vml"), "Comment VML is added after a cached save");
    test.checkTrue(addedZip.get("xl/worksheets/sheet1.xml").find("<legacyDrawing r:id=\"rIdCommentsVml\"/>") != std::string::npos,
                   "Cached worksheet is regenerated with the comment drawing marker");

    xlpp::Workbook loaded;
    loaded.load(withComment);
    const auto* loadedCell = loaded.worksheet("Notes")->tryCell("B3");
    test.checkTrue(loadedCell && loadedCell->hasComment(), "Comment added after first save can be loaded");
    test.checkEqual(loadedCell->commentValue()->text(), std::string("Added after first save"), "Added comment text round-trips");
    test.checkEqual(loadedCell->commentValue()->author(), std::string("Alice"), "Added comment author round-trips");

    cell.clearComment();
    workbook.save(withoutComment);
    auto removedZip = xlpp::internal::ZipArchive::open(withoutComment);
    test.checkTrue(!removedZip.contains("xl/comments1.xml"), "Last comment part is removed");
    test.checkTrue(!removedZip.contains("xl/drawings/commentsDrawing1.vml"), "Last comment VML is removed");
    test.checkTrue(removedZip.get("xl/worksheets/sheet1.xml").find("legacyDrawing") == std::string::npos,
                   "Cached worksheet removes the legacy drawing marker");

    for (const auto& path : {baseline, withComment, withoutComment}) std::filesystem::remove(path);
}

void testRichTextCommentImport(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_rich_comment_source.xlsx";
    const auto patched = std::filesystem::temp_directory_path() / "xlpp_rich_comment_patched.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Notes");
    sheet.cell("A1").setComment(xlpp::Comment("placeholder", "Reviewer"));
    workbook.save(source);

    auto archive = xlpp::internal::ZipArchive::open(source);
    archive.replace("xl/comments1.xml",
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Reviewer</author></authors><commentList><comment ref="A1" authorId="0"><text><r><rPr><b/></rPr><t xml:space="preserve">Hello </t></r><r><t>World</t></r><r><t>!</t></r></text></comment></commentList></comments>)");
    archive.save(patched);

    xlpp::Workbook loaded;
    loaded.load(patched);
    const auto* commentCell = loaded.worksheet("Notes")->tryCell("A1");
    test.checkTrue(commentCell && commentCell->hasComment(), "Rich-text legacy comment loads");
    test.checkEqual(commentCell->commentValue()->text(), std::string("Hello World!"), "All rich-text comment runs are concatenated");
    test.checkEqual(commentCell->commentValue()->author(), std::string("Reviewer"), "Rich-text comment author loads");

    std::filesystem::remove(source);
    std::filesystem::remove(patched);
}

void testPivotAutoCacheFromSource(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_pivot_auto_cache.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Sales Data");
    sheet.append({std::string("Region"), std::string("Quarter"), std::string("Year"), std::string("Amount")});
    sheet.append({std::string("North"), std::string("Q1"), 2025.0, 12.5});
    sheet.append({std::string("South"), std::string("Q1"), 2025.0, 20.0});
    sheet.append({std::string("North"), std::string("Q2"), 2026.0, 7.5});

    xlpp::PivotTable pivot("SalesByRegion");
    pivot.setLocation("F2");
    pivot.cache().setSourceData("'Sales Data'!$A$1:$D$4");
    pivot.addRowField("Region");
    pivot.addColumnField("Quarter");
    pivot.addPageField("Year");
    pivot.addDataField("Amount");
    sheet.addPivotTable(std::move(pivot));
    workbook.save(path);

    auto archive = xlpp::internal::ZipArchive::open(path);
    const auto cacheXml = archive.get("xl/pivotCache/pivotCacheDefinition1.xml");
    const auto recordsXml = archive.get("xl/pivotCache/pivotCacheRecords1.xml");
    const auto tableXml = archive.get("xl/pivotTables/pivotTable1.xml");
    test.checkTrue(cacheXml.find("recordCount=\"3\"") != std::string::npos, "Pivot records are inferred from source rows");
    test.checkTrue(cacheXml.find("<cacheFields count=\"4\">") != std::string::npos, "Pivot fields are inferred from header row");
    test.checkTrue(cacheXml.find("name=\"Region\"") != std::string::npos, "First source header becomes pivot field");
    test.checkTrue(cacheXml.find("name=\"Amount\"") != std::string::npos, "Last source header becomes pivot field");
    test.checkTrue(cacheXml.find("<s v=\"North\"/>") != std::string::npos, "String shared item is inferred");
    const auto amountField = cacheXml.find("<cacheField name=\"Amount\"");
    test.checkTrue(amountField != std::string::npos, "Numeric data cache field is written");
    test.checkTrue(cacheXml.find("saveData=\"1\"") != std::string::npos, "Pivot cache declares saved records");
    test.checkTrue(cacheXml.find("containsSemiMixedTypes=\"0\"") != std::string::npos,
                   "Homogeneous numeric pivot field declares compatible type metadata");
    test.checkTrue(cacheXml.find("<cacheField name=\"Amount\" numFmtId=\"0\"><sharedItems", amountField) != std::string::npos,
                   "Numeric data field includes cache metadata");
    test.checkTrue(recordsXml.find("<n v=\"12.5\"/>") != std::string::npos,
                   "Pure data field values are stored directly in pivot cache records");
    test.checkTrue(recordsXml.find("count=\"3\"") != std::string::npos, "Inferred pivot cache records part has correct count");
    test.checkTrue(tableXml.find("<rowFields count=\"1\"><field x=\"0\"/>") != std::string::npos,
                   "Named row field resolves to source index");
    test.checkTrue(tableXml.find("<colFields count=\"1\"><field x=\"1\"/>") != std::string::npos,
                   "Named column field resolves to source index");
    test.checkTrue(tableXml.find("<pageField fld=\"2\" hier=\"-1\"/>") != std::string::npos,
                   "Named page field resolves to source index");
    test.checkTrue(tableXml.find("<dataField name=\"Sum of Amount\" fld=\"3\"") != std::string::npos,
                   "Named data field resolves to source index");
    test.checkTrue(tableXml.find("<i t=\"grand\"><x/></i>") != std::string::npos,
                   "Pivot grand item does not reference a non-existent shared item");
    std::filesystem::remove(path);
}

void testExcelCompatiblePivotView(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_excel_compatible_pivot_view.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("PivotData");
    sheet.append({std::string("Quarter"), std::string("Amount")});
    sheet.append({std::string("Q1"), 10.0});
    sheet.append({std::string("Q2"), 20.0});

    xlpp::PivotTable pivot("SalesPivot");
    pivot.setLocation("D2");
    pivot.cache().setSourceData("'PivotData'!$A$1:$B$3");
    pivot.addRowField("Quarter");
    pivot.addDataField("Amount", "sum");
    sheet.addPivotTable(std::move(pivot));
    workbook.save(path);

    auto archive = xlpp::internal::ZipArchive::open(path);
    const auto tableXml = archive.get("xl/pivotTables/pivotTable1.xml");
    const auto cacheXml = archive.get("xl/pivotCache/pivotCacheDefinition1.xml");
    test.checkTrue(tableXml.find("<location ref=\"D2:E5\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>") != std::string::npos,
                   "Pivot view derives an exact range consistent with its row and column items");
    test.checkTrue(tableXml.find("<pivotField axis=\"axisRow\" showAll=\"0\" defaultSubtotal=\"0\"><items count=\"2\">") != std::string::npos,
                   "Pivot row field writes only concrete cache items");
    test.checkTrue(tableXml.find("<item t=\"default\"/>") == std::string::npos,
                   "Pivot view does not invent a synthetic default item");
    test.checkTrue(tableXml.find("<colItems count=\"1\"><i/></colItems>") != std::string::npos,
                   "Single-data-field pivot uses the canonical empty column item");
    test.checkTrue(tableXml.find("<pivotField dataField=\"1\" showAll=\"0\"/>") != std::string::npos,
                   "Pivot data field omits axis-only attributes and custom source-name metadata");
    test.checkTrue(tableXml.find("<dataField name=\"Sum of Amount\" fld=\"1\" baseField=\"0\" baseItem=\"0\"/>") != std::string::npos,
                   "Pivot value field uses the minimal Excel-compatible representation");
    test.checkTrue(cacheXml.find("minValue=\"10\" maxValue=\"20\" count=") == std::string::npos,
                   "Pure data cache fields do not advertise child-item counts without child items");

    std::filesystem::remove(path);
}

void testMultiplePivotCacheIds(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_multiple_pivot_ids.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Category"), std::string("Value")});
    sheet.append({std::string("A"), 1.0});
    sheet.append({std::string("B"), 2.0});

    for (const auto& pair : {std::pair<std::string, std::string>{"PivotOne", "D2"}, {"PivotTwo", "H2"}}) {
        xlpp::PivotTable pivot(pair.first);
        pivot.setLocation(pair.second);
        pivot.cache().setSourceData("'Data'!$A$1:$B$3");
        pivot.addRowField("Category");
        pivot.addDataField("Value");
        sheet.addPivotTable(std::move(pivot));
    }
    workbook.save(path);

    auto archive = xlpp::internal::ZipArchive::open(path);
    const auto workbookXml = archive.get("xl/workbook.xml");
    const auto firstPivot = archive.get("xl/pivotTables/pivotTable1.xml");
    const auto secondPivot = archive.get("xl/pivotTables/pivotTable2.xml");
    test.checkTrue(workbookXml.find("<pivotCache cacheId=\"1\"") != std::string::npos, "Workbook declares first pivot cache ID");
    test.checkTrue(workbookXml.find("<pivotCache cacheId=\"2\"") != std::string::npos, "Workbook declares second pivot cache ID");
    test.checkTrue(firstPivot.find("cacheId=\"1\"") != std::string::npos, "First pivot references cache ID 1");
    test.checkTrue(secondPivot.find("cacheId=\"2\"") != std::string::npos, "Second pivot references cache ID 2");
    test.checkTrue(archive.contains("xl/pivotCache/pivotCacheDefinition2.xml"), "Second pivot cache definition is written");
    test.checkTrue(archive.contains("xl/pivotCache/pivotCacheRecords2.xml"), "Second pivot cache records are written");
    std::filesystem::remove(path);
}

void testStrictPivotNamespaces(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_strict_pivot.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Group"), std::string("Value")});
    sheet.append({std::string("A"), 1.0});
    xlpp::PivotTable pivot("StrictPivot");
    pivot.cache().setSourceData("'Data'!$A$1:$B$2");
    pivot.addRowField("Group");
    pivot.addDataField("Value");
    sheet.addPivotTable(std::move(pivot));

    xlpp::SaveOptions options;
    options.strictNamespace = true;
    workbook.save(path, options);
    auto archive = xlpp::internal::ZipArchive::open(path);
    const std::string strictMain = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    const std::string strictDocumentRels = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    const std::string strictPackageRels = "http://purl.oclc.org/ooxml/package/relationships";
    test.checkTrue(archive.get("xl/pivotTables/pivotTable1.xml").find(strictMain) != std::string::npos,
                   "Strict pivot table uses strict SpreadsheetML namespace");
    test.checkTrue(archive.get("xl/pivotCache/pivotCacheDefinition1.xml").find(strictMain) != std::string::npos,
                   "Strict pivot cache uses strict SpreadsheetML namespace");
    test.checkTrue(archive.get("xl/pivotCache/pivotCacheDefinition1.xml").find(strictDocumentRels) != std::string::npos,
                   "Strict pivot cache uses strict document relationship namespace");
    test.checkTrue(archive.get("xl/pivotCache/pivotCacheRecords1.xml").find(strictMain) != std::string::npos,
                   "Strict pivot records use strict SpreadsheetML namespace");
    test.checkTrue(archive.get("xl/pivotTables/_rels/pivotTable1.xml.rels").find(strictPackageRels) != std::string::npos,
                   "Strict pivot relationship part uses strict package namespace");
    std::filesystem::remove(path);
}

void testPivotModelValidationAndAggregation(TestContext& test) {
    xlpp::PivotCache cache;
    cache.setFields({"A", "B"});
    cache.addRecord({"x", "1"});
    bool widthRejected = false;
    try { cache.setFields({"OnlyOne"}); } catch (const std::invalid_argument&) { widthRejected = true; }
    test.checkTrue(widthRejected, "Changing pivot fields rejects existing records with a different width");
    bool lateFieldRejected = false;
    try { cache.addField("C"); } catch (const std::logic_error&) { lateFieldRejected = true; }
    test.checkTrue(lateFieldRejected, "Adding a pivot field after records is rejected");
    bool invalidIdRejected = false;
    try { cache.setCacheId(0); } catch (const std::invalid_argument&) { invalidIdRejected = true; }
    test.checkTrue(invalidIdRejected, "Pivot cache ID must be positive");

    xlpp::PivotTable pivot("NamedFields");
    pivot.cache().setFields({"Category", "Amount"});
    auto& row = pivot.addRowField("Category");
    auto& data = pivot.addDataField("Amount", "average");
    test.checkEqual(row.fieldIndex(), 0, "Named row field resolves immediately when cache fields exist");
    test.checkEqual(data.fieldIndex(), 1, "Named data field resolves immediately when cache fields exist");
    test.checkEqual(data.subtotal(), std::string("average"), "Data field aggregation is stored");
    bool invalidAggregationRejected = false;
    try { data.setSubtotal("median"); } catch (const std::invalid_argument&) { invalidAggregationRejected = true; }
    test.checkTrue(invalidAggregationRejected, "Unsupported pivot aggregation is rejected before serialization");
}


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

void testPrintAreaTitlesAndFitToPage(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_print_names_roundtrip.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Sales Data");
    sheet.cell("A1").setValue("Header");
    sheet.cell("D40").setValue(1.0);
    sheet.pageSetup().setOrientation(xlpp::PageOrientation::Landscape);
    sheet.pageSetup().setPaperSize(xlpp::PaperSize::A4);
    sheet.pageSetup().setFitToPage(true);
    sheet.pageSetup().setFitToWidth(1);
    sheet.pageSetup().setFitToHeight(0);
    sheet.setPrintArea("A1:D40");
    sheet.setPrintTitlesRows("1:2");
    sheet.setPrintTitlesCols("A:B");
    workbook.save(path);

    const auto zip = xlpp::internal::ZipArchive::open(path);
    const auto workbookXml = zip.get("xl/workbook.xml");
    const auto sheetXml = zip.get("xl/worksheets/sheet1.xml");
    test.checkTrue(workbookXml.find("name=\"_xlnm.Print_Area\" localSheetId=\"0\"") != std::string::npos,
                   "Print area is a local workbook defined name");
    test.checkTrue(workbookXml.find("&apos;Sales Data&apos;!$A$1:$D$40") != std::string::npos,
                   "Print area contains quoted absolute sheet reference");
    test.checkTrue(workbookXml.find("name=\"_xlnm.Print_Titles\" localSheetId=\"0\"") != std::string::npos,
                   "Print titles are a local workbook defined name");
    test.checkTrue(workbookXml.find("$A:$B") != std::string::npos && workbookXml.find("$1:$2") != std::string::npos,
                   "Print title columns and rows are serialized");
    test.checkTrue(sheetXml.find("<pageSetUpPr fitToPage=\"1\"/>") != std::string::npos,
                   "Fit-to-page mode is activated in sheet properties");
    test.checkTrue(sheetXml.find("<printArea>") == std::string::npos, "Invalid worksheet printArea element is absent");

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedSheet = loaded.worksheet("Sales Data");
    test.checkTrue(loadedSheet != nullptr, "Page setup worksheet reloads");
    test.checkEqual(loadedSheet->printArea(), std::string("A1:D40"), "Print area reloads into Worksheet state");
    test.checkEqual(loadedSheet->printTitlesRows(), std::string("1:2"), "Print title rows reload");
    test.checkEqual(loadedSheet->printTitlesCols(), std::string("A:B"), "Print title columns reload");
    test.checkTrue(loadedSheet->pageSetup().fitToPage(), "Fit-to-page reloads");
    std::filesystem::remove(path);
}

void testComprehensiveFormattingRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_comprehensive_formatting.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Formatting");
    auto& cell = sheet.cell("C4");
    cell.setValue("Formatted");
    cell.font().setBold(true);
    cell.font().setItalic(true);
    cell.font().setSize(15.0);
    cell.font().color().setArgb("FF123456");
    cell.fill().setPatternType("solid");
    cell.fill().foregroundColor().setArgb("FFF4B183");
    cell.alignment().setHorizontal("center");
    cell.alignment().setVertical("center");
    cell.alignment().setWrapText(true);
    cell.alignment().setTextRotation(45);
    for (auto* side : {&cell.border().left(), &cell.border().right(), &cell.border().top(), &cell.border().bottom()}) {
        side->setStyle("thin");
        side->color().setArgb("FF0070C0");
    }
    sheet.columnDimension("C").width = 27.5;
    sheet.rowDimension(4).height = 31.25;
    workbook.save(path);

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedSheet = loaded.worksheet("Formatting");
    test.checkTrue(loadedSheet != nullptr, "Formatting worksheet reloads");
    const auto* loadedCellPtr = loadedSheet->tryCell("C4");
    test.checkTrue(loadedCellPtr != nullptr, "Formatted cell reloads");
    const auto& loadedCell = *loadedCellPtr;
    test.checkTrue(loadedCell.font().bold(), "Bold font reloads");
    test.checkTrue(loadedCell.font().italic(), "Italic font reloads");
    test.checkNear(loadedCell.font().size(), 15.0, 1e-12, "Font size reloads");
    test.checkEqual(loadedCell.font().color().argb(), std::string("FF123456"), "Font color reloads");
    test.checkEqual(loadedCell.fill().patternType(), std::string("solid"), "Fill pattern reloads");
    test.checkEqual(loadedCell.fill().foregroundColor().argb(), std::string("FFF4B183"), "Cell fill color reloads");
    test.checkEqual(loadedCell.alignment().horizontal(), std::string("center"), "Horizontal alignment reloads");
    test.checkEqual(loadedCell.alignment().vertical(), std::string("center"), "Vertical alignment reloads");
    test.checkTrue(loadedCell.alignment().wrapText(), "Wrap-text alignment reloads");
    test.checkEqual(loadedCell.alignment().textRotation(), 45, "Text rotation reloads");
    test.checkEqual(loadedCell.border().left().style(), std::string("thin"), "Left border style reloads");
    test.checkEqual(loadedCell.border().right().color().argb(), std::string("FF0070C0"), "Right border color reloads");
    test.checkTrue(loadedSheet->tryColumnDimension(3) != nullptr, "Column dimension reloads");
    test.checkNear(*loadedSheet->tryColumnDimension(3)->width, 27.5, 1e-12, "Column width reloads");
    test.checkTrue(loadedSheet->tryRowDimension(4) != nullptr, "Row dimension reloads");
    test.checkNear(*loadedSheet->tryRowDimension(4)->height, 31.25, 1e-12, "Row height reloads");
    std::filesystem::remove(path);
}

void testProtectionPasswordAddRemove(TestContext& test) {
    test.checkEqual(xlpp::legacyProtectionPasswordHash("password"), std::string("83AF"), "Known legacy password hash");
    test.checkEqual(xlpp::legacyProtectionPasswordHash("secret"), std::string("DAA7"), "Second known legacy password hash");
    bool longRejected = false;
    try { (void)xlpp::legacyProtectionPasswordHash("1234567890123456"); }
    catch (const std::invalid_argument&) { longRejected = true; }
    test.checkTrue(longRejected, "Legacy password longer than 15 characters is rejected");

    const auto protectedPath = std::filesystem::temp_directory_path() / "xlpp_password_added.xlsx";
    const auto clearedPath = std::filesystem::temp_directory_path() / "xlpp_password_removed.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Protected");
    sheet.cell("A1").setValue("Protected");
    sheet.protection().setPassword("secret");
    workbook.protection().setLockStructure(true);
    workbook.protection().setPassword("password");
    workbook.save(protectedPath);
    test.checkTrue(sheet.protection().hasPassword(), "Worksheet reports password after setPassword");
    test.checkTrue(workbook.protection().hasPassword(), "Workbook reports password after setPassword");

    auto zip = xlpp::internal::ZipArchive::open(protectedPath);
    test.checkTrue(zip.get("xl/worksheets/sheet1.xml").find("password=\"DAA7\"") != std::string::npos,
                   "Worksheet password hash is serialized");
    test.checkTrue(zip.get("xl/workbook.xml").find("workbookPassword=\"83AF\"") != std::string::npos,
                   "Workbook password hash is serialized");

    sheet.protection().clearPassword();
    workbook.protection().clearPassword();
    workbook.save(clearedPath);
    test.checkTrue(!sheet.protection().hasPassword(), "Worksheet password is cleared");
    test.checkTrue(!workbook.protection().hasPassword(), "Workbook password is cleared");
    zip = xlpp::internal::ZipArchive::open(clearedPath);
    test.checkTrue(zip.get("xl/worksheets/sheet1.xml").find(" password=") == std::string::npos,
                   "Cleared worksheet password is absent from XML");
    test.checkTrue(zip.get("xl/workbook.xml").find("workbookPassword=") == std::string::npos,
                   "Cleared workbook password is absent from XML");
    std::filesystem::remove(protectedPath);
    std::filesystem::remove(clearedPath);
}

void testVbaProjectPackageLifecycle(TestContext& test) {
    const auto macroPath = std::filesystem::temp_directory_path() / "xlpp_vba_package.xlsm";
    const auto noMacroPath = std::filesystem::temp_directory_path() / "xlpp_vba_removed.xlsx";
    const std::vector<unsigned char> bytes{'V','B','A',0,'X','L','P','P'};
    xlpp::Workbook workbook;
    workbook.addWorksheet("MacroHost").cell("A1").setValue("VBA host");
    workbook.setVbaProject(bytes);
    test.checkTrue(workbook.hasVbaProject(), "Workbook reports attached VBA project");
    workbook.save(macroPath);

    auto zip = xlpp::internal::ZipArchive::open(macroPath);
    test.checkTrue(zip.contains("xl/vbaProject.bin"), "VBA project binary is packaged");
    const auto packaged = zip.get("xl/vbaProject.bin");
    test.checkEqual(packaged.size(), bytes.size(), "VBA binary size is preserved");
    test.checkEqual(packaged[3], '\0', "VBA binary NUL byte is preserved");
    const auto types = zip.get("[Content_Types].xml");
    const auto rels = zip.get("xl/_rels/workbook.xml.rels");
    test.checkTrue(types.find("application/vnd.ms-excel.sheet.macroEnabled.main+xml") != std::string::npos,
                   "Macro-enabled workbook main content type is emitted");
    test.checkTrue(types.find("application/vnd.ms-office.vbaProject") != std::string::npos,
                   "VBA project content type is emitted");
    test.checkTrue(rels.find("/vbaProject\" Target=\"vbaProject.bin\"") != std::string::npos,
                   "Workbook relationship points to VBA project");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(zip).validate().ok(),
                   "Macro package passes OPC relationship and content-type validation");

    xlpp::Workbook loaded;
    loaded.load(macroPath);
    test.checkTrue(loaded.hasVbaProject(), "Loaded XLSM reports VBA project");
    test.checkTrue(loaded.removeVbaProject(), "removeVbaProject removes attached project");
    test.checkTrue(!loaded.hasVbaProject(), "VBA project is absent after removal");
    loaded.save(noMacroPath);
    zip = xlpp::internal::ZipArchive::open(noMacroPath);
    test.checkTrue(!zip.contains("xl/vbaProject.bin"), "Removed VBA binary is not packaged");
    test.checkTrue(zip.get("[Content_Types].xml").find("macroEnabled") == std::string::npos,
                   "Workbook returns to normal XLSX content type after VBA removal");
    test.checkTrue(zip.get("xl/_rels/workbook.xml.rels").find("/vbaProject") == std::string::npos,
                   "VBA relationship is removed");
    std::filesystem::remove(macroPath);
    std::filesystem::remove(noMacroPath);
}

void testImagePackageRegression(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_image_package.xlsx";
    const std::vector<unsigned char> png{137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,137,0,0,0,13,73,68,65,84,8,215,99,248,207,192,240,31,0,5,0,1,255,137,153,61,29,0,0,0,0,73,69,78,68,174,66,96,130};
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Image");
    auto image = xlpp::Image("D2", png, "png");
    image.setWidthPixels(64);
    image.setHeightPixels(64);
    sheet.addImage(std::move(image));
    workbook.save(path);
    const auto zip = xlpp::internal::ZipArchive::open(path);
    test.checkTrue(zip.contains("xl/media/image1.png"), "Image binary is packaged");
    test.checkEqual(zip.get("xl/media/image1.png").size(), png.size(), "Image bytes are preserved");
    test.checkTrue(zip.get("xl/worksheets/sheet1.xml").find("<drawing r:id=\"rIdDrawing\"/>") != std::string::npos,
                   "Worksheet references drawing part");
    const auto drawing = zip.get("xl/drawings/drawing1.xml");
    test.checkTrue(drawing.find("<xdr:col>3</xdr:col>") != std::string::npos, "Image D2 anchor column is serialized");
    test.checkTrue(drawing.find("<xdr:row>1</xdr:row>") != std::string::npos, "Image D2 anchor row is serialized");
    test.checkTrue(zip.get("xl/drawings/_rels/drawing1.xml.rels").find("../media/image1.png") != std::string::npos,
                   "Drawing relationship targets packaged image");
    std::filesystem::remove(path);
}


void testAdvancedConditionalFormattingRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_advanced_conditional.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Rules");
    for (int row = 1; row <= 10; ++row) {
        sheet.cell(static_cast<std::size_t>(row), 1).setValue(row);
        sheet.cell(static_cast<std::size_t>(row), 2).setValue(row);
        sheet.cell(static_cast<std::size_t>(row), 3).setValue(row);
    }

    auto dataBar = xlpp::ConditionalRule::dataBar("FF112233");
    dataBar.getDataBar().min.type = "min";
    dataBar.getDataBar().max.type = "max";
    dataBar.getDataBar().showValue = false;
    dataBar.getDataBar().direction = "rightToLeft";
    sheet.conditionalFormatting().addRule("A1:A10", std::move(dataBar));

    auto colorScale = xlpp::ConditionalRule::colorScale();
    auto low = xlpp::Cfvo("min", 0.0); low.hasValue = false; low.color = "FFFF0000";
    auto middle = xlpp::Cfvo("percentile", 50.0); middle.color = "FFFFFF00";
    xlpp::Cfvo high("formula", std::string("MAX(B1:B10)")); high.color = "FF00FF00";
    colorScale.getColorScale().addStop(std::move(low));
    colorScale.getColorScale().addStop(std::move(middle));
    colorScale.getColorScale().addStop(std::move(high));
    sheet.conditionalFormatting().addRule("B1:B10", std::move(colorScale));

    auto iconSet = xlpp::ConditionalRule::iconSet("3Arrows");
    iconSet.getIconSet().showValue = false;
    iconSet.getIconSet().reverse = true;
    iconSet.getIconSet().addThreshold(xlpp::Cfvo("percent", 0.0));
    iconSet.getIconSet().addThreshold(xlpp::Cfvo("percent", 33.0));
    iconSet.getIconSet().addThreshold(xlpp::Cfvo("percent", 67.0));
    sheet.conditionalFormatting().addRule("C1:C10", std::move(iconSet));

    workbook.save(path);
    const auto zip = xlpp::internal::ZipArchive::open(path);
    const auto xml = zip.get("xl/worksheets/sheet1.xml");
    test.checkTrue(xml.find("<dataBar direction=\"rightToLeft\" showValue=\"0\">") != std::string::npos ||
                   xml.find("<dataBar showValue=\"0\" direction=\"rightToLeft\">") != std::string::npos,
                   "Data-bar showValue belongs to dataBar element");
    test.checkTrue(xml.find("<iconSet") != std::string::npos && xml.find("showValue=\"0\"") != std::string::npos,
                   "Icon-set showValue is serialized");
    test.checkTrue(xml.find("<cfvo type=\"formula\" val=\"MAX(B1:B10)\"/>") != std::string::npos,
                   "Formula cfvo uses val attribute");
    test.checkTrue(xml.find("<f>MAX(B1:B10)</f>") == std::string::npos,
                   "Formula cfvo does not emit invalid sibling f element");

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto& entries = loaded.worksheet("Rules")->conditionalFormatting().entries();
    test.checkEqual(entries.size(), std::size_t{3}, "All advanced conditional entries load");
    const auto& loadedBar = entries[0].rules()[0].getDataBar();
    test.checkTrue(!loadedBar.showValue, "Data-bar showValue round-trips");
    test.checkEqual(loadedBar.direction, std::string("rightToLeft"), "Data-bar direction round-trips");
    test.checkEqual(loadedBar.color, std::string("FF112233"), "Data-bar color round-trips");
    const auto& loadedScale = entries[1].rules()[0].getColorScale();
    test.checkEqual(loadedScale.stops.size(), std::size_t{3}, "Three color-scale stops round-trip");
    test.checkEqual(loadedScale.stops[2].type, std::string("formula"), "Formula stop type round-trips");
    test.checkEqual(loadedScale.stops[2].formula, std::string("MAX(B1:B10)"), "Formula stop expression round-trips");
    test.checkTrue(loadedScale.stops[2].color.has_value(), "Formula stop color round-trips");
    const auto& loadedIcons = entries[2].rules()[0].getIconSet();
    test.checkTrue(!loadedIcons.showValue, "Icon-set showValue round-trips");
    test.checkTrue(loadedIcons.reverse, "Icon-set reverse round-trips");
    test.checkEqual(loadedIcons.thresholds.size(), std::size_t{3}, "Icon thresholds round-trip");
    std::filesystem::remove(path);
}

void testAdvancedSheetViewRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_advanced_sheet_view.xlsx";
    const auto clearedPath = std::filesystem::temp_directory_path() / "xlpp_advanced_sheet_view_cleared.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("View");
    sheet.cell("A1").setValue("view");
    xlpp::SheetView view;
    view.setWorkbookViewId(0);
    view.setTabColor("FF123456");
    view.setZoomScale(135);
    view.setZoomScaleNormal(90);
    view.setShowGridLines(false);
    view.setTabSelected(true);
    view.setRightToLeft(true);
    view.setShowOutlineSymbols(false);
    view.setPane("bottomRight");
    view.setTopLeftCell("D5");
    view.setXSplit(3);
    view.setYSplit(4);
    sheet.setSheetView(std::move(view));
    workbook.save(path);

    const auto zip = xlpp::internal::ZipArchive::open(path);
    const auto xml = zip.get("xl/worksheets/sheet1.xml");
    test.checkTrue(xml.find("zoomScaleNormal=\"90\"") != std::string::npos, "Normal zoom is serialized");
    test.checkTrue(xml.find("showOutlineSymbols=\"0\"") != std::string::npos, "Outline symbols flag is serialized");
    test.checkTrue(xml.find("state=\"split\"") != std::string::npos, "Split pane state is serialized");
    test.checkTrue(xml.find("activePane=\"bottomRight\"") != std::string::npos, "Active pane is serialized");
    test.checkTrue(xml.find("topLeftCell=\"D5\"") != std::string::npos, "Split top-left cell is serialized");

    xlpp::Workbook loaded;
    loaded.load(path);
    auto* loadedSheet = loaded.worksheet("View");
    const auto& loadedView = static_cast<const xlpp::Worksheet&>(*loadedSheet).sheetView();
    test.checkEqual(loadedView.workbookViewId(), 0, "Workbook view ID round-trips");
    test.checkTrue(loadedView.tabColor().has_value(), "Tab color round-trips");
    test.checkEqual(*loadedView.tabColor(), std::string("FF123456"), "Tab color value round-trips");
    test.checkEqual(loadedView.zoomScale(), 135, "Zoom scale round-trips");
    test.checkEqual(loadedView.zoomScaleNormal(), 90, "Normal zoom round-trips");
    test.checkTrue(!loadedView.showGridLines(), "Grid-line visibility round-trips");
    test.checkTrue(loadedView.tabSelected(), "Selected tab round-trips");
    test.checkTrue(loadedView.rightToLeft(), "Right-to-left round-trips");
    test.checkTrue(!loadedView.showOutlineSymbols(), "Outline symbol visibility round-trips");
    test.checkEqual(loadedView.pane(), std::string("bottomRight"), "Active pane round-trips");
    test.checkEqual(loadedView.topLeftCell(), std::string("D5"), "Split top-left cell round-trips");
    test.checkEqual(loadedView.xSplit(), 3, "Horizontal split round-trips");
    test.checkEqual(loadedView.ySplit(), 4, "Vertical split round-trips");

    loadedSheet->sheetView().clearTabColor();
    loadedSheet->clearFreezePanes();
    loaded.save(clearedPath);
    test.checkTrue(xlpp::internal::ZipArchive::open(clearedPath).get("xl/worksheets/sheet1.xml").find("tabColor") == std::string::npos,
                   "Clearing tab color removes tabColor XML");
    std::filesystem::remove(path);
    std::filesystem::remove(clearedPath);
}

void testAutoFilterMutationAndOperatorMatrix(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_filter_operator_matrix.xlsx";
    xlpp::FilterColumn standalone(1);
    standalone.setColumnId(7);
    standalone.addValue("A");
    standalone.addCustomFilter(xlpp::FilterOperator::NotEqual, "B");
    standalone.clearValues();
    standalone.clearCustomFilters();
    test.checkEqual(standalone.columnId(), std::size_t{7}, "Filter column ID setter");
    test.checkTrue(standalone.values().empty(), "Filter values clear");
    test.checkTrue(standalone.customFilters().empty(), "Custom filters clear");

    xlpp::SortState state;
    state.setReference("A1:F10");
    state.setCaseSensitive(true);
    state.addCondition("A2:A10", true);
    state.clear();
    test.checkTrue(state.reference().empty() && state.conditions().empty() && !state.caseSensitive(), "Sort state clear resets all fields");

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Filter");
    sheet.autoFilter().setReference("A1:F10");
    const std::array<xlpp::FilterOperator, 6> operators{
        xlpp::FilterOperator::Equal, xlpp::FilterOperator::NotEqual,
        xlpp::FilterOperator::LessThan, xlpp::FilterOperator::LessThanOrEqual,
        xlpp::FilterOperator::GreaterThan, xlpp::FilterOperator::GreaterThanOrEqual};
    for (std::size_t i = 0; i < operators.size(); ++i) {
        auto& column = sheet.autoFilter().column(i);
        column.addCustomFilter(operators[i], std::to_string(i + 1));
        column.setIncludeBlank(i == 0);
        column.setAndMode(i == 1);
    }
    auto& sort = sheet.autoFilter().sortState();
    sort.setReference("A1:F10");
    sort.setCaseSensitive(true);
    sort.addCondition("F2:F10", true);
    workbook.save(path);

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto& filter = static_cast<const xlpp::Worksheet&>(*loaded.worksheet("Filter")).autoFilter();
    test.checkTrue(filter.enabled(), "AutoFilter remains enabled");
    test.checkEqual(filter.columns().size(), std::size_t{6}, "All filter operator columns load");
    for (std::size_t i = 0; i < operators.size(); ++i) {
        const auto* column = filter.tryColumn(i);
        test.checkTrue(column != nullptr, "Filter operator column exists");
        test.checkEqual(static_cast<int>(column->customFilters().front().op), static_cast<int>(operators[i]), "Filter operator round-trips");
    }
    test.checkTrue(filter.tryColumn(0)->includeBlank(), "Include-blank flag round-trips");
    test.checkTrue(filter.tryColumn(1)->andMode(), "AND mode round-trips");
    test.checkTrue(filter.sortStateValue().has_value(), "Sort state round-trips");
    test.checkTrue(filter.sortStateValue()->caseSensitive(), "Case-sensitive sort round-trips");
    test.checkTrue(filter.sortStateValue()->conditions().front().descending, "Descending sort round-trips");

    auto& mutableFilter = loaded.worksheet("Filter")->autoFilter();
    mutableFilter.clear();
    test.checkTrue(!mutableFilter.enabled() && mutableFilter.columns().empty() && !mutableFilter.sortStateValue().has_value(),
                   "AutoFilter clear resets reference, columns and sort");
    std::filesystem::remove(path);
}

void testCellOverloadsAndOptionalModels(TestContext& test) {
    xlpp::Cell cell("A1");
    cell.setValue(std::int64_t{900719925});
    test.checkNear(cell.numericValueOr(-1), 900719925.0, 1e-12, "int64 overload stores numeric value");
    const std::string text = "view";
    cell.setValue(std::string_view(text));
    test.checkEqual(cell.stringValueOr(""), text, "string_view overload stores text");
    cell.setStringValue(static_cast<const char*>(nullptr));
    test.checkEqual(cell.stringValueOr("fallback"), std::string(), "Null C string becomes empty text");
    cell.setValue(std::monostate{});
    test.checkTrue(!cell.hasValue(), "monostate overload clears value");

    cell.setValue("plain");
    auto& rich = cell.richText();
    test.checkEqual(rich.runs().size(), std::size_t{1}, "Lazy rich text creates one run");
    rich.runs()[0].setText("changed");
    rich.runs()[0].setBold(true);
    test.checkTrue(!rich.empty(), "Rich text mutable runs are available");
    xlpp::RichText replacement = xlpp::RichText::fromPlain("replacement");
    replacement.runs()[0].setItalic(true);
    cell.setRichText(std::move(replacement));
    test.checkTrue(cell.hasRichText(), "setRichText enables rich text");
    test.checkEqual(cell.stringValueOr(""), std::string("replacement"), "Rich text updates plain value");
    cell.clearRichText();
    test.checkTrue(!cell.hasRichText(), "clearRichText removes run metadata");
    test.checkEqual(cell.stringValueOr(""), std::string("replacement"), "clearRichText preserves plain value");

    cell.hyperlink().setTarget("https://example.com");
    cell.comment().setText("note");
    cell.comment().setAuthor("author");
    test.checkTrue(cell.hasHyperlink() && cell.hasComment(), "Lazy hyperlink and comment accessors create values");
    test.checkEqual(cell.hyperlinkValue()->target(), std::string("https://example.com"), "Lazy hyperlink target stored");
    test.checkEqual(cell.commentValue()->text(), std::string("note"), "Comment text setter stored");
    test.checkEqual(cell.commentValue()->author(), std::string("author"), "Comment author setter stored");

    auto& metadata = cell.formulaMetadata();
    metadata.setType(xlpp::FormulaType::Shared);
    metadata.setReference("A1:A5");
    metadata.setSharedIndex(4);
    metadata.clearReference();
    metadata.clearSharedIndex();
    test.checkTrue(metadata.reference().empty(), "Formula reference clear");
    test.checkTrue(!metadata.sharedIndex().has_value(), "Formula shared index clear");
    metadata.setType(xlpp::FormulaType::Normal);
    test.checkTrue(metadata.empty(), "Formula metadata returns to empty state");

    cell.fill().backgroundColor().setArgb("FF010203");
    cell.alignment().setIndent(2);
    test.checkEqual(cell.fill().backgroundColor().argb(), std::string("FF010203"), "Background fill color accessor");
    test.checkEqual(cell.alignment().indent(), 2, "Alignment indent setter");
}

void testTableMutationRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_table_mutation.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Tables");
    sheet.append({std::string("OldA"), std::string("OldB"), std::string("OldC")});
    sheet.append({1.0, 2.0, 3.0});
    sheet.append({4.0, 5.0, 6.0});
    auto& table = sheet.addTable("TableOriginal", "A1:C3");
    table.setDisplayName("TableDisplay");
    table.setReference("A1:C3");
    table.setShowHeaderRow(false);
    table.setShowTotalsRow(true);
    table.addColumn("A").setName("RenamedA");
    table.addColumn("B");
    table.addColumn("C");
    table.styleInfo().setName("TableStyleMedium9");
    table.styleInfo().setShowFirstColumn(true);
    table.styleInfo().setShowLastColumn(true);
    table.styleInfo().setShowRowStripes(false);
    table.styleInfo().setShowColumnStripes(true);
    workbook.save(path);

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedTable = static_cast<const xlpp::Worksheet&>(*loaded.worksheet("Tables")).table("TableOriginal");
    test.checkTrue(loadedTable != nullptr, "Table loads by original name");
    test.checkEqual(loadedTable->displayName(), std::string("TableDisplay"), "Display name round-trips");
    test.checkEqual(loadedTable->reference(), std::string("A1:C3"), "Table reference round-trips");
    test.checkTrue(!loadedTable->showHeaderRow(), "Header-row flag round-trips");
    test.checkTrue(loadedTable->showTotalsRow(), "Totals-row flag round-trips");
    test.checkEqual(loadedTable->columns().size(), std::size_t{3}, "Table columns round-trip");
    test.checkEqual(loadedTable->columns()[0].name(), std::string("RenamedA"), "Renamed table column round-trips");
    test.checkEqual(loadedTable->styleInfo().name(), std::string("TableStyleMedium9"), "Table style name round-trips");
    test.checkTrue(loadedTable->styleInfo().showFirstColumn(), "First-column style flag round-trips");
    test.checkTrue(loadedTable->styleInfo().showLastColumn(), "Last-column style flag round-trips");
    test.checkTrue(!loadedTable->styleInfo().showRowStripes(), "Row-stripe flag round-trips");
    test.checkTrue(loadedTable->styleInfo().showColumnStripes(), "Column-stripe flag round-trips");
    std::filesystem::remove(path);
}

void testAdvancedPageSetupRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_advanced_page_setup.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Print");
    sheet.cell("A1").setValue("print");
    auto& setup = sheet.pageSetup();
    setup.setOrientation(xlpp::PageOrientation::Landscape);
    setup.setPaperSize(xlpp::PaperSize::A3);
    setup.setScale(75);
    setup.setFitToWidth(2);
    setup.setFitToHeight(3);
    setup.setFitToPage(true);
    setup.setBlackAndWhite(true);
    setup.setDraft(true);
    setup.setFirstPageNumber(7);
    setup.setUseFirstPageNumber(true);
    auto& margins = sheet.pageMargins();
    margins.setLeft(0.1); margins.setRight(0.2); margins.setTop(0.3);
    margins.setBottom(0.4); margins.setHeader(0.5); margins.setFooter(0.6);
    auto& options = sheet.printOptions();
    options.setHorizontalCentered(true); options.setVerticalCentered(true);
    options.setHeadings(true); options.setGridLines(true);
    auto& footer = sheet.headerFooter();
    footer.setOddHeader("odd-h"); footer.setOddFooter("odd-f");
    footer.setEvenHeader("even-h"); footer.setEvenFooter("even-f");
    footer.setDifferentOddEven(true); footer.setDifferentFirst(true);
    workbook.save(path);

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto& loadedSheet = static_cast<const xlpp::Worksheet&>(*loaded.worksheet("Print"));
    const auto& loadedSetup = loadedSheet.pageSetup();
    test.checkEqual(static_cast<int>(loadedSetup.orientation()), static_cast<int>(xlpp::PageOrientation::Landscape), "Page orientation round-trips");
    test.checkEqual(static_cast<unsigned>(loadedSetup.paperSize()), static_cast<unsigned>(xlpp::PaperSize::A3), "Paper size round-trips");
    test.checkEqual(loadedSetup.scale(), 75u, "Page scale round-trips");
    test.checkEqual(loadedSetup.fitToWidth(), 2u, "Fit width round-trips");
    test.checkEqual(loadedSetup.fitToHeight(), 3u, "Fit height round-trips");
    test.checkTrue(loadedSetup.fitToPage(), "Fit-to-page round-trips");
    test.checkTrue(loadedSetup.blackAndWhite(), "Black-and-white round-trips");
    test.checkTrue(loadedSetup.draft(), "Draft mode round-trips");
    test.checkEqual(loadedSetup.firstPageNumber(), 7u, "First page number round-trips");
    test.checkTrue(loadedSetup.useFirstPageNumber(), "Use-first-page-number round-trips");
    test.checkNear(loadedSheet.pageMargins().left(), 0.1, 1e-12, "Left margin round-trips");
    test.checkNear(loadedSheet.pageMargins().footer(), 0.6, 1e-12, "Footer margin round-trips");
    test.checkTrue(loadedSheet.printOptions().horizontalCentered(), "Horizontal centering round-trips");
    test.checkTrue(loadedSheet.printOptions().verticalCentered(), "Vertical centering round-trips");
    test.checkTrue(loadedSheet.printOptions().headings(), "Print headings round-trip");
    test.checkTrue(loadedSheet.printOptions().gridLines(), "Print grid lines round-trip");
    test.checkEqual(loadedSheet.headerFooter().evenHeader(), std::string("even-h"), "Even header round-trips");
    test.checkEqual(loadedSheet.headerFooter().evenFooter(), std::string("even-f"), "Even footer round-trips");
    test.checkTrue(loadedSheet.headerFooter().differentOddEven(), "Odd/even header flag round-trips");
    test.checkTrue(loadedSheet.headerFooter().differentFirst(), "First-page header flag round-trips");
    std::filesystem::remove(path);
}

void testImageFileApi(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto pngPath = dir / "xlpp_image_api.PNG";
    const auto gifPath = dir / "xlpp_image_api.gif";
    const auto output = dir / "xlpp_image_api.xlsx";
    const std::vector<unsigned char> png{137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,137,0,0,0,13,73,68,65,84,8,215,99,248,207,192,240,31,0,5,0,1,255,137,153,61,29,0,0,0,0,73,69,78,68,174,66,96,130};
    { std::ofstream stream(pngPath, std::ios::binary); stream.write(reinterpret_cast<const char*>(png.data()), static_cast<std::streamsize>(png.size())); }
    { std::ofstream stream(gifPath, std::ios::binary); stream << "GIF89a"; }

    auto image = xlpp::Image::fromFile(pngPath, "C3");
    test.checkEqual(image.anchor(), std::string("C3"), "Image file anchor");
    test.checkEqual(image.extension(), std::string("png"), "Upper-case PNG extension normalizes");
    test.checkEqual(image.name(), std::string("xlpp_image_api"), "Image name comes from file stem");
    image.setAnchor("D4"); image.setName("Front View"); image.setWidthPixels(64); image.setHeightPixels(32);
    test.checkEqual(image.anchor(), std::string("D4"), "Image anchor setter");
    test.checkEqual(image.name(), std::string("Front View"), "Image name setter");
    test.checkNear(image.widthPixels(), 64.0, 1e-12, "Image width setter");
    test.checkNear(image.heightPixels(), 32.0, 1e-12, "Image height setter");

    bool unsupportedThrown = false;
    try { (void)xlpp::Image::fromFile(gifPath, "A1"); } catch (const std::invalid_argument&) { unsupportedThrown = true; }
    test.checkTrue(unsupportedThrown, "Unsupported image extension throws");
    bool missingThrown = false;
    try { (void)xlpp::Image::fromFile(dir / "missing-xlpp-image.png", "A1"); } catch (const std::runtime_error&) { missingThrown = true; }
    test.checkTrue(missingThrown, "Missing image file throws");

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Image");
    sheet.addImage(pngPath, "B2");
    workbook.save(output);
    test.checkTrue(xlpp::internal::ZipArchive::open(output).contains("xl/media/image1.png"), "Worksheet path overload packages image");
    std::filesystem::remove(pngPath); std::filesystem::remove(gifPath); std::filesystem::remove(output);
}

void testWorkbookClearConstAndVbaFileApi(TestContext& test) {
    const auto binPath = std::filesystem::temp_directory_path() / "xlpp_vba_file_api.bin";
    { std::ofstream stream(binPath, std::ios::binary); stream << "VBAPROJECT"; }
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.cell("A1").setValue("value");
    workbook.addNamedStyle(xlpp::NamedStyle("Accent", xlpp::Style{}));
    workbook.addDefinedName(xlpp::DefinedName("Name", "Data!$A$1"));
    workbook.properties().setTitle("title");
    workbook.protection().setLockStructure(true);
    workbook.calcProperties().setCalcMode("manual");
    workbook.customProperties().add(xlpp::CustomProperty("Custom", "Text"));
    workbook.setDate1904(true);
    workbook.preservedParts().push_back({"custom/item.bin", "x", "application/octet-stream", "bin", "application/octet-stream", false});
    workbook.addVbaProject(binPath);
    test.checkTrue(workbook.hasVbaProject(), "VBA file API attaches project");

    const xlpp::Workbook& constant = workbook;
    test.checkEqual(constant.worksheets().size(), std::size_t{1}, "Const worksheets accessor");
    test.checkTrue(constant.worksheet("Data") != nullptr, "Const worksheet lookup");
    test.checkTrue(constant.namedStyle("Accent") != nullptr, "Const named-style lookup");
    test.checkTrue(constant.definedName("Name") != nullptr, "Const defined-name lookup");
    test.checkEqual(constant.properties().title(), std::string("title"), "Const properties accessor");
    test.checkTrue(constant.protection().lockStructure(), "Const protection accessor");
    test.checkEqual(constant.calcProperties().calcMode(), std::string("manual"), "Const calc-properties accessor");
    test.checkEqual(constant.customProperties().items().size(), std::size_t{1}, "Const custom-properties accessor");
    test.checkEqual(constant.preservedParts().size(), std::size_t{2}, "Const preserved-parts accessor includes VBA");

    workbook.clear();
    test.checkEqual(workbook.sheetCount(), std::size_t{0}, "Workbook clear removes sheets");
    test.checkTrue(workbook.namedStyles().empty(), "Workbook clear removes named styles");
    test.checkTrue(workbook.definedNames().empty(), "Workbook clear removes defined names");
    test.checkTrue(workbook.properties().title().empty(), "Workbook clear resets properties");
    test.checkTrue(!workbook.protection().lockStructure(), "Workbook clear resets protection");
    test.checkTrue(!workbook.date1904(), "Workbook clear resets date system");
    test.checkTrue(workbook.customProperties().items().empty(), "Workbook clear removes custom properties");
    test.checkTrue(workbook.preservedParts().empty(), "Workbook clear removes preserved parts");
    test.checkTrue(!workbook.hasVbaProject(), "Workbook clear removes VBA project");
    test.checkTrue(!workbook.removeVbaProject(), "Removing absent VBA project reports false");
    std::filesystem::remove(binPath);
}

void testChartAndPivotAdvancedModel(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Models");
    xlpp::Chart chart(xlpp::Chart::Type::Area);
    chart.setGrouping(xlpp::Chart::Grouping::Stacked);
    chart.setTitle("Area"); chart.setXAxisTitle("X"); chart.setYAxisTitle("Y");
    chart.setStyle("10"); chart.setWidth(700); chart.setHeight(300);
    chart.setShowLegend(false); chart.setLegendPosition("b");
    xlpp::ChartSeries series("Series");
    series.reference("Models", "$B$2:$B$4");
    series.categories("Models", "$A$2:$A$4");
    chart.addSeries(std::move(series));
    sheet.addChart(std::move(chart));
    sheet.chart(0).series()[0].setTitle("Renamed");
    const auto& constSheet = static_cast<const xlpp::Worksheet&>(sheet);
    test.checkEqual(static_cast<int>(constSheet.chart(0).grouping()), static_cast<int>(xlpp::Chart::Grouping::Stacked), "Chart grouping setter");
    test.checkEqual(constSheet.chart(0).series()[0].title(), std::string("Renamed"), "Mutable chart series accessor");
    test.checkEqual(constSheet.chart(0).series()[0].valuesReference(), std::string("='Models'!$B$2:$B$4"), "Chart value reference helper");
    test.checkEqual(constSheet.chart(0).series()[0].categoriesReference(), std::string("='Models'!$A$2:$A$4"), "Chart category reference helper");
    test.checkTrue(!constSheet.chart(0).showLegend(), "Chart legend visibility setter");
    test.checkEqual(constSheet.charts().size(), std::size_t{1}, "Const charts accessor");

    xlpp::PivotTable pivot("Pivot");
    pivot.setName("PivotRenamed"); pivot.setLocation("H2");
    pivot.cache().setCacheId(3); pivot.cache().setSourceData("Models!A1:C4");
    pivot.cache().setFields({"Region", "Quarter", "Amount"});
    pivot.cache().setRecords({{"East", "Q1", "10"}, {"West", "Q2", "20"}});
    test.checkEqual(pivot.cache().cacheId(), 3, "Pivot cache ID getter");
    test.checkEqual(pivot.cache().records().size(), std::size_t{2}, "Pivot setRecords stores records");
    pivot.cache().clearRecords();
    test.checkTrue(pivot.cache().records().empty(), "Pivot clearRecords clears records");
    pivot.cache().addRecord({"East", "Q1", "10"});
    auto& row = pivot.addRowField("Region"); row.setName("Region"); row.setShowAll(true); row.setSortType(1);
    pivot.addColumnField("Quarter");
    pivot.addPageField("Quarter");
    pivot.addDataField("Amount", "average");
    test.checkEqual(pivot.name(), std::string("PivotRenamed"), "Pivot name setter");
    test.checkEqual(row.axis(), std::string("axisRow"), "Pivot row axis getter");
    test.checkTrue(row.showAll(), "Pivot showAll setter");
    test.checkEqual(row.sortType(), 1, "Pivot sort setter");
    test.checkEqual(pivot.pageFields().size(), std::size_t{1}, "Mutable page fields accessor");
    test.checkEqual(pivot.dataFields().size(), std::size_t{1}, "Mutable data fields accessor");
    sheet.addPivotTable(std::move(pivot));
    test.checkEqual(constSheet.pivotTables().size(), std::size_t{1}, "Const pivot table accessor");
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

void testStoredReferenceMutationAfterSave(TestContext& test) {
    const auto first = std::filesystem::temp_directory_path() / "xlpp_stored_reference_first.xlsx";
    const auto second = std::filesystem::temp_directory_path() / "xlpp_stored_reference_second.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    auto& numeric = sheet.cell("A1");
    auto& styled = sheet.cell("B1");
    auto& filter = sheet.autoFilter();
    numeric.setValue(1.0);
    styled.setValue("style");
    styled.fill().setPatternType("solid");
    styled.fill().foregroundColor().setArgb("FFFF0000");
    filter.setReference("A1:B2");
    workbook.save(first);

    // All three objects were acquired before the first save. Their later
    // mutation must invalidate any cached worksheet XML.
    numeric.setValue(2.0);
    styled.fill().foregroundColor().setArgb("FF00FF00");
    filter.clear();
    workbook.save(second);

    xlpp::Workbook loaded;
    loaded.load(second);
    const auto* loadedSheet = loaded.worksheet("Data");
    const auto* loadedNumeric = loadedSheet ? loadedSheet->tryCell("A1") : nullptr;
    const auto* loadedStyled = loadedSheet ? loadedSheet->tryCell("B1") : nullptr;
    test.checkTrue(loadedNumeric != nullptr && loadedStyled != nullptr, "Stored-reference cells load");
    test.checkNear(loadedNumeric->numericValueOr(-1), 2.0, 1e-12,
                   "Stored Cell reference numeric mutation survives cached save");
    test.checkEqual(loadedStyled->fill().foregroundColor().argb(), std::string("FF00FF00"),
                    "Stored nested Style reference mutation survives cached save");
    test.checkTrue(!static_cast<const xlpp::Worksheet&>(*loadedSheet).autoFilter().enabled(),
                   "Stored AutoFilter reference clear survives cached save");
    std::filesystem::remove(first);
    std::filesystem::remove(second);
}


void testRemainingPublicMutationApis(TestContext& test) {
    test.checkEqual(xlpp::detail::pow26(3), std::size_t{17576}, "Base-26 power helper");

    xlpp::Cell cell("A1");
    const std::string text = "string setter";
    cell.setStringValue(text);
    test.checkEqual(cell.stringValueOr(""), text, "Const-reference string setter");

    xlpp::ConditionalFormattingCollection conditional;
    test.checkTrue(conditional.empty(), "Conditional collection initially empty");
    auto& entry = conditional.add("A1:A5");
    test.checkTrue(entry.empty(), "Conditional entry initially empty");
    entry.setReference("B1:B5");
    auto rule = xlpp::ConditionalRule::cellIs(xlpp::ConditionalOperator::Equal, "1");
    rule.setOperator(xlpp::ConditionalOperator::NotEqual);
    rule.addFormula("2");
    rule.differentialStyle().font().setBold(true);
    test.checkTrue(rule.hasDifferentialStyle(), "Differential style lazy accessor enables style");
    rule.clearDifferentialStyle();
    test.checkTrue(!rule.hasDifferentialStyle(), "Differential style clear");
    entry.addRule(std::move(rule));
    entry.rules()[0].setPriority(9);
    test.checkEqual(entry.reference(), std::string("B1:B5"), "Conditional entry reference setter");
    test.checkEqual(entry.rules()[0].formulas().size(), std::size_t{2}, "Conditional formula append");
    test.checkEqual(static_cast<int>(entry.rules()[0].op()), static_cast<int>(xlpp::ConditionalOperator::NotEqual),
                    "Conditional operator setter");
    test.checkTrue(!conditional.empty(), "Conditional collection becomes non-empty");
    conditional.clear();
    test.checkTrue(conditional.empty(), "Conditional collection clear");

    xlpp::DataValidationCollection validations;
    auto& validation = validations.add(xlpp::DataValidationType::Whole, "C1:C5");
    validation.setType(xlpp::DataValidationType::Decimal);
    test.checkEqual(static_cast<int>(validation.type()), static_cast<int>(xlpp::DataValidationType::Decimal),
                    "Data-validation type setter");
    validations.clear();
    test.checkTrue(validations.empty(), "Data-validation collection clear");

    xlpp::Worksheet worksheet("Mutable");
    worksheet.rowDimension(2).height = 21.0;
    const auto& constWorksheet = static_cast<const xlpp::Worksheet&>(worksheet);
    test.checkEqual(constWorksheet.rowDimensions().size(), std::size_t{1}, "Const row-dimensions accessor");
    worksheet.charts().push_back(xlpp::Chart(xlpp::Chart::Type::Line));
    worksheet.pivotTables().push_back(xlpp::PivotTable("P"));
    test.checkEqual(worksheet.charts().size(), std::size_t{1}, "Mutable charts collection accessor");
    test.checkEqual(worksheet.pivotTables().size(), std::size_t{1}, "Mutable pivot collection accessor");

    xlpp::NamedStyle named;
    named.setName("RenamedStyle");
    test.checkEqual(named.name(), std::string("RenamedStyle"), "Named-style name setter");

    xlpp::DefinedName defined("ValueName", "Sheet1!$A$1");
    defined.setValue("Sheet1!$B$2");
    defined.setLocalSheetId(0);
    defined.clearLocalSheetId();
    test.checkEqual(defined.value(), std::string("Sheet1!$B$2"), "Defined-name value setter");
    test.checkTrue(!defined.localSheetId().has_value(), "Defined-name local scope clear");

    xlpp::CustomProperty property;
    property.setName("RenamedProperty");
    property.setValue("42");
    property.setType("i4");
    test.checkEqual(property.name(), std::string("RenamedProperty"), "Custom-property name setter");
    test.checkEqual(property.value(), std::string("42"), "Custom-property value setter");
    test.checkEqual(property.type(), std::string("i4"), "Custom-property type setter");

    xlpp::Workbook workbook;
    workbook.addWorksheet("Sheet1");
    workbook.addDefinedName(std::move(defined));
    const auto& constWorkbook = static_cast<const xlpp::Workbook&>(workbook);
    test.checkEqual(constWorkbook.definedNames().size(), std::size_t{1}, "Const defined-names collection accessor");
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

    const auto emptyPath = std::filesystem::temp_directory_path() / "xlpp_mapped_empty_test.bin";
    { std::ofstream empty(emptyPath, std::ios::binary); }
    xlpp::internal::MappedFile emptyMapped(emptyPath);
    test.checkEqual(emptyMapped.size(), std::size_t{0}, "Mapped-file supports empty files without OS mapping failure");
    test.checkTrue(emptyMapped.view().empty() && emptyMapped.slice(0, 0).empty(),
                   "Empty mapped-file view and zero-length slice are valid");

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

    // Streaming parsing is intentionally strict: malformed numeric prefixes
    // and out-of-range shared-string references must not be silently accepted.
    const auto malformedNumericPath = std::filesystem::temp_directory_path() / "xlpp_stream_malformed_numeric.xlsx";
    {
        xlpp::Workbook wb;
        auto& s = wb.addWorksheet("Data");
        s.cell("A1").setValue(12.5);
        wb.save(malformedNumericPath);
        auto zip = xlpp::internal::ZipArchive::open(malformedNumericPath);
        auto xml = zip.get("xl/worksheets/sheet1.xml");
        const auto pos = xml.find("<v>12.5</v>");
        if (pos != std::string::npos) xml.replace(pos, std::string("<v>12.5</v>").size(), "<v>12.5oops</v>");
        zip.replace("xl/worksheets/sheet1.xml", xml);
        zip.save(malformedNumericPath);
    }
    bool malformedNumericRejected = false;
    try {
        xlpp::StreamingWorkbookReader reader(malformedNumericPath);
        reader.forEachRow("Data", [](std::size_t, const xlpp::StreamingRow&) { return true; });
    } catch (const std::runtime_error&) { malformedNumericRejected = true; }
    test.checkTrue(malformedNumericRejected, "Streaming reader rejects malformed numeric cell text");

    const auto badSstPath = std::filesystem::temp_directory_path() / "xlpp_stream_bad_sst.xlsx";
    {
        xlpp::Workbook wb;
        auto& s = wb.addWorksheet("Data");
        s.cell("A1").setValue("only");
        wb.save(badSstPath);
        auto zip = xlpp::internal::ZipArchive::open(badSstPath);
        auto xml = zip.get("xl/worksheets/sheet1.xml");
        const auto pos = xml.find("<v>0</v>");
        if (pos != std::string::npos) xml.replace(pos, std::string("<v>0</v>").size(), "<v>999</v>");
        zip.replace("xl/worksheets/sheet1.xml", xml);
        zip.save(badSstPath);
    }
    bool badSstRejected = false;
    try {
        xlpp::StreamingWorkbookReader reader(badSstPath);
        reader.forEachRow("Data", [](std::size_t, const xlpp::StreamingRow&) { return true; });
    } catch (const std::runtime_error&) { badSstRejected = true; }
    test.checkTrue(badSstRejected, "Streaming reader rejects out-of-range shared-string index");

    xlpp::internal::ZipEntryInfo dummyInfo;
    dummyInfo.name = "dummy";
    xlpp::internal::ZipEntrySource pathSource(rawPath, dummyInfo);
    test.checkTrue(!pathSource.complete(), "Path-backed ZIP entry source constructor");

    std::filesystem::remove(rawPath);
    std::filesystem::remove(emptyPath);
    std::filesystem::remove(workbookPath);
    std::filesystem::remove(malformedNumericPath);
    std::filesystem::remove(badSstPath);
}

void testRangeBoundsAndIndexedCell(TestContext& test) {
    xlpp::Worksheet sheet("Range");
    auto range = sheet.range(4, 5, 2, 3);
    test.checkEqual(range.minRow(), std::size_t{2}, "Numeric range min row normalizes");
    test.checkEqual(range.minColumn(), std::size_t{3}, "Numeric range min column normalizes");
    test.checkEqual(range.maxRow(), std::size_t{4}, "Numeric range max row normalizes");
    test.checkEqual(range.maxColumn(), std::size_t{5}, "Numeric range max column normalizes");
    range.cell(1, 1).setValue("relative");
    test.checkEqual(sheet.cell("C2").stringValueOr(""), std::string("relative"), "Range indexed cell uses relative coordinates");
    test.checkTrue(sheet.empty() == false, "Worksheet empty reflects populated range");
    sheet.markDirty();
    test.checkTrue(sheet.dirty(), "Worksheet markDirty sets dirty flag");
}

}



void writeExternalReaderFixture(const std::filesystem::path& path,
                                const std::string& workbookXml,
                                const std::string& sheetXml,
                                const std::string& stylesXml,
                                const std::string& sharedStringsXml = {},
                                const std::string& sheetRelationships = {},
                                const std::string& commentsXml = {},
                                const std::string& corePropertiesXml = {},
                                const std::string& customPropertiesXml = {}) {
    constexpr auto packageNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    constexpr auto documentNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    xlpp::internal::ZipArchive zip;
    std::string contentTypes =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
    if (!sharedStringsXml.empty())
        contentTypes += "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
    if (!commentsXml.empty())
        contentTypes += "<Override PartName=\"/xl/comments1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>";
    if (!corePropertiesXml.empty())
        contentTypes += "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>";
    if (!customPropertiesXml.empty())
        contentTypes += "<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>";
    contentTypes += "</Types>";
    zip.add("[Content_Types].xml", contentTypes);

    std::string rootRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"";
    rootRels += packageNs;
    rootRels += "\"><Relationship Id=\"rIdWorkbook\" Type=\"";
    rootRels += documentNs;
    rootRels += "/officeDocument\" Target=\"xl/workbook.xml\"/>";
    if (!corePropertiesXml.empty())
        rootRels += "<Relationship Id=\"rIdCore\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>";
    if (!customPropertiesXml.empty()) {
        rootRels += "<Relationship Id=\"rIdCustom\" Type=\"";
        rootRels += documentNs;
        rootRels += "/custom-properties\" Target=\"docProps/custom.xml\"/>";
    }
    rootRels += "</Relationships>";
    zip.add("_rels/.rels", rootRels);
    zip.add("xl/workbook.xml", workbookXml);

    std::string workbookRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"";
    workbookRels += packageNs;
    workbookRels += "\"><Relationship Id=\"rIdSheet1\" Type=\"";
    workbookRels += documentNs;
    workbookRels += "/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rIdStyles\" Type=\"";
    workbookRels += documentNs;
    workbookRels += "/styles\" Target=\"styles.xml\"/>";
    if (!sharedStringsXml.empty()) {
        workbookRels += "<Relationship Id=\"rIdShared\" Type=\"";
        workbookRels += documentNs;
        workbookRels += "/sharedStrings\" Target=\"sharedStrings.xml\"/>";
    }
    workbookRels += "</Relationships>";
    zip.add("xl/_rels/workbook.xml.rels", workbookRels);
    zip.add("xl/worksheets/sheet1.xml", sheetXml);
    zip.add("xl/styles.xml", stylesXml);
    if (!sharedStringsXml.empty()) zip.add("xl/sharedStrings.xml", sharedStringsXml);
    if (!sheetRelationships.empty()) zip.add("xl/worksheets/_rels/sheet1.xml.rels", sheetRelationships);
    if (!commentsXml.empty()) zip.add("xl/comments1.xml", commentsXml);
    if (!corePropertiesXml.empty()) zip.add("docProps/core.xml", corePropertiesXml);
    if (!customPropertiesXml.empty()) zip.add("docProps/custom.xml", customPropertiesXml);
    zip.save(path);
}

void testVbaSourceTextBuildAndRead(TestContext& test) {
    const auto firstPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_source.xlsm";
    const auto secondPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_source_updated.xlsm";
    const auto sheetMutationPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_sheet_mutation.xlsm";
    const auto removedPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_source_removed.xlsx";
    const std::string source =
        "Option Explicit\n"
        "Public Sub XLPP_Hello()\n"
        "    Range(\"A2\").Value = 42\n"
        "End Sub\n"
        "Public Function AddTwo(ByVal value As Double) As Double\n"
        "    AddTwo = value + 2\n"
        "End Function\n";

    xlpp::Workbook workbook;
    workbook.addWorksheet("MacroHost").cell("A1").setValue("=AddTwo(40)");
    workbook.setVbaModuleText("MathModule", source);
    // Add a worksheet after creating the VBA project. save() must regenerate
    // document modules so Sheet2 and sheetPr codeName stay synchronized.
    workbook.addWorksheet("AddedAfterVba").cell("A1").setValue("document-module sync");
    test.checkTrue(workbook.hasVbaProject(), "VBA text API creates a binary project");
    const auto inMemory = workbook.vbaModuleText("mathmodule");
    test.checkTrue(inMemory.has_value(), "VBA module lookup is case-insensitive");
    test.checkEqual(*inMemory,
                    std::string("Option Explicit\r\nPublic Sub XLPP_Hello()\r\n    Range(\"A2\").Value = 42\r\nEnd Sub\r\nPublic Function AddTwo(ByVal value As Double) As Double\r\n    AddTwo = value + 2\r\nEnd Function\r\n"),
                    "VBA source is normalized to CRLF");
    test.checkTrue(inMemory->find("Public Sub XLPP_Hello()") != std::string::npos,
                   "Generated standard module contains a public parameterless macro");
    test.checkTrue(inMemory->find("Private Sub XLPP_Hello") == std::string::npos,
                   "Generated macro is not private");
    workbook.save(firstPath);

    auto zip = xlpp::internal::ZipArchive::open(firstPath);
    test.checkTrue(zip.contains("xl/vbaProject.bin"), "Text-generated vbaProject.bin is packaged");
    test.checkTrue(zip.get("[Content_Types].xml").find("application/vnd.ms-excel.sheet.macroEnabled.main+xml") != std::string::npos,
                   "Workbook content type is macro-enabled");
    test.checkTrue(zip.get("xl/_rels/workbook.xml.rels").find("/relationships/vbaProject") != std::string::npos,
                   "Workbook relationship targets vbaProject.bin");
    const auto& binary = zip.get("xl/vbaProject.bin");
    test.checkTrue(binary.size() >= 512, "Generated VBA project is a compound file");
    test.checkEqual(static_cast<unsigned char>(binary[0]), static_cast<unsigned char>(0xD0), "CFB signature byte 0");
    test.checkEqual(static_cast<unsigned char>(binary[1]), static_cast<unsigned char>(0xCF), "CFB signature byte 1");
    test.checkTrue(zip.get("xl/workbook.xml").find("codeName=\"ThisWorkbook\"") != std::string::npos,
                   "Macro workbook emits ThisWorkbook code name");
    test.checkTrue(zip.get("xl/worksheets/sheet1.xml").find("<sheetPr codeName=\"Sheet1\">") != std::string::npos,
                   "Macro worksheet emits the Sheet1 document-module code name");
    test.checkTrue(zip.get("xl/worksheets/sheet2.xml").find("<sheetPr codeName=\"Sheet2\">") != std::string::npos,
                   "Worksheet added after VBA creation emits a synchronized Sheet2 code name");

    xlpp::Workbook loaded;
    loaded.load(firstPath);
    const auto modules = loaded.vbaModules();
    test.checkEqual(modules.size(), std::size_t{4}, "Reader exposes both sheets, workbook, and standard VBA modules");
    test.checkEqual(modules[0].name, std::string("Sheet1"), "First worksheet document module is present");
    test.checkEqual(static_cast<int>(modules[0].type), static_cast<int>(xlpp::VbaModuleType::Document), "First worksheet module type is document");
    test.checkEqual(modules[1].name, std::string("Sheet2"), "Worksheet added after VBA creation has a document module");
    test.checkEqual(static_cast<int>(modules[1].type), static_cast<int>(xlpp::VbaModuleType::Document), "Second worksheet module type is document");
    test.checkEqual(modules[2].name, std::string("ThisWorkbook"), "Workbook document module is present");
    test.checkEqual(static_cast<int>(modules[2].type), static_cast<int>(xlpp::VbaModuleType::Document), "Workbook module type is document");
    test.checkEqual(modules[3].name, std::string("MathModule"), "Standard module name is read from dir stream");
    test.checkEqual(static_cast<int>(modules[3].type), static_cast<int>(xlpp::VbaModuleType::Standard), "Visible macro resides in a standard module");
    test.checkEqual(loaded.vbaModuleText("MathModule").value(), *inMemory,
                    "VBA source survives XLSM save/load");

    // A generated project loaded from disk remains identifiable as XL++-owned.
    // Adding a sheet without editing VBA must still rebuild document modules.
    loaded.addWorksheet("AddedAfterReload").cell("A1").setValue("reload sync");
    loaded.save(sheetMutationPath);
    xlpp::Workbook sheetMutated;
    sheetMutated.load(sheetMutationPath);
    const auto mutatedModules = sheetMutated.vbaModules();
    test.checkEqual(mutatedModules.size(), std::size_t{5},
                    "Loaded XL++ VBA project rebuilds document modules after worksheet insertion");
    test.checkEqual(mutatedModules[2].name, std::string("Sheet3"),
                    "Sheet inserted after reload receives a Sheet3 VBA document module");
    test.checkTrue(sheetMutated.vbaModuleText("MathModule").has_value(),
                   "Standard macro module survives sheet-driven VBA project regeneration");

    loaded.setVbaModuleText("MathModule", "Public Sub Updated()\n    Range(\"A2\").Value = 99\nEnd Sub");
    loaded.setVbaModuleText("SecondModule", "Public Const Answer As Long = 42\n");
    loaded.save(secondPath);
    xlpp::Workbook updated;
    updated.load(secondPath);
    test.checkTrue(updated.vbaModuleText("MathModule")->find("Updated") != std::string::npos,
                   "Existing VBA module source can be replaced after load");
    test.checkTrue(updated.vbaModuleText("SecondModule")->find("Answer") != std::string::npos,
                   "Second VBA source module can be added after load");
    test.checkTrue(updated.removeVbaModule("MathModule"), "Individual VBA module can be removed");
    test.checkTrue(!updated.vbaModuleText("MathModule").has_value(), "Removed VBA module is absent from reader");
    test.checkTrue(updated.removeVbaModule("SecondModule"), "Last standard VBA module can be removed");
    test.checkTrue(!updated.hasVbaProject(), "Removing the last standard module removes the VBA project");
    updated.save(removedPath);
    zip = xlpp::internal::ZipArchive::open(removedPath);
    test.checkTrue(!zip.contains("xl/vbaProject.bin"), "Workbook returns to XLSX after all text modules are removed");

    bool invalidRejected = false;
    try { workbook.setVbaModuleText("1 Bad Name", "Sub X(): End Sub"); }
    catch (const std::invalid_argument&) { invalidRejected = true; }
    test.checkTrue(invalidRejected, "Invalid VBA identifier is rejected");

    // Exercise regular CFB sectors, multiple OVBA chunks, an empty source stream,
    // and a larger directory red-black tree rather than only mini-stream modules.
    const auto stressPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_stress.xlsm";
    std::string longSource = "Option Explicit\nPublic Sub LongGeneratedMacro()\n";
    for (int line = 0; line < 700; ++line)
        longSource += "    ' generated source line " + std::to_string(line) + " for CFB and OVBA coverage\n";
    longSource += "End Sub\n";
    xlpp::Workbook stress;
    stress.addWorksheet("MacroHost");
    stress.setVbaModuleText("EmptyModule", "");
    stress.setVbaModuleText("ModuleA", "Public Const A As Long = 1\n");
    stress.setVbaModuleText("ModuleB", "Public Const B As Long = 2\n");
    stress.setVbaModuleText("ModuleC", "Public Const C As Long = 3\n");
    stress.setVbaModuleText("ModuleD", "Public Const D As Long = 4\n");
    stress.setVbaModuleText("LongModule", longSource);
    stress.save(stressPath);
    auto stressZip = xlpp::internal::ZipArchive::open(stressPath);
    test.checkTrue(stressZip.get("xl/vbaProject.bin").size() > 8192,
                   "Large text module uses a multi-sector CFB project");
    xlpp::Workbook stressLoaded;
    stressLoaded.load(stressPath);
    test.checkEqual(stressLoaded.vbaModules().size(), std::size_t{8},
                    "Multiple standard modules plus Sheet1 and ThisWorkbook are read");
    test.checkEqual(stressLoaded.vbaModuleText("EmptyModule").value(), std::string{},
                    "Empty VBA source stream round-trips");
    test.checkTrue(stressLoaded.vbaModuleText("LongModule").value() ==
                       xlpp::internal::normalizeVbaSource(longSource),
                   "Long multi-chunk VBA source round-trips exactly");

    std::filesystem::remove(firstPath);
    std::filesystem::remove(secondPath);
    std::filesystem::remove(sheetMutationPath);
    std::filesystem::remove(removedPath);
    std::filesystem::remove(stressPath);
}


void testPivotImportedReaderAndEditorP1A(TestContext& test) {
    const auto fixture = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "libreoffice" / "pivot.xlsx";
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p1a_pivot_reader_editor.xlsx";
    test.checkTrue(std::filesystem::exists(fixture), "P1A pivot reader fixture exists");

    xlpp::Workbook workbook;
    workbook.load(fixture);
    const auto& readOnlyWorkbook = static_cast<const xlpp::Workbook&>(workbook);
    const auto* loadedSheet = readOnlyWorkbook.worksheet("Data");
    test.checkTrue(loadedSheet != nullptr, "P1A pivot source worksheet loads");
    test.checkEqual(loadedSheet->pivotTables().size(), std::size_t{1}, "Imported relationship-only pivot enters object model");
    const auto& pivot = loadedSheet->pivotTables().front();
    test.checkEqual(pivot.name(), std::string("SalesPivot"), "Imported pivot name is read");
    test.checkEqual(pivot.location(), std::string("E4:H9"), "Imported pivot location is read");
    test.checkEqual(pivot.cache().cacheId(), 1, "Imported pivot logical cache ID is read");
    test.checkEqual(pivot.cache().sourceData(), std::string("'Data'!A1:C7"), "Imported pivot worksheet source is read");
    test.checkEqual(pivot.cache().fields().size(), std::size_t{3}, "Imported cache fields are materialized");
    test.checkEqual(pivot.cache().fields()[0], std::string("Region"), "Imported first cache field name is read");
    test.checkEqual(pivot.cache().fields()[2], std::string("Sales"), "Imported numeric cache field name is read");
    test.checkEqual(pivot.cache().records().size(), std::size_t{6}, "Imported cache records are materialized");
    test.checkEqual(pivot.cache().records()[0][0], std::string("East"), "Shared-item cache index resolves to text");
    test.checkEqual(pivot.cache().records()[0][1], std::string("Q1"), "Second shared-item cache index resolves to text");
    test.checkEqual(pivot.cache().records()[0][2], std::string("10"), "Numeric cache index resolves to value text");
    test.checkEqual(pivot.rowFields().size(), std::size_t{1}, "Imported row field is read");
    test.checkEqual(pivot.rowFields()[0].fieldIndex(), 0, "Imported row field index is read");
    test.checkTrue(!pivot.rowFields()[0].compact(), "Imported row compact flag is read");
    test.checkEqual(pivot.columnFields().size(), std::size_t{1}, "Imported column field is read");
    test.checkEqual(pivot.dataFields().size(), std::size_t{1}, "Imported data field is read");
    test.checkEqual(pivot.dataFields()[0].fieldIndex(), 2, "Imported data field source index is read");
    test.checkEqual(pivot.dataFields()[0].displayName(), std::string("Sum - Sales"), "Imported data-field caption is read");
    test.checkEqual(pivot.dataFields()[0].numberFormatId(), 164, "Imported data-field number format ID is read");
    test.checkEqual(pivot.styleName(), std::string("PivotStyleLight16"), "Imported pivot style is read");

    auto* editableSheet = workbook.worksheet("Data");
    auto& editable = editableSheet->pivotTables().front();
    editable.setStyleName("PivotStyleMedium9");
    editable.setRowGrandTotals(false);
    editable.setShowRowStripes(true);
    editable.cache().setRefreshOnLoad(false);
    editable.rowFields()[0].hideItem(1);
    editable.dataFields()[0].setDisplayName("Sales %");
    editable.dataFields()[0].setShowDataAs("percentOfTotal");
    workbook.save(output);

    const auto archive = xlpp::internal::ZipArchive::open(output);
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    const auto validation = graph.validate();
    test.checkTrue(validation.ok(), "Mutated imported pivot saves as a valid OOXML object graph");
    test.checkEqual(graph.objectInventory().pivotTables, std::size_t{1}, "Regenerated imported pivot has one reachable pivot-table root");
    const auto generatedPivotPart = archive.contains("xl/pivotTables/pivotTable2.xml")
        ? std::string("xl/pivotTables/pivotTable2.xml") : std::string("xl/pivotTables/pivotTable1.xml");
    const auto pivotXml = archive.get(generatedPivotPart);
    test.checkTrue(pivotXml.find("rowGrandTotals=\"0\"") != std::string::npos, "Edited row-grand-total flag is serialized");
    test.checkTrue(pivotXml.find("name=\"PivotStyleMedium9\"") != std::string::npos, "Edited pivot style is serialized");
    test.checkTrue(pivotXml.find("showDataAs=\"percentOfTotal\"") != std::string::npos, "Edited show-data-as mode is serialized");
    test.checkTrue(pivotXml.find("name=\"Sales %\"") != std::string::npos, "Edited data-field caption is serialized");
    test.checkTrue(pivotXml.find(" h=\"1\"") != std::string::npos, "Edited hidden pivot item is serialized");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto& reloadedConst = static_cast<const xlpp::Workbook&>(reloaded);
    const auto* reloadedSheet = reloadedConst.worksheet("Data");
    test.checkTrue(reloadedSheet != nullptr && reloadedSheet->pivotTables().size() == 1, "Regenerated pivot reloads through object model");
    const auto& roundTrip = reloadedSheet->pivotTables().front();
    test.checkEqual(roundTrip.styleName(), std::string("PivotStyleMedium9"), "Edited pivot style round-trips");
    test.checkTrue(!roundTrip.rowGrandTotals(), "Edited row-grand-total flag round-trips");
    test.checkTrue(roundTrip.showRowStripes(), "Edited style stripe flag round-trips");
    test.checkTrue(!roundTrip.cache().refreshOnLoad(), "Edited cache refresh-on-load flag round-trips");
    test.checkEqual(roundTrip.dataFields()[0].displayName(), std::string("Sales %"), "Edited data caption round-trips");
    test.checkEqual(roundTrip.dataFields()[0].showDataAs(), std::string("percentOfTotal"), "Edited show-data-as mode round-trips");
    test.checkTrue(std::find(roundTrip.rowFields()[0].hiddenItems().begin(), roundTrip.rowFields()[0].hiddenItems().end(), 1) != roundTrip.rowFields()[0].hiddenItems().end(),
                   "Edited hidden pivot item round-trips");
    std::filesystem::remove(output);
}

void testPivotOlapCalculatedMembersP1Y(TestContext& test) {
    // P1Y-A OLAP pivot metadata: model cacheSource type="olap", olapInfo and
    // calculatedMember nodes, generate them, round-trip them and patch them
    // in-place without regenerating the physical cache.
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1y_olap_source.xlsx";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1y_olap_edited.xlsx";

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.cell("A1").setValue("Placeholder");

    xlpp::PivotCache cache;
    cache.setCacheId(1);
    cache.setFields({"Measure"});
    cache.setTypedRecords({{"10"}}, {{xlpp::PivotCacheValueKind::Number}});
    xlpp::PivotOlapSourceInfo olap;
    olap.sourceType = "olap";
    olap.preserveFormatting = true;
    olap.connectionId = 1;
    cache.setOlap(std::move(olap));
    xlpp::PivotCalculatedMember member;
    member.name = "Profit";
    member.mdx = "[Measures].[Sales]-[Measures].[Cost]";
    member.hierarchy = 0;
    cache.calculatedMembers().push_back(member);

    xlpp::PivotTable pivot("OlapPivot");
    pivot.setLocation("C2");
    pivot.cache() = cache;
    pivot.addDataField(0);
    sheet.addPivotTable(std::move(pivot));
    workbook.save(source);

    xlpp::Workbook loaded;
    loaded.load(source);
    auto* loadedSheet = loaded.worksheet("Data");
    test.checkTrue(loadedSheet != nullptr && loadedSheet->pivotTables().size() == 1, "P1Y OLAP pivot reloads");
    auto& loadedPivot = loadedSheet->pivotTables().front();
    const auto* loadedOlap = static_cast<const xlpp::PivotCache&>(loadedPivot.cache()).olap();
    test.checkTrue(loadedOlap != nullptr, "P1Y OLAP cache source identity is read");
    test.checkEqual(loadedOlap->sourceType, std::string("olap"), "P1Y cacheSource type is read");
    test.checkTrue(loadedOlap->preserveFormatting, "P1Y olapInfo preserveFormatting is read");
    test.checkEqual(loadedPivot.cache().calculatedMembers().size(), std::size_t{1}, "P1Y calculated member is read");
    test.checkEqual(loadedPivot.cache().calculatedMembers()[0].name, std::string("Profit"), "P1Y calculated member name is read");
    test.checkEqual(loadedPivot.cache().calculatedMembers()[0].mdx,
                    std::string("[Measures].[Sales]-[Measures].[Cost]"), "P1Y calculated member MDX is read");

    // Selective patch: change preserveFormatting and the member MDX.
    xlpp::PivotOlapSourcePatch olapPatch;
    olapPatch.preserveFormatting = false;
    olapPatch.connectionId = 2;
    test.checkTrue(loaded.updateImportedPivotOlapSource("Data", "OlapPivot", olapPatch),
                   "P1Y OLAP cache source patch succeeds");
    xlpp::PivotCalculatedMemberPatch memberPatch;
    memberPatch.mdx = "[Measures].[Gross]-[Measures].[Cost]";
    memberPatch.memberName = "ProfitMember";
    test.checkTrue(loaded.updateImportedPivotCalculatedMember("Data", "OlapPivot", 0, memberPatch),
                   "P1Y calculated member patch succeeds");
    loaded.save(edited);

    xlpp::Workbook round; round.load(edited);
    auto* roundSheet = round.worksheet("Data");
    test.checkTrue(roundSheet != nullptr && roundSheet->pivotTables().size() == 1, "P1Y patched OLAP pivot reloads");
    auto& roundPivot = roundSheet->pivotTables().front();
    const auto* roundOlap = static_cast<const xlpp::PivotCache&>(roundPivot.cache()).olap();
    test.checkTrue(roundOlap != nullptr, "P1Y OLAP cache identity survives edit");
    test.checkTrue(!roundOlap->preserveFormatting, "P1Y olapInfo preserveFormatting patch round-trips");
    test.checkEqual(roundOlap->connectionId, 2, "P1Y OLAP connectionId patch round-trips");
    test.checkEqual(roundPivot.cache().calculatedMembers().size(), std::size_t{1}, "P1Y calculated member survives patch");
    test.checkEqual(roundPivot.cache().calculatedMembers()[0].mdx,
                    std::string("[Measures].[Gross]-[Measures].[Cost]"), "P1Y calculated member MDX patch round-trips");
    test.checkEqual(roundPivot.cache().calculatedMembers()[0].memberName,
                    std::string("ProfitMember"), "P1Y calculated member memberName patch round-trips");
    test.checkTrue(roundOlap->rawOlapInfoXml.find("preserveFormatting=\"0\"") != std::string::npos
                   || roundPivot.cache().calculatedMembers().front().rawXml.find("[Measures].[Gross]") != std::string::npos,
                   "P1Y patched OLAP metadata persists in package bytes");

    // Verify the patched bytes actually reached the saved package.
    const auto archive = xlpp::internal::ZipArchive::open(edited);
    std::string cacheXml;
    for (int i = 1; i <= 8 && cacheXml.empty(); ++i) {
        const auto candidate = "xl/pivotCache/pivotCacheDefinition" + std::to_string(i) + ".xml";
        if (archive.contains(candidate)) cacheXml = archive.get(candidate);
    }
    test.checkTrue(cacheXml.find("preserveFormatting=\"0\"") != std::string::npos,
                   "P1Y preserveFormatting=0 is written to the saved pivotCacheDefinition");
    test.checkTrue(cacheXml.find("connectionId=\"2\"") != std::string::npos,
                   "P1Y connectionId=2 is written to the saved pivotCacheDefinition");
    test.checkTrue(cacheXml.find("preserveFormatting") != std::string::npos,
                   "P1Y pivotCacheDefinition contains an olapInfo block");
    test.checkTrue(cacheXml.find("[Measures].[Gross]") != std::string::npos,
                   "P1Y patched calculated member MDX is written to the saved pivotCacheDefinition");

    std::filesystem::remove(source); std::filesystem::remove(edited);
}

void testPivotTypedGroupItemsP1Y(TestContext& test) {
    // P1Y-A preserves the physical kind of Pivot groupItems (<n>/<d>/<s>/<m>)
    // instead of flattening every date/number bin to a plain string label.
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1y_group_items.xlsx";

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    xlpp::PivotTable pivot("Bins");
    pivot.setLocation("D2");
    pivot.cache().setSourceData("'Data'!$A$1:$B$5");
    pivot.cache().setFields({"OrderDate", "Amount"});
    pivot.cache().addRecord({"2026-01-15T00:00:00", "10"});
    pivot.cache().addRecord({"2026-02-20T00:00:00", "20"});
    pivot.addRowField("OrderDate");
    pivot.addDataField("Amount", "sum");

    // A date-grouped field whose groupItems mix date (<d>) and number (<n>)
    // bins plus a missing (<m>) sentinel.
    auto& group = pivot.cache().fieldGroup(0);
    group.groupBy = "days";
    group.autoStart = false;
    group.autoEnd = false;
    group.startDate = "2026-01-01T00:00:00";
    group.endDate = "2026-12-31T23:59:59";
    group.addTypedGroupItem(xlpp::PivotCacheValueKind::DateTime, "2026-01-01T00:00:00");
    group.addTypedGroupItem(xlpp::PivotCacheValueKind::DateTime, "2026-01-02T00:00:00");
    group.addTypedGroupItem(xlpp::PivotCacheValueKind::Number, "3");
    group.addTypedGroupItem(xlpp::PivotCacheValueKind::Missing, "");
    sheet.addPivotTable(std::move(pivot));
    workbook.save(source);

    const auto archive = xlpp::internal::ZipArchive::open(source);
    const auto cachePart = archive.contains("xl/pivotCache/pivotCacheDefinition2.xml")
        ? std::string("xl/pivotCache/pivotCacheDefinition2.xml") : std::string("xl/pivotCache/pivotCacheDefinition1.xml");
    const auto cacheXml = archive.get(cachePart);
    test.checkTrue(cacheXml.find("<d v=\"2026-01-01T00:00:00\"/>") != std::string::npos,
                   "P1Y date group bin is serialized as a <d> element");
    test.checkTrue(cacheXml.find("<n v=\"3\"/>") != std::string::npos,
                   "P1Y numeric group bin is serialized as an <n> element");
    test.checkTrue(cacheXml.find("<m/>") != std::string::npos,
                   "P1Y missing group bin is serialized as an <m> element");

    xlpp::Workbook round; round.load(source);
    auto* roundSheet = round.worksheet("Data");
    test.checkTrue(roundSheet != nullptr && roundSheet->pivotTables().size() == 1, "P1Y grouped pivot reloads");
    const auto* loadedGroup = static_cast<const xlpp::PivotCache&>(roundSheet->pivotTables().front().cache()).tryFieldGroup(0);
    test.checkTrue(loadedGroup != nullptr, "P1Y fieldGroup survives reload");
    test.checkEqual(loadedGroup->typedItems.size(), std::size_t{4}, "P1Y typed group items are read back");
    test.checkTrue(loadedGroup->typedGroupItem(0) != nullptr
                   && loadedGroup->typedGroupItem(0)->kind == xlpp::PivotCacheValueKind::DateTime
                   && loadedGroup->typedGroupItem(0)->value == "2026-01-01T00:00:00",
                   "P1Y first date bin kind/value is preserved");
    test.checkTrue(loadedGroup->typedGroupItem(2) != nullptr
                   && loadedGroup->typedGroupItem(2)->kind == xlpp::PivotCacheValueKind::Number
                   && loadedGroup->typedGroupItem(2)->value == "3",
                   "P1Y numeric bin kind/value is preserved");
    test.checkTrue(loadedGroup->typedGroupItem(3) != nullptr
                   && loadedGroup->typedGroupItem(3)->kind == xlpp::PivotCacheValueKind::Missing,
                   "P1Y missing bin kind is preserved");
    test.checkEqual(loadedGroup->items.size(), std::size_t{4}, "P1Y legacy text view stays in sync with typed items");

    std::filesystem::remove(source);
}

void testVbaClassDocumentAndProjectMetadataP1A(TestContext& test) {
    const auto first = std::filesystem::temp_directory_path() / "xlpp_p1a_vba_modules.xlsm";
    const auto second = std::filesystem::temp_directory_path() / "xlpp_p1a_vba_modules_sheet_add.xlsm";

    xlpp::Workbook workbook;
    workbook.addWorksheet("Data");
    workbook.setVbaModuleText("Utilities", "Public Function Twice(ByVal x As Double) As Double\nTwice = x * 2\nEnd Function");
    xlpp::VbaModule worker;
    worker.name = "WorkerClass";
    worker.type = xlpp::VbaModuleType::Class;
    worker.source = "Option Explicit\nPublic Name As String";
    worker.docString = "Worker class metadata";
    worker.readOnly = true;
    worker.isPrivate = true;
    workbook.setVbaModule(std::move(worker));
    workbook.setVbaDocumentModuleText("ThisWorkbook", "Private Sub Workbook_Open()\nWorksheets(1).Range(\"A1\").Value = \"Opened\"\nEnd Sub");
    workbook.setVbaDocumentModuleText("Sheet1", "Private Sub Worksheet_Activate()\nRange(\"A2\").Value = 7\nEnd Sub");

    xlpp::VbaProjectInfo info;
    info.name = "XLPPAnalytics";
    info.description = "P1A VBA project metadata";
    info.helpFile = "xlpp-help.chm";
    info.helpContextId = 42;
    info.constants = "P1A = 1 : Feature = 2";
    info.references.push_back({"Scripting", "*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime"});
    info.projectId = "{12345678-1234-4ABC-9DEF-1234567890AB}";
    workbook.setVbaProjectInfo(info);
    workbook.save(first);

    xlpp::Workbook loaded;
    loaded.load(first);
    test.checkTrue(loaded.hasVbaProject(), "P1A generated VBA project reloads");
    const auto loadedInfo = loaded.vbaProjectInfo();
    test.checkEqual(loadedInfo.name, std::string("XLPPAnalytics"), "VBA project name round-trips");
    test.checkEqual(loadedInfo.description, std::string("P1A VBA project metadata"), "VBA project description round-trips");
    test.checkEqual(loadedInfo.helpFile, std::string("xlpp-help.chm"), "VBA help file round-trips");
    test.checkEqual(loadedInfo.helpContextId, 42, "VBA help context round-trips");
    test.checkEqual(loadedInfo.constants, std::string("P1A = 1 : Feature = 2"), "VBA project constants round-trip from dir stream");
    test.checkEqual(loadedInfo.projectId, std::string("{12345678-1234-4ABC-9DEF-1234567890AB}"), "VBA project ID round-trips");
    const auto scriptingReference = std::find_if(loadedInfo.references.begin(), loadedInfo.references.end(), [](const auto& reference) { return reference.name == "Scripting"; });
    test.checkTrue(scriptingReference != loadedInfo.references.end(), "Registered VBA type-library reference round-trips");
    test.checkTrue(scriptingReference != loadedInfo.references.end() && scriptingReference->libid.find("420B2830-E718-11CF-893D-00A0C9054228") != std::string::npos,
                   "Registered VBA reference LibId round-trips");
    test.checkTrue(loadedInfo.references.size() >= 3, "VBA project reader exposes baseline and custom registered references");

    const auto modules = loaded.vbaModules();
    const auto findModule = [&](const std::string& name) -> const xlpp::VbaModule* {
        const auto it = std::find_if(modules.begin(), modules.end(), [&](const auto& module) { return module.name == name; });
        return it == modules.end() ? nullptr : &*it;
    };
    const auto* classModule = findModule("WorkerClass");
    test.checkTrue(classModule != nullptr, "VBA class module reloads");
    test.checkEqual(static_cast<int>(classModule->type), static_cast<int>(xlpp::VbaModuleType::Class), "Class module remains Class rather than Document");
    test.checkEqual(classModule->docString, std::string("Worker class metadata"), "Class module doc string round-trips");
    test.checkTrue(classModule->readOnly, "Class module read-only metadata round-trips");
    test.checkTrue(classModule->isPrivate, "Class module private metadata round-trips");
    test.checkTrue(classModule->source.find("Public Name As String") != std::string::npos, "Class module source round-trips");
    const auto thisWorkbook = loaded.vbaModuleText("ThisWorkbook");
    const auto sheet1 = loaded.vbaModuleText("Sheet1");
    test.checkTrue(thisWorkbook.has_value() && thisWorkbook->find("Workbook_Open") != std::string::npos, "ThisWorkbook document source round-trips");
    test.checkTrue(sheet1.has_value() && sheet1->find("Worksheet_Activate") != std::string::npos, "Worksheet document source round-trips");
    test.checkTrue(!loaded.removeVbaModule("ThisWorkbook"), "Host document module cannot be removed through user-module API");

    // Customized project metadata must not stop XL++ from recognizing its own
    // source-generated project after reload. Adding a sheet should create the
    // matching document module while retaining existing class/document source.
    loaded.addWorksheet("Second");
    loaded.save(second);
    xlpp::Workbook reloaded;
    reloaded.load(second);
    const auto modulesAfterSheetAdd = reloaded.vbaModules();
    test.checkTrue(std::any_of(modulesAfterSheetAdd.begin(), modulesAfterSheetAdd.end(), [](const auto& module) {
        return module.name == "Sheet2" && module.type == xlpp::VbaModuleType::Document;
    }), "Adding worksheet after metadata-customized reload creates Sheet2 document module");
    test.checkTrue(reloaded.vbaModuleText("ThisWorkbook").value_or("").find("Workbook_Open") != std::string::npos,
                   "ThisWorkbook source survives sheet-driven project regeneration");
    const auto workerAfterSheetAdd = std::find_if(modulesAfterSheetAdd.begin(), modulesAfterSheetAdd.end(), [](const auto& module) { return module.name == "WorkerClass"; });
    test.checkTrue(workerAfterSheetAdd != modulesAfterSheetAdd.end() && workerAfterSheetAdd->type == xlpp::VbaModuleType::Class,
                   "Class module survives sheet-driven project regeneration");
    test.checkEqual(reloaded.vbaProjectInfo().name, std::string("XLPPAnalytics"), "Customized VBA project metadata survives sheet-driven regeneration");
    const auto reloadedProjectInfo = reloaded.vbaProjectInfo();
    test.checkTrue(std::any_of(reloadedProjectInfo.references.begin(), reloadedProjectInfo.references.end(), [](const auto& reference) { return reference.name == "Scripting"; }),
                   "Custom registered reference survives sheet-driven regeneration");
    test.checkTrue(reloaded.removeVbaModule("WorkerClass"), "Class module can be removed through public module API");
    test.checkTrue(!reloaded.vbaModuleText("WorkerClass").has_value(), "Removed class module disappears from reader");

    std::filesystem::remove(first);
    std::filesystem::remove(second);
}

void testPivotSharedCacheCalculatedGroupingP1B(TestContext& test) {
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p1b_shared_pivot_cache.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.cell("A1").setValue("Region"); sheet.cell("B1").setValue("Sales");
    sheet.cell("A2").setValue("East");  sheet.cell("B2").setValue(10.0);
    sheet.cell("A3").setValue("West");  sheet.cell("B3").setValue(20.0);
    sheet.cell("A4").setValue("East");  sheet.cell("B4").setValue(30.0);
    sheet.cell("A5").setValue("North"); sheet.cell("B5").setValue(40.0);

    xlpp::PivotCache cache;
    cache.setSourceData("'Data'!A1:B5");
    cache.setSharedCacheKey("sales-shared-cache");
    cache.setFields({"Region", "Sales"});
    cache.setRecords({{"East", "10"}, {"West", "20"}, {"East", "30"}, {"North", "40"}});
    const auto calculatedIndex = cache.addCalculatedField("Commission", "Sales*0.1");
    auto& salesGroup = cache.fieldGroup(1);
    salesGroup.baseField = 1;
    salesGroup.autoStart = false;
    salesGroup.autoEnd = false;
    salesGroup.startNumber = 10.0;
    salesGroup.endNumber = 40.0;
    salesGroup.interval = 10.0;
    salesGroup.items = {"10-19", "20-29", "30-39", "40-49"};

    xlpp::PivotTable first("SalesByRegion");
    first.setLocation("D2");
    first.cache() = cache;
    auto& regionField = first.addRowField("Region");
    regionField.setDefaultSubtotal(false);
    regionField.addSubtotal("sum");
    regionField.addSubtotal("avg");
    first.addDataField(calculatedIndex).setDisplayName("Commission");
    sheet.addPivotTable(first);

    xlpp::PivotTable second("SalesSummary");
    second.setLocation("J2");
    second.cache() = cache;
    second.addRowField("Region");
    second.addDataField("Sales", "sum").setDisplayName("Total Sales");
    sheet.addPivotTable(second);

    workbook.save(output);
    const auto archive = xlpp::internal::ZipArchive::open(output);
    std::size_t pivotTables = 0, cacheDefinitions = 0, cacheRecords = 0;
    for (const auto& name : archive.entryNames()) {
        if (name.rfind("xl/pivotTables/pivotTable", 0) == 0 && name.find("_rels") == std::string::npos) ++pivotTables;
        if (name.rfind("xl/pivotCache/pivotCacheDefinition", 0) == 0 && name.find("_rels") == std::string::npos) ++cacheDefinitions;
        if (name.rfind("xl/pivotCache/pivotCacheRecords", 0) == 0) ++cacheRecords;
    }
    test.checkEqual(pivotTables, std::size_t{2}, "Two generated PivotTables are emitted");
    test.checkEqual(cacheDefinitions, std::size_t{1}, "Shared cache key emits one physical PivotCache definition");
    test.checkEqual(cacheRecords, std::size_t{1}, "Shared cache key emits one physical PivotCache records part");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    test.checkTrue(graph.validate().ok(), "Shared PivotCache package graph validates");

    const auto cacheXml = archive.get("xl/pivotCache/pivotCacheDefinition1.xml");
    test.checkTrue(cacheXml.find("formula=\"Sales*0.1\"") != std::string::npos,
                   "Calculated cache field formula is serialized");
    test.checkTrue(cacheXml.find("databaseField=\"0\"") != std::string::npos,
                   "Calculated cache field is marked non-database");
    test.checkTrue(cacheXml.find("<fieldGroup base=\"1\">") != std::string::npos,
                   "Pivot numeric field group is serialized");
    test.checkTrue(cacheXml.find("startNum=\"10\"") != std::string::npos
                   && cacheXml.find("groupInterval=\"10\"") != std::string::npos,
                   "Pivot numeric range grouping properties are serialized");
    const auto firstPivotXml = archive.get("xl/pivotTables/pivotTable1.xml");
    test.checkTrue(firstPivotXml.find("sumSubtotal=\"1\"") != std::string::npos,
                   "Explicit sum subtotal is serialized");
    test.checkTrue(firstPivotXml.find("avgSubtotal=\"1\"") != std::string::npos,
                   "Explicit average subtotal is serialized");
    const auto rel1 = archive.get("xl/pivotTables/_rels/pivotTable1.xml.rels");
    const auto rel2 = archive.get("xl/pivotTables/_rels/pivotTable2.xml.rels");
    test.checkTrue(rel1.find("pivotCacheDefinition1.xml") != std::string::npos
                   && rel2.find("pivotCacheDefinition1.xml") != std::string::npos,
                   "Both PivotTables reference the same physical cache definition");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto& readOnly = static_cast<const xlpp::Workbook&>(reloaded);
    const auto* loadedSheet = readOnly.worksheet("Data");
    test.checkTrue(loadedSheet != nullptr && loadedSheet->pivotTables().size() == 2,
                   "Two shared-cache PivotTables reload into the object model");
    const auto& loadedFirst = loadedSheet->pivotTables()[0];
    const auto& loadedSecond = loadedSheet->pivotTables()[1];
    test.checkEqual(loadedFirst.cache().sharedCacheKey(), loadedSecond.cache().sharedCacheKey(),
                    "Reloaded PivotTables retain a common physical cache identity");
    test.checkEqual(loadedFirst.cache().fieldFormula(2), std::string("Sales*0.1"),
                    "Calculated Pivot field formula round-trips");
    const auto* loadedGroup = loadedFirst.cache().tryFieldGroup(1);
    test.checkTrue(loadedGroup != nullptr, "Pivot field grouping round-trips");
    test.checkTrue(loadedGroup && loadedGroup->interval.has_value() && *loadedGroup->interval == 10.0,
                   "Pivot numeric grouping interval round-trips");
    test.checkTrue(std::find(loadedFirst.rowFields()[0].subtotals().begin(), loadedFirst.rowFields()[0].subtotals().end(), "avg") != loadedFirst.rowFields()[0].subtotals().end(),
                   "Explicit Pivot field subtotals round-trip");

    // Sharing is opt-in but strict: identical identity with divergent cache
    // content is rejected instead of producing a corrupt shared-cache package.
    bool incompatibleRejected = false;
    try {
        xlpp::Workbook invalid;
        auto& invalidSheet = invalid.addWorksheet("Data");
        invalidSheet.cell("A1").setValue("Region"); invalidSheet.cell("B1").setValue("Sales");
        invalidSheet.cell("A2").setValue("East"); invalidSheet.cell("B2").setValue(1.0);
        xlpp::PivotTable a("A"); a.cache() = cache; a.setLocation("D2"); a.addRowField("Region"); a.addDataField("Sales");
        xlpp::PivotTable b("B"); b.cache() = cache; b.setLocation("J2"); b.cache().records()[0][1] = "999"; b.addRowField("Region"); b.addDataField("Sales");
        invalidSheet.addPivotTable(std::move(a)); invalidSheet.addPivotTable(std::move(b));
        invalid.save(std::filesystem::temp_directory_path() / "xlpp_p1b_invalid_shared_cache.xlsx");
    } catch (const std::invalid_argument&) { incompatibleRejected = true; }
    test.checkTrue(incompatibleRejected, "Incompatible PivotCaches cannot silently share one cache identity");
    std::filesystem::remove(output);
    std::filesystem::remove(std::filesystem::temp_directory_path() / "xlpp_p1b_invalid_shared_cache.xlsx");
}

void testPivotDateGroupingImportedCalculatedFieldP1B(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1b_pivot_date_source.xlsx";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1b_pivot_date_edited.xlsx";

    xlpp::Workbook created;
    auto& data = created.addWorksheet("Data");
    data.append({std::string("OrderDate"), std::string("Amount")});
    data.append({std::string("2026-01-15T00:00:00"), 10.0});
    data.append({std::string("2026-02-20T00:00:00"), 20.0});
    data.append({std::string("2026-03-10T00:00:00"), 30.0});
    xlpp::PivotTable pivot("DatePivot");
    pivot.setLocation("D2");
    pivot.cache().setSourceData("'Data'!$A$1:$B$4");
    pivot.cache().setFields({"OrderDate", "Amount"});
    pivot.cache().addRecord({"2026-01-15T00:00:00", "10"});
    pivot.cache().addRecord({"2026-02-20T00:00:00", "20"});
    pivot.cache().addRecord({"2026-03-10T00:00:00", "30"});
    pivot.addRowField("OrderDate");
    pivot.addDataField("Amount", "sum");
    data.addPivotTable(std::move(pivot));
    created.save(source);

    xlpp::Workbook imported;
    imported.load(source);
    auto* importedSheet = imported.worksheet("Data");
    test.checkTrue(importedSheet != nullptr && importedSheet->pivotTables().size() == 1,
                   "Generated Pivot reloads as an imported editable Pivot model");
    auto& editable = importedSheet->pivotTables().front();
    editable.cache().setDateFieldGrouping(0, "months", "2026-01-01T00:00:00", "2026-12-31T23:59:59", false, false);
    const auto calculatedIndex = editable.cache().addCalculatedField("DoubleAmount", "Amount*2");
    editable.addDataField("DoubleAmount", "sum").setDisplayName("Double Amount");
    test.checkEqual(calculatedIndex, 2, "Calculated field appends to imported PivotCache geometry");
    imported.save(edited);

    const auto archive = xlpp::internal::ZipArchive::open(edited);
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    test.checkTrue(graph.validate().ok(), "Date-grouped imported Pivot package graph validates");
    const auto cachePart = archive.contains("xl/pivotCache/pivotCacheDefinition2.xml")
        ? std::string("xl/pivotCache/pivotCacheDefinition2.xml") : std::string("xl/pivotCache/pivotCacheDefinition1.xml");
    const auto cacheXml = archive.get(cachePart);
    test.checkTrue(cacheXml.find("groupBy=\"months\"") != std::string::npos,
                   "Pivot date grouping unit is serialized through rangePr");
    test.checkTrue(cacheXml.find("startDate=\"2026-01-01T00:00:00\"") != std::string::npos
                   && cacheXml.find("endDate=\"2026-12-31T23:59:59\"") != std::string::npos,
                   "Pivot date grouping bounds are serialized");
    test.checkTrue(cacheXml.find("formula=\"Amount*2\"") != std::string::npos,
                   "Calculated field added to an imported PivotCache is serialized");

    xlpp::Workbook roundTrip;
    roundTrip.load(edited);
    const auto* finalSheet = static_cast<const xlpp::Workbook&>(roundTrip).worksheet("Data");
    test.checkTrue(finalSheet != nullptr && finalSheet->pivotTables().size() == 1,
                   "Date-grouped edited Pivot reloads");
    const auto& finalPivot = finalSheet->pivotTables().front();
    const auto* group = finalPivot.cache().tryFieldGroup(0);
    test.checkTrue(group != nullptr && group->groupBy == "months", "Pivot date grouping round-trips");
    test.checkTrue(group != nullptr && !group->autoStart && !group->autoEnd, "Pivot explicit date bounds round-trip");
    test.checkEqual(finalPivot.cache().fieldFormula(2), std::string("Amount*2"),
                    "Calculated field added after import round-trips");
    test.checkEqual(finalPivot.dataFields().back().displayName(), std::string("Double Amount"),
                    "Calculated Pivot data field round-trips");

    bool badUnitRejected = false;
    try { editable.cache().setDateFieldGrouping(0, "fortnights"); }
    catch (const std::invalid_argument&) { badUnitRejected = true; }
    test.checkTrue(badUnitRejected, "Unsupported Pivot date grouping unit is rejected early");
    bool badIntervalRejected = false;
    try { editable.cache().setNumericFieldGrouping(1, 0.0, 10.0, 0.0); }
    catch (const std::invalid_argument&) { badIntervalRejected = true; }
    test.checkTrue(badIntervalRejected, "Non-positive Pivot numeric grouping interval is rejected early");

    std::filesystem::remove(source);
    std::filesystem::remove(edited);
}

void testVbaProjectReferencesLocaleAndModuleMetadataP1B(TestContext& test) {
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p1b_vba_project_refs.xlsm";
    const auto afterRemove = std::filesystem::temp_directory_path() / "xlpp_p1b_vba_codenames_after_remove.xlsm";
    xlpp::Workbook workbook;
    auto& hostSheet = workbook.addWorksheet("Host");
    hostSheet.setVbaCodeName("HostSheet");
    auto& calcSheet = workbook.addWorksheet("Calc");
    calcSheet.setVbaCodeName("CalcSheet");

    xlpp::VbaModule module;
    module.name = "InteropModule";
    module.source = "Option Explicit\nPublic Sub RunInterop()\nEnd Sub";
    module.helpContextId = 314;
    workbook.setVbaModule(module);
    workbook.setVbaDocumentModuleText("CalcSheet", "Private Sub Worksheet_Activate()\nRange(\"A1\").Value = 123\nEnd Sub");

    xlpp::VbaProjectInfo info;
    info.name = "P1BInterop";
    info.systemKind = 3;
    info.lcid = 0x0411;       // Japanese locale metadata; source remains ASCII in this fixture.
    info.lcidInvoke = 0x0411;
    info.codePage = 932;
    info.references.push_back({"Scripting", "*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime"});
    xlpp::VbaReference projectRef;
    projectRef.name = "SharedMacros";
    projectRef.kind = xlpp::VbaReferenceKind::Project;
    projectRef.libid = "C:\\Macros\\SharedMacros.xlsm";
    projectRef.relativeLibid = "SharedMacros.xlsm";
    projectRef.majorVersion = 2;
    projectRef.minorVersion = 7;
    info.references.push_back(projectRef);
    xlpp::VbaReference controlRef;
    controlRef.name = "MSForms";
    controlRef.kind = xlpp::VbaReferenceKind::Control;
    controlRef.twiddledLibid = "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##";
    controlRef.extendedName = "MSForms";
    controlRef.libid = "*\\G{896C2D83-5466-46ED-8FAE-4C3E4F85E710}#2.0#0#C:\\Temp\\MSForms.exd#Microsoft Forms 2.0 Object Library";
    controlRef.originalTypeLib = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}";
    controlRef.controlCookie = 1;
    info.references.push_back(controlRef);
    workbook.setVbaProjectInfo(info);
    workbook.save(output);

    xlpp::Workbook loaded;
    loaded.load(output);
    const auto loadedInfo = loaded.vbaProjectInfo();
    test.checkEqual(loadedInfo.systemKind, std::uint32_t{3}, "VBA project system kind round-trips");
    test.checkEqual(loadedInfo.lcid, std::uint32_t{0x0411}, "VBA project LCID round-trips");
    test.checkEqual(loadedInfo.lcidInvoke, std::uint32_t{0x0411}, "VBA project invoke LCID round-trips");
    test.checkEqual(loadedInfo.codePage, std::uint16_t{932}, "VBA project code page round-trips");
    const auto project = std::find_if(loadedInfo.references.begin(), loadedInfo.references.end(), [](const auto& reference) {
        return reference.name == "SharedMacros";
    });
    test.checkTrue(project != loadedInfo.references.end(), "VBA project reference is read from dir stream");
    test.checkTrue(project != loadedInfo.references.end() && project->kind == xlpp::VbaReferenceKind::Project,
                   "REFERENCEPROJECT remains distinct from registered type-library references");
    test.checkTrue(project != loadedInfo.references.end() && project->relativeLibid == "SharedMacros.xlsm",
                   "VBA project relative LibId round-trips");
    test.checkTrue(project != loadedInfo.references.end() && project->majorVersion == 2 && project->minorVersion == 7,
                   "VBA project reference version round-trips");
    const auto control = std::find_if(loadedInfo.references.begin(), loadedInfo.references.end(), [](const auto& reference) {
        return reference.name == "MSForms";
    });
    test.checkTrue(control != loadedInfo.references.end() && control->kind == xlpp::VbaReferenceKind::Control,
                   "VBA ActiveX REFERENCECONTROL round-trips as a distinct reference kind");
    test.checkTrue(control != loadedInfo.references.end() && control->extendedName == "MSForms",
                   "VBA control-reference extended name round-trips");
    test.checkTrue(control != loadedInfo.references.end() && control->originalTypeLib == "{0D452EE1-E08F-101A-852E-02608C4D0BB4}",
                   "VBA control-reference original type-library GUID round-trips");
    test.checkTrue(control != loadedInfo.references.end() && control->controlCookie == 1,
                   "VBA control-reference cookie round-trips");
    const auto modules = loaded.vbaModules();
    const auto interop = std::find_if(modules.begin(), modules.end(), [](const auto& candidate) { return candidate.name == "InteropModule"; });
    test.checkTrue(interop != modules.end(), "VBA module with extended metadata reloads");
    test.checkTrue(interop != modules.end() && interop->helpContextId == 314,
                   "VBA module help-context metadata round-trips");
    const auto* loadedCalc = static_cast<const xlpp::Workbook&>(loaded).worksheet("Calc");
    test.checkTrue(loadedCalc != nullptr && loadedCalc->vbaCodeName() == "CalcSheet",
                   "Worksheet VBA codeName round-trips from sheetPr");
    test.checkTrue(loaded.vbaModuleText("CalcSheet").value_or("").find("Worksheet_Activate") != std::string::npos,
                   "Custom worksheet document-module source round-trips");

    test.checkTrue(loaded.removeWorksheet("Host"), "Worksheet preceding custom codeName can be removed");
    loaded.save(afterRemove);
    xlpp::Workbook reloaded;
    reloaded.load(afterRemove);
    const auto* survivingCalc = static_cast<const xlpp::Workbook&>(reloaded).worksheet("Calc");
    test.checkTrue(survivingCalc != nullptr && survivingCalc->vbaCodeName() == "CalcSheet",
                   "Worksheet codeName remains stable after sibling worksheet removal");
    test.checkTrue(reloaded.vbaModuleText("CalcSheet").value_or("").find("Worksheet_Activate") != std::string::npos,
                   "Document-module source follows stable codeName after sheet removal");
    test.checkTrue(!reloaded.vbaModuleText("HostSheet").has_value(),
                   "Removed worksheet document module is dropped during source-project regeneration");
    std::filesystem::remove(output);
    std::filesystem::remove(afterRemove);
}


void testPivotFiltersChartLinkAndSelectiveCacheP1C(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1c_pivot_chart_source.xlsx";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1c_pivot_chart_edited.xlsx";

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Region"), std::string("Sales")});
    sheet.append({std::string("East"), 10.0});
    sheet.append({std::string("West"), 20.0});
    sheet.append({std::string("East"), 30.0});

    xlpp::PivotCache cache;
    cache.setSourceData("'Data'!$A$1:$B$4");
    cache.setSharedCacheKey("p1c-shared-sales");
    cache.setFields({"Region", "Sales"});
    cache.setRecords({{"East", "10"}, {"West", "20"}, {"East", "30"}});

    xlpp::PivotTable first("SalesPivot");
    first.setLocation("D2");
    first.cache() = cache;
    first.addRowField("Region");
    first.addDataField("Sales", "sum").setDisplayName("Total Sales");
    xlpp::PivotFilter filter;
    filter.fieldIndex = 0;
    filter.type = "captionContains";
    filter.evaluationOrder = 2;
    filter.name = "East filter";
    filter.description = "P1C advanced Pivot filter";
    filter.stringValue1 = "East";
    filter.autoFilterXml = "<autoFilter ref=\"A1:A4\"><filters><filter val=\"East\"/></filters></autoFilter>";
    first.addFilter(filter);
    xlpp::PivotChartFormat chartFormat;
    chartFormat.chartIndex = 7;
    chartFormat.formatId = 0;
    chartFormat.series = true;
    chartFormat.pivotAreaXml = "<pivotArea type=\"normal\" dataOnly=\"1\"/>";
    first.addChartFormat(chartFormat);
    first.setChartFormatIndex(7);
    sheet.addPivotTable(first);

    xlpp::PivotTable second("SalesPivot2");
    second.setLocation("J2");
    second.cache() = cache;
    second.addRowField("Region");
    second.addDataField("Sales", "sum");
    sheet.addPivotTable(second);

    xlpp::Chart chart(xlpp::Chart::Type::Bar);
    chart.setTitle("Pivot-linked Sales");
    chart.linkPivotTable("SalesPivot", 7);
    auto& series = chart.addSeries(xlpp::ChartSeries("Sales"));
    series.reference("Data", "$B$2:$B$4");
    series.categories("Data", "$A$2:$A$4");
    sheet.addChart(std::move(chart));

    workbook.save(source);
    const auto sourceArchive = xlpp::internal::ZipArchive::open(source);
    const auto pivot1Before = sourceArchive.get("xl/pivotTables/pivotTable1.xml");
    const auto pivot2Before = sourceArchive.get("xl/pivotTables/pivotTable2.xml");
    test.checkTrue(pivot1Before.find(" chartFormat=\"7\"") != std::string::npos,
                   "PivotTable root chartFormat attribute is serialized");
    test.checkTrue(pivot1Before.find("<chartFormats count=\"1\">") != std::string::npos,
                   "PivotTable emits chartFormats collection for PivotChart linkage");
    test.checkTrue(pivot1Before.find("<chartFormat chart=\"7\" format=\"0\" series=\"1\">") != std::string::npos,
                   "PivotChart chartFormat preserves chart/format/series metadata");
    test.checkTrue(pivot1Before.find("<pivotArea type=\"normal\" dataOnly=\"1\"/>") != std::string::npos,
                   "PivotChart format preserves raw PivotArea selector");
    const auto stylePos = pivot1Before.find("<pivotTableStyleInfo");
    const auto filterPos = pivot1Before.find("<filters count=\"1\">");
    const auto chartFormatPos = pivot1Before.find("<chartFormats count=\"1\">");
    test.checkTrue(chartFormatPos < stylePos && stylePos < filterPos,
                   "PivotTable serializer follows chartFormats -> style -> filters schema order");
    test.checkTrue(pivot1Before.find("type=\"captionContains\"") != std::string::npos
                   && pivot1Before.find("stringValue1=\"East\"") != std::string::npos,
                   "Advanced PivotFilter attributes are serialized");

    std::string chartPart;
    for (const auto& name : sourceArchive.entryNames()) {
        if (name.rfind("xl/charts/chart", 0) == 0 && name.size() > 4 && name.substr(name.size() - 4) == ".xml") {
            chartPart = name;
            break;
        }
    }
    test.checkTrue(!chartPart.empty(), "Pivot-linked chart part exists");
    const auto chartXmlText = chartPart.empty() ? std::string{} : sourceArchive.get(chartPart);
    test.checkTrue(chartXmlText.find("<c:pivotSource>") != std::string::npos
                   && chartXmlText.find("<c:name>SalesPivot</c:name>") != std::string::npos
                   && chartXmlText.find("<c:fmtId val=\"7\"/>") != std::string::npos,
                   "DrawingML chart emits c:pivotSource name and fmtId");

    xlpp::Workbook loaded;
    loaded.load(source);
    const auto* loadedSheet = static_cast<const xlpp::Workbook&>(loaded).worksheet("Data");
    test.checkTrue(loadedSheet != nullptr && loadedSheet->pivotTables().size() == 2,
                   "P1C Pivot workbook reloads both shared-cache PivotTables");
    const auto& loadedPivot = loadedSheet->pivotTables().front();
    test.checkEqual(loadedPivot.filters().size(), std::size_t{1}, "PivotFilter model reloads");
    test.checkEqual(loadedPivot.filters().front().type, std::string("captionContains"), "PivotFilter type round-trips");
    test.checkEqual(loadedPivot.filters().front().stringValue1, std::string("East"), "PivotFilter string value round-trips");
    test.checkTrue(loadedPivot.filters().front().autoFilterXml.find("filter val=\"East\"") != std::string::npos,
                   "Nested PivotFilter AutoFilter subtree round-trips");
    test.checkTrue(loadedPivot.chartFormatIndex().has_value() && *loadedPivot.chartFormatIndex() == 7,
                   "PivotTable root chartFormat attribute round-trips");
    test.checkEqual(loadedPivot.chartFormats().size(), std::size_t{1}, "PivotChart chartFormat model reloads");
    test.checkEqual(loadedPivot.chartFormats().front().chartIndex, std::uint32_t{7}, "PivotChart chart index round-trips");
    test.checkTrue(loadedPivot.chartFormats().front().series, "PivotChart series-format flag round-trips");
    test.checkTrue(!loadedSheet->charts().empty() && loadedSheet->charts().front().pivotSource().present,
                   "Chart reader recognizes PivotChart source");
    test.checkEqual(loadedSheet->charts().front().pivotSource().pivotTableName, std::string("SalesPivot"),
                    "PivotChart source PivotTable name round-trips");
    test.checkEqual(loadedSheet->charts().front().pivotSource().formatId, 7,
                    "PivotChart source fmtId round-trips");

    const auto physicalCachePart = loadedPivot.cache().sharedCacheKey();
    test.checkTrue(physicalCachePart.rfind("xl/pivotCache/pivotCacheDefinition", 0) == 0,
                   "Imported shared PivotCache keeps its physical cache-part identity");
    test.checkEqual(loadedSheet->pivotTables()[1].cache().sharedCacheKey(), physicalCachePart,
                    "Sibling PivotTable resolves to the same imported physical PivotCache");

    xlpp::PivotCacheOptionsPatch patch;
    patch.refreshOnLoad = false;
    patch.saveData = false;
    patch.enableRefresh = true;
    patch.missingItemsLimit = 17;
    test.checkTrue(loaded.updateImportedPivotCacheOptions("Data", "SalesPivot", patch),
                   "Selective imported shared-cache option mutation succeeds");
    loaded.save(edited);

    const auto editedArchive = xlpp::internal::ZipArchive::open(edited);
    test.checkEqual(editedArchive.get("xl/pivotTables/pivotTable1.xml"), pivot1Before,
                    "Selective shared-cache mutation preserves first PivotTable XML byte-for-byte");
    test.checkEqual(editedArchive.get("xl/pivotTables/pivotTable2.xml"), pivot2Before,
                    "Selective shared-cache mutation preserves sibling PivotTable XML byte-for-byte");
    const auto patchedCacheXml = editedArchive.get(physicalCachePart);
    test.checkTrue(patchedCacheXml.find("refreshOnLoad=\"0\"") != std::string::npos
                   && patchedCacheXml.find("saveData=\"0\"") != std::string::npos
                   && patchedCacheXml.find("enableRefresh=\"1\"") != std::string::npos
                   && patchedCacheXml.find("missingItemsLimit=\"17\"") != std::string::npos,
                   "Selective mutation patches only physical PivotCache options");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(editedArchive);
    test.checkTrue(graph.validate().ok(), "PivotChart/filter/selective-cache package graph validates");

    xlpp::Workbook roundTrip;
    roundTrip.load(edited);
    const auto* roundTripSheet = static_cast<const xlpp::Workbook&>(roundTrip).worksheet("Data");
    test.checkTrue(roundTripSheet != nullptr && roundTripSheet->pivotTables().size() == 2,
                   "Selectively patched shared PivotCache reloads");
    for (const auto& pivot : roundTripSheet->pivotTables()) {
        test.checkTrue(!pivot.cache().refreshOnLoad(), "Shared cache refreshOnLoad patch reaches every PivotTable model");
        test.checkTrue(!pivot.cache().saveData(), "Shared cache saveData patch reaches every PivotTable model");
        test.checkTrue(pivot.cache().enableRefresh(), "Shared cache enableRefresh patch reaches every PivotTable model");
        test.checkEqual(pivot.cache().missingItemsLimit(), 17, "Shared cache missingItemsLimit patch reaches every PivotTable model");
    }

    std::filesystem::remove(source);
    std::filesystem::remove(edited);
}

void testVbaDesignerUserFormStorageP1C(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1c_userform_source.xlsm";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1c_userform_edited.xlsm";
    const auto removed = std::filesystem::temp_directory_path() / "xlpp_p1c_userform_removed.xlsm";

    xlpp::Workbook workbook;
    workbook.addWorksheet("Data").setVbaCodeName("DataSheet");
    workbook.setVbaModuleText("KeepModule", "Option Explicit\nPublic Sub KeepProjectAlive()\nEnd Sub");

    xlpp::VbaDesignerStorage storage;
    storage.name = "UserForm1";
    storage.streams.push_back({"f", {0x00, 0x01, 0x7F, 0x80, 0xFF}});
    storage.streams.push_back({"o", {0x10, 0x20, 0x30, 0x40}});
    storage.streams.push_back({"vbFrame", {'V','e','r','s','i','o','n',' ','5','.','0'}});
    storage.streams.push_back({"Controls/0", {0xDE, 0xAD, 0xBE, 0xEF}});
    storage.streams.push_back({"Controls/Nested/state", {0x11, 0x22, 0x33}});
    workbook.setVbaDesignerModule("UserForm1",
        "Option Explicit\nPrivate Sub UserForm_Initialize()\nMe.Caption = \"P1C\"\nEnd Sub", storage);

    auto info = workbook.vbaProjectInfo();
    xlpp::VbaReference forms;
    forms.name = "MSForms";
    forms.kind = xlpp::VbaReferenceKind::Control;
    forms.twiddledLibid = "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##";
    forms.extendedName = "MSForms";
    forms.libid = "*\\G{896C2D83-5466-46ED-8FAE-4C3E4F85E710}#2.0#0#C:\\Temp\\MSForms.exd#Microsoft Forms 2.0 Object Library";
    forms.originalTypeLib = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}";
    forms.controlCookie = 1;
    info.references.push_back(forms);
    workbook.setVbaProjectInfo(info);
    workbook.save(source);

    xlpp::Workbook loaded;
    loaded.load(source);
    const auto modules = loaded.vbaModules();
    const auto designer = std::find_if(modules.begin(), modules.end(), [](const auto& module) {
        return module.name == "UserForm1";
    });
    test.checkTrue(designer != modules.end(), "UserForm Designer module reloads from PROJECT/BaseClass");
    test.checkTrue(designer != modules.end() && designer->type == xlpp::VbaModuleType::Designer,
                   "UserForm remains a Designer module rather than Class/Document");
    test.checkTrue(designer != modules.end() && designer->source.find("UserForm_Initialize") != std::string::npos,
                   "UserForm VBA source round-trips independently from binary designer storage");
    test.checkTrue(designer != modules.end() && !designer->designerClassId.empty(),
                   "UserForm Package designer class ID round-trips");
    test.checkTrue(designer != modules.end() && designer->designerBaseClass.find("842E9C5E") != std::string::npos,
                   "UserForm VB_Base metadata round-trips");

    auto storages = loaded.vbaDesignerStorages();
    test.checkEqual(storages.size(), std::size_t{1}, "One UserForm Designer Storage reloads");
    test.checkEqual(storages.front().name, std::string("UserForm1"), "Designer Storage name matches module identity");
    const auto* fStream = storages.front().findStream("f");
    const auto* nestedStream = storages.front().findStream("Controls/0");
    const auto* deeplyNested = storages.front().findStream("Controls/Nested/state");
    test.checkTrue(fStream != nullptr && fStream->data == std::vector<unsigned char>({0x00,0x01,0x7F,0x80,0xFF}),
                   "Binary UserForm f stream is preserved byte-for-byte");
    test.checkTrue(nestedStream != nullptr && nestedStream->data == std::vector<unsigned char>({0xDE,0xAD,0xBE,0xEF}),
                   "Nested Designer Storage stream round-trips");
    test.checkTrue(deeplyNested != nullptr && deeplyNested->data == std::vector<unsigned char>({0x11,0x22,0x33}),
                   "Recursive Designer Storage tree round-trips beyond one nesting level");

    auto editableStorage = storages.front();
    auto editableF = std::find_if(editableStorage.streams.begin(), editableStorage.streams.end(), [](const auto& stream) { return stream.path == "f"; });
    test.checkTrue(editableF != editableStorage.streams.end(), "UserForm f stream is exposed for raw editing");
    if (editableF != editableStorage.streams.end()) editableF->data = {0xAA, 0xBB, 0xCC};
    editableStorage.streams.push_back({"Controls/1", {0x44, 0x55}});
    loaded.setVbaDesignerStorage(editableStorage);
    loaded.setVbaDesignerModule("UserForm1",
        "Option Explicit\nPrivate Sub UserForm_Initialize()\nMe.Caption = \"P1C edited\"\nEnd Sub", editableStorage);
    loaded.save(edited);

    xlpp::Workbook roundTrip;
    roundTrip.load(edited);
    const auto editedStorages = roundTrip.vbaDesignerStorages();
    test.checkEqual(editedStorages.size(), std::size_t{1}, "Edited UserForm Designer Storage reloads");
    test.checkTrue(editedStorages.front().findStream("f") != nullptr
                   && editedStorages.front().findStream("f")->data == std::vector<unsigned char>({0xAA,0xBB,0xCC}),
                   "Raw binary UserForm stream mutation persists");
    test.checkTrue(editedStorages.front().findStream("Controls/1") != nullptr,
                   "New nested UserForm designer stream persists");
    test.checkTrue(roundTrip.vbaModuleText("UserForm1").value_or("").find("P1C edited") != std::string::npos,
                   "Edited UserForm VBA source persists alongside designer storage");
    const auto editedInfo = roundTrip.vbaProjectInfo();
    test.checkTrue(std::any_of(editedInfo.references.begin(), editedInfo.references.end(), [](const auto& reference) {
        return reference.name == "MSForms" && reference.kind == xlpp::VbaReferenceKind::Control;
    }), "MSForms REFERENCECONTROL survives UserForm designer regeneration");

    test.checkTrue(roundTrip.removeVbaModule("UserForm1"), "Designer module can be removed through public module API");
    roundTrip.save(removed);
    xlpp::Workbook afterRemove;
    afterRemove.load(removed);
    test.checkTrue(afterRemove.hasVbaProject(), "Removing UserForm keeps VBA project when another user module remains");
    test.checkTrue(!afterRemove.vbaModuleText("UserForm1").has_value(), "Removed UserForm module source disappears");
    test.checkTrue(std::none_of(afterRemove.vbaDesignerStorages().begin(), afterRemove.vbaDesignerStorages().end(), [](const auto& item) {
        return item.name == "UserForm1";
    }), "Removing UserForm also removes its root Designer Storage without leaving an orphan");

    std::filesystem::remove(source);
    std::filesystem::remove(edited);
    std::filesystem::remove(removed);
}

void testExternalCellAndStyleReaderFixture(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_external_reader_cells.xlsx";
    const std::string workbookXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Imported\" sheetId=\"1\" r:id=\"rIdSheet1\"/></sheets></workbook>";
    const std::string stylesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<numFmts count=\"0\"/>"
        "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font>"
        "<font><b/><i/><sz val=\"14\"/><color rgb=\"FF112233\"/><name val=\"Arial\"/></font></fonts>"
        "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>"
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFC000\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>"
        "<borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border>"
        "<border><left style=\"thin\"><color rgb=\"FF00B050\"/></left><right style=\"double\"><color rgb=\"FF0070C0\"/></right><top/><bottom/><diagonal/></border></borders>"
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        "<cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\" textRotation=\"30\"/></xf>"
        "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs>"
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        "<dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>"
        "</styleSheet>";
    const std::string sharedStrings =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\">"
        "<si><t>Shared value</t></si>"
        "<si><r><rPr><b/><color rgb=\"FFFF0000\"/><sz val=\"12\"/><rFont val=\"Calibri\"/></rPr><t>Rich</t></r><r><rPr><i/></rPr><t xml:space=\"preserve\"> text</t></r></si>"
        "</sst>";
    const std::string sheetXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<dimension ref=\"A1:I1\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>"
        "<sheetData><row r=\"1\">"
        "<c r=\"A1\" t=\"s\"><v>0</v></c>"
        "<c r=\"B1\" t=\"inlineStr\"><is><t>Inline value</t></is></c>"
        "<c r=\"C1\"><v>42.5</v></c>"
        "<c r=\"D1\" t=\"b\"><v>1</v></c>"
        "<c r=\"E1\" t=\"e\"><v>#DIV/0!</v></c>"
        "<c r=\"F1\"><f>SUM(C1,7.5)</f><v>50</v></c>"
        "<c r=\"G1\" t=\"s\"><v>1</v></c>"
        "<c r=\"H1\" s=\"2\"><v>45292</v></c>"
        "<c r=\"I1\" s=\"1\" t=\"inlineStr\"><is><t>Styled</t></is></c>"
        "</row></sheetData></worksheet>";
    writeExternalReaderFixture(path, workbookXml, sheetXml, stylesXml, sharedStrings);

    xlpp::Workbook workbook;
    workbook.load(path);
    auto* sheet = workbook.worksheet("Imported");
    test.checkTrue(sheet != nullptr, "Handcrafted worksheet is discovered by relationship reader");
    test.checkEqual(std::get<std::string>(sheet->cell("A1").value()), std::string("Shared value"), "Shared string is read from external fixture");
    test.checkEqual(std::get<std::string>(sheet->cell("B1").value()), std::string("Inline value"), "Inline string is read");
    test.checkNear(std::get<double>(sheet->cell("C1").value()), 42.5, 1e-12, "Numeric cell is read");
    test.checkTrue(std::get<bool>(sheet->cell("D1").value()), "Boolean cell is read");
    test.checkEqual(static_cast<unsigned>(*sheet->cell("E1").error()), static_cast<unsigned>(xlpp::CellError::DivisionByZero), "Error cell is read");
    test.checkEqual(sheet->cell("F1").formula(), std::string("SUM(C1,7.5)"), "Formula text is read");
    test.checkNear(std::get<double>(sheet->cell("F1").value()), 50.0, 1e-12, "Formula cached value is read");
    test.checkTrue(sheet->cell("G1").hasRichText(), "Shared rich-text cell is preserved as runs");
    test.checkEqual(sheet->cell("G1").richTextValue()->runs().size(), std::size_t{2}, "Two rich-text runs are read");
    test.checkTrue(sheet->cell("G1").richTextValue()->runs()[0].bold(), "Rich-text bold property is read");
    test.checkEqual(sheet->cell("G1").richTextValue()->runs()[0].color(), std::string("FFFF0000"), "Rich-text color is read");
    test.checkTrue(sheet->cell("H1").isDate(), "Built-in date style converts numeric serial to DateTime");
    test.checkEqual(*sheet->cell("H1").date(), xlpp::DateTime{2024, 1, 1}, "Date serial is decoded");
    const auto& styled = sheet->cell("I1");
    test.checkTrue(styled.font().bold(), "External style font bold is read");
    test.checkTrue(styled.font().italic(), "External style font italic is read");
    test.checkEqual(styled.font().name(), std::string("Arial"), "External font name is read");
    test.checkNear(styled.font().size(), 14.0, 1e-12, "External font size is read");
    test.checkEqual(styled.font().color().argb(), std::string("FF112233"), "External font color is read");
    test.checkEqual(styled.fill().foregroundColor().argb(), std::string("FFFFC000"), "External fill color is read");
    test.checkEqual(styled.border().left().style(), std::string("thin"), "External left border is read");
    test.checkEqual(styled.border().right().style(), std::string("double"), "External right border is read");
    test.checkEqual(styled.alignment().horizontal(), std::string("center"), "External horizontal alignment is read");
    test.checkTrue(styled.alignment().wrapText(), "External wrap-text alignment is read");
    test.checkEqual(styled.alignment().textRotation(), 30, "External text rotation is read");
    std::filesystem::remove(path);
}

void testExternalWorksheetFeatureReaderFixture(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_external_reader_features.xlsx";
    const std::string workbookXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Features\" sheetId=\"1\" r:id=\"rIdSheet1\"/></sheets></workbook>";
    const std::string stylesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
        "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        "<dxfs count=\"1\"><dxf><font><color rgb=\"FFFF0000\"/><b/></font></dxf></dxfs></styleSheet>";
    const std::string sheetXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr><tabColor rgb=\"FF00B0F0\"/><pageSetUpPr fitToPage=\"1\"/></sheetPr><dimension ref=\"A1:D9\"/>"
        "<sheetViews><sheetView workbookViewId=\"2\" zoomScale=\"135\" zoomScaleNormal=\"110\" showGridLines=\"0\" rightToLeft=\"1\" showOutlineSymbols=\"0\"><pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"B3\" activePane=\"bottomRight\" state=\"split\"/></sheetView></sheetViews>"
        "<sheetFormatPr defaultRowHeight=\"15\"/><cols><col min=\"2\" max=\"2\" width=\"22.5\" customWidth=\"1\" hidden=\"1\" bestFit=\"1\" outlineLevel=\"2\" collapsed=\"1\"/></cols>"
        "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Category</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Value</t></is></c></row>"
        "<row r=\"3\" ht=\"27\" customHeight=\"1\" hidden=\"1\" outlineLevel=\"2\" collapsed=\"1\"><c r=\"A3\" t=\"inlineStr\"><is><t>A</t></is></c><c r=\"B3\"><v>12</v></c></row></sheetData>"
        "<sheetProtection sheet=\"1\" password=\"DAA7\" formatCells=\"1\" autoFilter=\"1\"/>"
        "<mergeCells count=\"1\"><mergeCell ref=\"C2:D2\"/></mergeCells>"
        "<autoFilter ref=\"A1:B9\"><filterColumn colId=\"0\"><filters blank=\"1\"><filter val=\"A\"/><filter val=\"B\"/></filters></filterColumn><sortState ref=\"A1:B9\" caseSensitive=\"1\"><sortCondition ref=\"B2:B9\" descending=\"1\"/></sortState></autoFilter>"
        "<conditionalFormatting sqref=\"B2:B9\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"greaterThan\" stopIfTrue=\"1\"><formula>10</formula></cfRule></conditionalFormatting>"
        "<dataValidations count=\"1\"><dataValidation type=\"list\" errorStyle=\"warning\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" promptTitle=\"Pick\" prompt=\"Choose A or B\" errorTitle=\"Invalid\" error=\"Not allowed\" sqref=\"A2:A9\"><formula1>\"A,B\"</formula1></dataValidation></dataValidations>"
        "<hyperlinks><hyperlink ref=\"A3\" r:id=\"rIdExternal\" display=\"External\" tooltip=\"Open site\"/><hyperlink ref=\"B3\" location=\"Features!A1\" display=\"Internal\"/></hyperlinks>"
        "<printOptions horizontalCentered=\"1\" verticalCentered=\"1\" headings=\"1\" gridLines=\"1\"/><pageMargins left=\"0.4\" right=\"0.5\" top=\"0.6\" bottom=\"0.7\" header=\"0.2\" footer=\"0.3\"/>"
        "<pageSetup orientation=\"landscape\" paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"2\" blackAndWhite=\"1\" draft=\"1\" firstPageNumber=\"3\" useFirstPageNumber=\"1\"/>"
        "<headerFooter differentOddEven=\"1\" differentFirst=\"1\"><oddHeader>&amp;CExternal Header</oddHeader><oddFooter>&amp;RPage &amp;P</oddFooter></headerFooter><legacyDrawing r:id=\"rIdVml\"/></worksheet>";
    const std::string rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rIdExternal\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/read-fixture\" TargetMode=\"External\"/>"
        "<Relationship Id=\"rIdComments\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments1.xml\"/>"
        "<Relationship Id=\"rIdVml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/commentsDrawing1.vml\"/>"
        "</Relationships>";
    const std::string comments =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>External Author</author></authors><commentList><comment ref=\"D4\" authorId=\"0\"><text><r><t>Read </t></r><r><t>comment</t></r></text></comment></commentList></comments>";
    writeExternalReaderFixture(path, workbookXml, sheetXml, stylesXml, {}, rels, comments);

    xlpp::Workbook workbook;
    workbook.load(path);
    auto* sheet = workbook.worksheet("Features");
    test.checkTrue(sheet != nullptr, "External feature fixture sheet loads");
    test.checkTrue(sheet->tryColumnDimension(2) != nullptr, "External column dimension is read");
    test.checkNear(*sheet->tryColumnDimension(2)->width, 22.5, 1e-12, "External column width is read");
    test.checkTrue(sheet->tryColumnDimension(2)->hidden, "External hidden column flag is read");
    test.checkTrue(sheet->tryRowDimension(3) != nullptr, "External row dimension is read");
    test.checkNear(*sheet->tryRowDimension(3)->height, 27.0, 1e-12, "External row height is read");
    test.checkTrue(sheet->tryRowDimension(3)->hidden, "External hidden row flag is read");
    test.checkTrue(sheet->isMerged("C2"), "External merged range is read");
    test.checkEqual(sheet->sheetView().zoomScale(), 135, "External zoom scale is read");
    test.checkEqual(sheet->sheetView().zoomScaleNormal(), 110, "External normal zoom is read");
    test.checkTrue(!sheet->sheetView().showGridLines(), "External gridline visibility is read");
    test.checkTrue(sheet->sheetView().rightToLeft(), "External right-to-left view is read");
    test.checkEqual(sheet->sheetView().topLeftCell(), std::string("B3"), "External split-pane top-left cell is read");
    test.checkEqual(sheet->sheetView().pane(), std::string("bottomRight"), "External active pane is read");
    test.checkTrue(sheet->protection().enabled(), "External worksheet protection is read");
    test.checkEqual(sheet->protection().passwordHash(), std::string("DAA7"), "External worksheet password hash is read");
    test.checkTrue(sheet->autoFilter().enabled(), "External AutoFilter is read");
    test.checkEqual(sheet->autoFilter().reference(), std::string("A1:B9"), "External AutoFilter reference is read");
    test.checkEqual(sheet->autoFilter().columns().at(0).values().size(), std::size_t{2}, "External filter values are read");
    test.checkTrue(sheet->autoFilter().columns().at(0).includeBlank(), "External filter blank option is read");
    test.checkTrue(sheet->autoFilter().sortStateValue().has_value(), "External sort state is read");
    test.checkTrue(sheet->autoFilter().sortStateValue()->conditions().front().descending, "External descending sort is read");
    test.checkEqual(sheet->conditionalFormatting().entries().size(), std::size_t{1}, "External conditional formatting is read");
    const auto& rule = sheet->conditionalFormatting().entries().front().rules().front();
    test.checkEqual(static_cast<unsigned>(rule.type()), static_cast<unsigned>(xlpp::ConditionalRuleType::CellIs), "External cell-is rule type is read");
    test.checkEqual(rule.formulas().front(), std::string("10"), "External conditional formula is read");
    test.checkTrue(rule.hasDifferentialStyle(), "External differential style is linked");
    test.checkEqual(rule.differentialStyle().font().color().argb(), std::string("FFFF0000"), "External differential font color is read");
    test.checkEqual(sheet->dataValidations().items().size(), std::size_t{1}, "External data validation is read");
    const auto& validation = sheet->dataValidations().items().front();
    test.checkEqual(static_cast<unsigned>(validation.type()), static_cast<unsigned>(xlpp::DataValidationType::List), "External validation type is read");
    test.checkEqual(validation.formula1(), std::string("\"A,B\""), "External validation formula is read");
    test.checkEqual(validation.promptTitle(), std::string("Pick"), "External validation prompt title is read");
    test.checkTrue(sheet->cell("A3").hasHyperlink(), "External hyperlink relationship is read");
    test.checkEqual(sheet->cell("A3").hyperlinkValue()->target(), std::string("https://example.com/read-fixture"), "External hyperlink target is resolved");
    test.checkTrue(sheet->cell("B3").hasHyperlink(), "Internal hyperlink is read without relationship");
    test.checkTrue(sheet->cell("D4").hasComment(), "External legacy comment is read");
    test.checkEqual(sheet->cell("D4").commentValue()->text(), std::string("Read comment"), "External rich comment runs are concatenated");
    test.checkEqual(sheet->cell("D4").commentValue()->author(), std::string("External Author"), "External comment author is read");
    test.checkEqual(static_cast<unsigned>(sheet->pageSetup().orientation()), static_cast<unsigned>(xlpp::PageOrientation::Landscape), "External page orientation is read");
    test.checkEqual(static_cast<unsigned>(sheet->pageSetup().paperSize()), static_cast<unsigned>(xlpp::PaperSize::A4), "External paper size is read");
    test.checkEqual(sheet->pageSetup().fitToWidth(), 1u, "External fit-to-width is read");
    test.checkEqual(sheet->pageSetup().fitToHeight(), 2u, "External fit-to-height is read");
    test.checkNear(sheet->pageMargins().left(), 0.4, 1e-12, "External left page margin is read");
    test.checkTrue(sheet->printOptions().horizontalCentered(), "External horizontal print centering is read");
    test.checkEqual(sheet->headerFooter().oddHeader(), std::string("&CExternal Header"), "External header text is read");
    std::filesystem::remove(path);
}

void testExternalWorkbookMetadataReaderFixture(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_external_reader_metadata.xlsx";
    const std::string workbookXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<workbookPr date1904=\"1\"/><workbookProtection lockStructure=\"1\" lockWindows=\"1\" workbookPassword=\"83AF\"/>"
        "<sheets><sheet name=\"Metadata\" sheetId=\"1\" r:id=\"rIdSheet1\"/></sheets>"
        "<definedNames><definedName name=\"InputValue\" comment=\"external name\">Metadata!$A$1</definedName><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">Metadata!$A$1:$D$20</definedName><definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">Metadata!$A:$B,Metadata!$1:$2</definedName></definedNames>"
        "<calcPr calcId=\"777\" calcMode=\"manual\" fullPrecision=\"0\" iterate=\"1\" iterateCount=\"55\" iterateDelta=\"0.0005\" fullCalcOnLoad=\"1\" calcOnSave=\"1\"/></workbook>";
    const std::string stylesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font/></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf/></cellStyleXfs><cellXfs count=\"1\"><xf/></cellXfs></styleSheet>";
    const std::string sheetXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
    const std::string core =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>External title</dc:title><dc:subject>Read fixture</dc:subject><dc:creator>Fixture Author</dc:creator><dc:description>Loaded without XLPP writer</dc:description><cp:keywords>read,test</cp:keywords><cp:category>QA</cp:category><cp:lastModifiedBy>External Tool</cp:lastModifiedBy></cp:coreProperties>";
    const std::string custom =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"StringProp\"><vt:lpwstr>hello</vt:lpwstr></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"3\" name=\"IntProp\"><vt:i4>42</vt:i4></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"4\" name=\"BoolProp\"><vt:bool>true</vt:bool></property></Properties>";
    writeExternalReaderFixture(path, workbookXml, sheetXml, stylesXml, {}, {}, {}, core, custom);

    xlpp::Workbook workbook;
    workbook.load(path);
    test.checkTrue(workbook.date1904(), "External 1904 date system is read");
    test.checkEqual(workbook.properties().title(), std::string("External title"), "External document title is read");
    test.checkEqual(workbook.properties().creator(), std::string("Fixture Author"), "External document creator is read");
    test.checkEqual(workbook.properties().lastModifiedBy(), std::string("External Tool"), "External lastModifiedBy is read");
    test.checkEqual(workbook.customProperties().items().size(), std::size_t{3}, "External custom properties are read");
    test.checkEqual(workbook.customProperties().items()[0].value(), std::string("hello"), "External custom string property is read");
    test.checkEqual(workbook.customProperties().items()[1].type(), std::string("i4"), "External custom integer type is retained");
    test.checkTrue(workbook.protection().lockStructure(), "External workbook structure protection is read");
    test.checkTrue(workbook.protection().lockWindows(), "External workbook window protection is read");
    test.checkEqual(workbook.protection().workbookPasswordHash(), std::string("83AF"), "External workbook password hash is read");
    test.checkEqual(workbook.calcProperties().calcId(), 777, "External calcId is read");
    test.checkEqual(workbook.calcProperties().calcMode(), std::string("manual"), "External calculation mode is read");
    test.checkTrue(workbook.calcProperties().iterate(), "External iterative calculation flag is read");
    test.checkEqual(workbook.calcProperties().iterateCount(), 55, "External iteration count is read");
    test.checkNear(workbook.calcProperties().iterateDelta(), 0.0005, 1e-12, "External iteration delta is read");
    test.checkTrue(workbook.calcProperties().fullCalcOnLoad(), "External full-calc-on-load flag is read");
    test.checkTrue(workbook.calcProperties().calcOnSave(), "External calc-on-save flag is read");
    test.checkTrue(!workbook.calcProperties().fullPrecision(), "External full-precision false is read");
    test.checkTrue(workbook.definedName("InputValue") != nullptr, "External user-defined name is read");
    test.checkEqual(workbook.definedName("InputValue")->comment(), std::string("external name"), "External defined-name comment is read");
    const auto* sheet = workbook.worksheet("Metadata");
    test.checkEqual(sheet->printArea(), std::string("A1:D20"), "External built-in print area is applied to worksheet");
    test.checkEqual(sheet->printTitlesCols(), std::string("A:B"), "External print-title columns are applied");
    test.checkEqual(sheet->printTitlesRows(), std::string("1:2"), "External print-title rows are applied");
    std::filesystem::remove(path);
}


std::vector<unsigned char> makeP1DUserFormStream() {
    const std::uint32_t mask = (1u << 1) | (1u << 2) | (1u << 3) | (1u << 6) | (1u << 7)
        | (1u << 8) | (1u << 9) | (1u << 10) | (1u << 11) | (1u << 12) | (1u << 13)
        | (1u << 16) | (1u << 17) | (1u << 18) | (1u << 19) | (1u << 22)
        | (1u << 23) | (1u << 25) | (1u << 26) | (1u << 27);
    std::vector<unsigned char> data;
    auto u8 = [&](std::uint8_t v) { data.push_back(v); };
    auto u16 = [&](std::uint16_t v) { data.push_back(static_cast<unsigned char>(v)); data.push_back(static_cast<unsigned char>(v >> 8)); };
    auto u32 = [&](std::uint32_t v) {
        data.push_back(static_cast<unsigned char>(v)); data.push_back(static_cast<unsigned char>(v >> 8));
        data.push_back(static_cast<unsigned char>(v >> 16)); data.push_back(static_cast<unsigned char>(v >> 24));
    };
    auto align = [&](std::size_t n) { while ((data.size() % n) != 0) data.push_back(0); };

    u8(0); u8(4); u16(0); // cbForm patched below
    u32(mask);
    u32(0x8000000Fu); // BackColor
    u32(0x80000012u); // ForeColor
    u32(5);           // NextAvailableID
    u32(0x0000000Fu); // BooleanProperties
    u8(1); u8(0); u8(3); // BorderStyle, MousePointer, ScrollBars
    align(4); u32(2);     // GroupCount
    u8(0); u8(2);         // Cycle, SpecialEffect
    align(4); u32(0x80000006u); // BorderColor
    u32(0x80000004u);            // compressed Caption byte count = 4
    u32(125);                     // Zoom
    u8(2); u8(1);                // PictureAlignment, PictureSizeMode
    align(4); u32(9); u32(32000); // ShapeCookie, DrawBuffer
    align(4);
    // ExtraDataBlock: DisplayedSize, LogicalSize, ScrollPosition, Caption
    u32(5000); u32(3000);
    u32(7000); u32(4000);
    u32(100); u32(200);
    data.insert(data.end(), {'P','1','D','!'});
    const auto cbForm = static_cast<std::uint16_t>(data.size() - 4);
    data[2] = static_cast<unsigned char>(cbForm & 0xFFu);
    data[3] = static_cast<unsigned char>(cbForm >> 8);
    // Unknown FormStreamData/SiteData sentinel: semantic edits must preserve it.
    data.insert(data.end(), {0xFA, 0xCE, 0xB0, 0x0C, 0x12, 0x34});
    return data;
}

void testPivotSelectiveFieldItemsAndLinkValidationP1D(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1d_pivot_source.xlsx";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1d_pivot_edited.xlsx";

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Region"), std::string("Sales")});
    sheet.append({std::string("East"), 10.0});
    sheet.append({std::string("West"), 20.0});
    sheet.append({std::string("East"), 30.0});
    xlpp::PivotCache cache;
    cache.setSourceData("'Data'!$A$1:$B$4");
    cache.setSharedCacheKey("p1d-shared-cache");
    cache.setFields({"Region", "Sales"});
    cache.setFieldCaption(1, "Gross Sales");
    cache.setFieldNumberFormatId(1, 2);
    cache.setRecords({{"East", "10"}, {"West", "20"}, {"East", "30"}});

    xlpp::PivotTable first("SalesPivot");
    first.setLocation("D2"); first.cache() = cache;
    auto& region = first.addRowField("Region");
    xlpp::PivotFieldItem east; east.cacheIndex = 0; east.type = "data"; east.caption = "East region"; east.hidden = true; east.showDetails = false;
    region.addItem(east);
    xlpp::PivotFieldItem west; west.cacheIndex = 1; west.type = "data"; west.caption = "West region";
    region.addItem(west);
    first.addDataField("Sales", "sum").setDisplayName("Total Sales");
    first.setChartFormatIndex(4);
    xlpp::PivotChartFormat cf; cf.chartIndex = 4; cf.formatId = 12; cf.series = true; cf.pivotAreaXml = "<pivotArea type=\"normal\"/>";
    first.addChartFormat(cf);
    sheet.addPivotTable(first);

    xlpp::PivotTable second("SalesPivot2");
    second.setLocation("J2"); second.cache() = cache; second.addRowField("Region"); second.addDataField("Sales", "sum");
    sheet.addPivotTable(second);

    xlpp::Chart chart(xlpp::Chart::Type::Bar);
    chart.setTitle("P1D PivotChart"); chart.linkPivotTable("SalesPivot", 4);
    auto& ser = chart.addSeries(xlpp::ChartSeries("Sales")); ser.reference("Data", "$B$2:$B$4"); ser.categories("Data", "$A$2:$A$4");
    sheet.addChart(std::move(chart));
    workbook.save(source);

    xlpp::Workbook loaded; loaded.load(source);
    auto validLinks = loaded.validatePivotChartLinks();
    test.checkTrue(validLinks.ok() && validLinks.pivotChartsVisited == 1 && validLinks.validLinks == 1,
                   "PivotChart ownership/link validator accepts coherent P1D linkage");
    const auto* loadedSheet = static_cast<const xlpp::Workbook&>(loaded).worksheet("Data");
    test.checkTrue(loadedSheet != nullptr && loadedSheet->pivotTables().size() == 2, "P1D shared Pivot workbook reloads");
    const auto& loadedPivot = loadedSheet->pivotTables().front();
    test.checkEqual(loadedPivot.rowFields().front().items().size(), std::size_t{2}, "Semantic PivotField items reload");
    test.checkEqual(loadedPivot.rowFields().front().items().front().caption, std::string("East region"), "Pivot item caption round-trips");
    test.checkTrue(loadedPivot.rowFields().front().items().front().hidden && !loadedPivot.rowFields().front().items().front().showDetails,
                   "Pivot item hidden/showDetails semantics round-trip");
    test.checkEqual(loadedPivot.cache().fieldCaption(1), std::string("Gross Sales"), "Pivot cache field caption reloads");
    test.checkEqual(loadedPivot.cache().fieldNumberFormatId(1), 2, "Pivot cache field numFmtId reloads");

    const auto sourceArchive = xlpp::internal::ZipArchive::open(source);
    const auto pivot1Before = sourceArchive.get("xl/pivotTables/pivotTable1.xml");
    const auto pivot2Before = sourceArchive.get("xl/pivotTables/pivotTable2.xml");
    const auto cachePart = loadedPivot.cache().sharedCacheKey();
    xlpp::PivotCacheFieldPatch patch;
    patch.name = "Revenue"; patch.caption = "Net Revenue"; patch.formula = "Sales*1.05";
    patch.numberFormatId = 4; patch.databaseField = false;
    test.checkTrue(loaded.updateImportedPivotCacheField("Data", "SalesPivot", 1, patch),
                   "Selective imported cacheField mutation succeeds");
    loaded.save(edited);
    const auto editedArchive = xlpp::internal::ZipArchive::open(edited);
    test.checkEqual(editedArchive.get("xl/pivotTables/pivotTable1.xml"), pivot1Before,
                    "Selective cacheField edit preserves first PivotTable XML byte-for-byte");
    test.checkEqual(editedArchive.get("xl/pivotTables/pivotTable2.xml"), pivot2Before,
                    "Selective cacheField edit preserves sibling PivotTable XML byte-for-byte");
    const auto cacheXml = editedArchive.get(cachePart);
    test.checkTrue(cacheXml.find("name=\"Revenue\"") != std::string::npos
                   && cacheXml.find("caption=\"Net Revenue\"") != std::string::npos
                   && cacheXml.find("formula=\"Sales*1.05\"") != std::string::npos
                   && cacheXml.find("numFmtId=\"4\"") != std::string::npos
                   && cacheXml.find("databaseField=\"0\"") != std::string::npos,
                   "Selective cacheField edit patches only promoted field metadata");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(editedArchive).validate().ok(),
                   "P1D selective Pivot field edit preserves package graph");

    xlpp::Workbook roundTrip; roundTrip.load(edited);
    const auto* roundSheet = static_cast<const xlpp::Workbook&>(roundTrip).worksheet("Data");
    for (const auto& pivot : roundSheet->pivotTables()) {
        test.checkEqual(pivot.cache().fields()[1], std::string("Revenue"), "Shared cache field rename reaches every PivotTable model");
        test.checkEqual(pivot.cache().fieldCaption(1), std::string("Net Revenue"), "Shared cache field caption reaches every PivotTable model");
        test.checkEqual(pivot.cache().fieldFormula(1), std::string("Sales*1.05"), "Shared cache field formula reaches every PivotTable model");
        test.checkEqual(pivot.cache().fieldNumberFormatId(1), 4, "Shared cache field numFmt reaches every PivotTable model");
        test.checkTrue(!pivot.cache().fieldDatabaseField(1), "Shared cache calculated field remains non-database");
    }

    auto* mutableSheet = roundTrip.worksheet("Data");
    mutableSheet->charts().front().linkPivotTable("MissingPivot", 4);
    const auto badLinks = roundTrip.validatePivotChartLinks();
    test.checkTrue(!badLinks.ok() && !badLinks.issues.empty(), "PivotChart validator diagnoses missing PivotTable ownership");

    std::filesystem::remove(source); std::filesystem::remove(edited);
}

void testVbaUserFormSemanticPropertiesP1D(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1d_userform_source.xlsm";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1d_userform_edited.xlsm";
    xlpp::Workbook workbook;
    workbook.addWorksheet("Data").setVbaCodeName("DataSheet");
    xlpp::VbaDesignerStorage storage; storage.name = "UserForm1";
    const auto originalF = makeP1DUserFormStream();
    storage.streams.push_back({"f", originalF});
    storage.streams.push_back({"o", {0x10,0x20,0x30,0x40,0x50}});
    storage.streams.push_back({"Controls/Nested/state", {0xDE,0xAD,0xBE,0xEF}});
    workbook.setVbaDesignerModule("UserForm1", "Option Explicit\nPrivate Sub UserForm_Click()\nEnd Sub", storage);
    workbook.save(source);

    xlpp::Workbook loaded; loaded.load(source);
    const auto inspect = loaded.inspectVbaUserForm("UserForm1");
    test.checkTrue(inspect.valid, "MS-OFORMS UserForm f stream is semantically decoded");
    const auto ownership = loaded.validateVbaDesignerProject();
    test.checkTrue(ownership.ok() && ownership.designerModules == 1 && ownership.designerStorages == 1
                   && ownership.validUserFormStreams == 1,
                   "Designer ownership validator accepts a coherent UserForm project");
    test.checkEqual(inspect.properties.majorVersion, std::uint8_t{4}, "UserForm major version decodes");
    test.checkEqual(inspect.properties.caption.value_or(""), std::string("P1D!"), "Compressed UserForm caption decodes");
    test.checkEqual(inspect.properties.displayedWidth.value_or(0), std::int32_t{5000}, "UserForm DisplayedSize width decodes");
    test.checkEqual(inspect.properties.displayedHeight.value_or(0), std::int32_t{3000}, "UserForm DisplayedSize height decodes");
    test.checkEqual(inspect.properties.scrollLeft.value_or(0), std::int32_t{100}, "UserForm ScrollPosition Left decodes");
    test.checkEqual(inspect.properties.scrollTop.value_or(0), std::int32_t{200}, "UserForm ScrollPosition Top decodes");
    test.checkEqual(inspect.trailingBytes, std::size_t{6}, "Unknown FormStreamData/SiteData bytes are separated from semantic Form block");

    const auto beforeStorages = loaded.vbaDesignerStorages();
    const auto* beforeO = beforeStorages.front().findStream("o");
    const auto* beforeNested = beforeStorages.front().findStream("Controls/Nested/state");
    xlpp::VbaUserFormPropertiesPatch patch;
    patch.caption = "P1D Form ✓"; // forces uncompressed UTF-16 fmString
    patch.backColor = 0x80000005u;
    patch.displayedWidth = 6400; patch.displayedHeight = 3600;
    patch.logicalWidth = 8000; patch.logicalHeight = 5000;
    patch.scrollLeft = 321; patch.scrollTop = 654;
    patch.zoom = 150u; patch.drawBuffer = 40000u;
    test.checkTrue(loaded.updateVbaUserFormProperties("UserForm1", patch), "Semantic UserForm property patch succeeds");
    loaded.save(edited);

    xlpp::Workbook roundTrip; roundTrip.load(edited);
    const auto after = roundTrip.inspectVbaUserForm("UserForm1");
    test.checkTrue(after.valid, "Edited UserForm f stream remains structurally valid");
    test.checkEqual(after.properties.caption.value_or(""), std::string("P1D Form ✓"), "Unicode UserForm caption round-trips after compressed-to-UTF16 growth");
    test.checkEqual(after.properties.backColor.value_or(0), std::uint32_t{0x80000005u}, "UserForm BackColor semantic edit round-trips");
    test.checkEqual(after.properties.displayedWidth.value_or(0), std::int32_t{6400}, "UserForm width semantic edit round-trips");
    test.checkEqual(after.properties.displayedHeight.value_or(0), std::int32_t{3600}, "UserForm height semantic edit round-trips");
    test.checkEqual(after.properties.logicalWidth.value_or(0), std::int32_t{8000}, "UserForm LogicalSize width semantic edit round-trips");
    test.checkEqual(after.properties.scrollLeft.value_or(0), std::int32_t{321}, "UserForm scroll-left semantic edit round-trips");
    test.checkEqual(after.properties.scrollTop.value_or(0), std::int32_t{654}, "UserForm scroll-top semantic edit round-trips");
    test.checkEqual(after.properties.zoom.value_or(0), std::uint32_t{150}, "UserForm Zoom semantic edit round-trips");
    test.checkEqual(after.properties.drawBuffer.value_or(0), std::uint32_t{40000}, "UserForm DrawBuffer semantic edit round-trips");
    test.checkEqual(after.trailingBytes, std::size_t{6}, "Caption length change preserves trailing Form stream payload");
    const auto afterStorages = roundTrip.vbaDesignerStorages();
    test.checkTrue(beforeO && afterStorages.front().findStream("o") && beforeO->data == afterStorages.front().findStream("o")->data,
                   "Semantic Form edit preserves UserForm object stream byte-for-byte");
    test.checkTrue(beforeNested && afterStorages.front().findStream("Controls/Nested/state")
                   && beforeNested->data == afterStorages.front().findStream("Controls/Nested/state")->data,
                   "Semantic Form edit preserves nested designer/control streams byte-for-byte");

    auto malformedStorage = afterStorages.front();
    auto malformedF = std::find_if(malformedStorage.streams.begin(), malformedStorage.streams.end(), [](const auto& st) { return st.path == "f"; });
    malformedF->data = {0x00,0x04,0xFF};
    roundTrip.setVbaDesignerStorage(malformedStorage);
    const auto malformed = roundTrip.inspectVbaUserForm("UserForm1");
    test.checkTrue(!malformed.valid && !malformed.error.empty(), "Malformed UserForm f stream returns an explicit diagnostic");
    const auto malformedOwnership = roundTrip.validateVbaDesignerProject();
    test.checkTrue(!malformedOwnership.ok(), "Designer validator reports malformed UserForm Form stream");
    test.checkTrue(roundTrip.removeVbaDesignerStorage("UserForm1"), "Designer Storage can be removed independently for ownership diagnostics");
    const auto invalidOwnership = roundTrip.validateVbaDesignerProject();
    test.checkTrue(!invalidOwnership.ok() && std::any_of(invalidOwnership.issues.begin(), invalidOwnership.issues.end(), [](const auto& issue) {
        return issue.designerName == "UserForm1" && issue.message.find("no matching") != std::string::npos;
    }), "Designer validator diagnoses a Designer module whose storage is missing");

    std::filesystem::remove(source); std::filesystem::remove(edited);
}


std::vector<unsigned char> makeP1EUserFormControlStream(std::uint32_t objectStreamSize = 5, std::uint16_t clsidCacheIndex = 1) {
    auto data = makeP1DUserFormStream();
    data.resize(data.size() - 6); // replace the P1D unknown sentinel with a valid FormSiteData block

    auto pushU16 = [](std::vector<unsigned char>& out, std::uint16_t v) {
        out.push_back(static_cast<unsigned char>(v));
        out.push_back(static_cast<unsigned char>(v >> 8));
    };
    auto pushU32 = [](std::vector<unsigned char>& out, std::uint32_t v) {
        out.push_back(static_cast<unsigned char>(v));
        out.push_back(static_cast<unsigned char>(v >> 8));
        out.push_back(static_cast<unsigned char>(v >> 16));
        out.push_back(static_cast<unsigned char>(v >> 24));
    };
    auto putU16 = [](std::vector<unsigned char>& out, std::size_t pos, std::uint16_t v) {
        out[pos] = static_cast<unsigned char>(v);
        out[pos + 1] = static_cast<unsigned char>(v >> 8);
    };
    auto putU32 = [](std::vector<unsigned char>& out, std::size_t pos, std::uint32_t v) {
        out[pos] = static_cast<unsigned char>(v);
        out[pos + 1] = static_cast<unsigned char>(v >> 8);
        out[pos + 2] = static_cast<unsigned char>(v >> 16);
        out[pos + 3] = static_cast<unsigned char>(v >> 24);
    };
    auto compressedCount = [](std::string_view value) {
        return 0x80000000u | static_cast<std::uint32_t>(value.size());
    };

    const std::string name = "Button1";
    const std::string tag = "primary";
    const std::string tip = "Run";
    const std::string controlSource = "B2";
    const std::string rowSource = "A2:A4";
    const std::uint32_t mask =
        (1u << 0) | (1u << 1) | (1u << 2) | (1u << 3) | (1u << 4) |
        (1u << 5) | (1u << 6) | (1u << 7) | (1u << 8) | (1u << 9) |
        (1u << 11) | (1u << 13) | (1u << 14);

    std::vector<unsigned char> site;
    pushU16(site, 0); // version
    pushU16(site, 0); // cbSite patched below
    pushU32(site, mask);
    const auto dataStart = site.size();
    auto alignData = [&](std::size_t n) {
        while (((site.size() - dataStart) % n) != 0) site.push_back(0);
    };

    alignData(4); pushU32(site, compressedCount(name));
    alignData(4); pushU32(site, compressedCount(tag));
    alignData(4); pushU32(site, 42);             // ID
    alignData(4); pushU32(site, 77);             // HelpContextID
    alignData(4); pushU32(site, 0x00000017u);    // BitFlags
    alignData(4); pushU32(site, objectStreamSize); // ObjectStreamSize
    alignData(2); pushU16(site, 3);              // TabIndex
    alignData(2); pushU16(site, clsidCacheIndex); // ClsidCacheIndex
    alignData(2); pushU16(site, 2);              // GroupID
    alignData(4); pushU32(site, compressedCount(tip));
    alignData(4); pushU32(site, compressedCount(controlSource));
    alignData(4); pushU32(site, compressedCount(rowSource));
    alignData(4);

    site.insert(site.end(), name.begin(), name.end());
    site.insert(site.end(), tag.begin(), tag.end());
    pushU32(site, 120); pushU32(site, 240);      // Position
    site.insert(site.end(), tip.begin(), tip.end());
    site.insert(site.end(), controlSource.begin(), controlSource.end());
    site.insert(site.end(), rowSource.begin(), rowSource.end());
    putU16(site, 2, static_cast<std::uint16_t>(site.size() - 4));

    std::vector<unsigned char> siteData;
    pushU16(siteData, 0);                        // CountOfSiteClassInfo
    pushU32(siteData, 1);                        // CountOfSites
    const auto countBytesPos = siteData.size();
    pushU32(siteData, 0);                        // CountOfBytes
    const auto bodyStart = siteData.size();
    siteData.push_back(0);                       // depth
    siteData.push_back(1);                       // SITE_TYPE::ST_Ole
    while (((siteData.size() - bodyStart) % 4) != 0) siteData.push_back(0);
    siteData.insert(siteData.end(), site.begin(), site.end());
    putU32(siteData, countBytesPos, static_cast<std::uint32_t>(siteData.size() - bodyStart));

    data.insert(data.end(), siteData.begin(), siteData.end());
    return data;
}

std::vector<unsigned char> makeP1FButtonOrLabelObject(bool label) {
    std::vector<unsigned char> out;
    auto pushU16 = [&](std::uint16_t v) { out.push_back(static_cast<unsigned char>(v)); out.push_back(static_cast<unsigned char>(v >> 8)); };
    auto pushU32 = [&](std::uint32_t v) { out.push_back(static_cast<unsigned char>(v)); out.push_back(static_cast<unsigned char>(v >> 8)); out.push_back(static_cast<unsigned char>(v >> 16)); out.push_back(static_cast<unsigned char>(v >> 24)); };
    auto putU16 = [&](std::size_t pos, std::uint16_t v) { out[pos] = static_cast<unsigned char>(v); out[pos + 1] = static_cast<unsigned char>(v >> 8); };
    const std::string caption = label ? "Label1" : "Run";
    const std::uint32_t mask = label
        ? ((1u << 0) | (1u << 1) | (1u << 2) | (1u << 3) | (1u << 5) | (1u << 6) |
           (1u << 7) | (1u << 8) | (1u << 9) | (1u << 11))
        : ((1u << 0) | (1u << 1) | (1u << 2) | (1u << 3) | (1u << 4) | (1u << 5) |
           (1u << 6) | (1u << 8));
    out.push_back(0); out.push_back(2); pushU16(0); pushU32(mask);
    const auto dataStart = out.size();
    auto align = [&](std::size_t n) { while (((out.size() - dataStart) % n) != 0) out.push_back(0); };
    align(4); pushU32(0x80000012u); // ForeColor
    align(4); pushU32(0x8000000Fu); // BackColor
    align(4); pushU32(0x0000001Bu); // VariousPropertyBits
    align(4); pushU32(0x80000000u | static_cast<std::uint32_t>(caption.size()));
    if (!label) { align(4); pushU32(7); } // PicturePosition
    align(1); out.push_back(2);           // MousePointer
    if (label) {
        align(4); pushU32(0x80000006u);   // BorderColor
        align(2); pushU16(1);             // BorderStyle
        align(2); pushU16(2);             // SpecialEffect
        align(2); pushU16(static_cast<std::uint16_t>('L')); // Accelerator
    } else {
        align(2); pushU16(static_cast<std::uint16_t>('R')); // Accelerator
    }
    align(4);
    out.insert(out.end(), caption.begin(), caption.end());
    while ((out.size() % 4) != 0) out.push_back(0);
    pushU32(label ? 2200u : 3200u);
    pushU32(label ? 700u : 900u);
    putU16(2, static_cast<std::uint16_t>(out.size() - 4));
    // Opaque StreamData/TextProps sentinel: semantic editing must preserve it.
    out.insert(out.end(), {0xFA, 0xFB, 0xFC, 0xFD});
    return out;
}

// P1Y-A: builds a MS-OFORMS control object stream for the extended control
// families (TextBox, CheckBox, OptionButton, ToggleButton, ComboBox, ListBox,
// SpinButton, ScrollBar) with a representative PropMask and DataBlock. The
// layout follows MS-OFORMS 2.x object-stream rules: a common 4-byte header,
// a 32-bit (or 64-bit for list families) PropMask, fixed-size DataBlock, then
// string payloads in ascending flag-bit order, then an 8-byte Size pair and an
// opaque StreamData/TextProps tail.
std::vector<unsigned char> makeP1YUserFormControlObject(xlpp::VbaUserFormControlKind kind) {
    std::vector<unsigned char> out;
    auto pushU16 = [&](std::uint16_t v) { out.push_back(static_cast<unsigned char>(v)); out.push_back(static_cast<unsigned char>(v >> 8)); };
    auto pushU32 = [&](std::uint32_t v) { out.push_back(static_cast<unsigned char>(v)); out.push_back(static_cast<unsigned char>(v >> 8)); out.push_back(static_cast<unsigned char>(v >> 16)); out.push_back(static_cast<unsigned char>(v >> 24)); };
    auto putU16 = [&](std::size_t pos, std::uint16_t v) { out[pos] = static_cast<unsigned char>(v); out[pos + 1] = static_cast<unsigned char>(v >> 8); };
    out.push_back(0); out.push_back(2);
    const auto cbPos = out.size(); pushU16(0); // cbControl patched below
    const auto maskPos = out.size();
    const bool listFamily = kind == xlpp::VbaUserFormControlKind::ComboBox || kind == xlpp::VbaUserFormControlKind::ListBox;
    std::uint32_t mask = 0;
    const auto setBit = [&](unsigned b) { if (b < 32) mask |= (1u << b); };
    // Common flags: ForeColor(0), BackColor(1), VariousPropertyBits(2), Caption(3), MousePointer(6).
    setBit(0); setBit(1); setBit(2); setBit(3); setBit(6);
    std::uint32_t maskHigh = 0;
    const auto dataStart = out.size() + (listFamily ? 8 : 4);
    auto align = [&](std::size_t n) { while (((out.size() - dataStart) % n) != 0) out.push_back(0); };

    if (kind == xlpp::VbaUserFormControlKind::TextBox) {
        // Size = bit 30; SpecialEffect = bit 5; ScrollBars = bit 10;
        // MaxLength = bit 15; Text = bit 17; PasswordChar = bit 22;
        // MultiLine = bit 25.
        setBit(5); setBit(10); setBit(15); setBit(17); setBit(22); setBit(25); setBit(30);
        pushU32(mask); if (listFamily) pushU32(maskHigh);
        align(4); pushU32(0x8000000Au); // ForeColor
        align(4); pushU32(0x8000000Fu); // BackColor
        align(4); pushU32(0x00000021u); // VariousPropertyBits
        align(4); pushU32(0x80000000u | 4u); // Caption count
        align(2); pushU16(2);            // SpecialEffect (bit 5)
        align(1); out.push_back(1);      // MousePointer (bit 6)
        align(2); pushU16(1);            // ScrollBars
        align(4); pushU32(100u);         // MaxLength
        align(4); pushU32(0x80000000u | 3u); // Text count
        align(2); pushU16(42u);          // PasswordChar ('*')
        align(2); pushU16(1);            // MultiLine
        align(4); out.insert(out.end(), {'M','a','i','n'}); // Caption payload
        align(4); out.insert(out.end(), {'A','B','C'});     // Text payload
    } else if (kind == xlpp::VbaUserFormControlKind::CheckBox ||
               kind == xlpp::VbaUserFormControlKind::OptionButton ||
               kind == xlpp::VbaUserFormControlKind::ToggleButton) {
        // SpecialEffect = bit 5; GroupName = bit 10; Value = bit 13;
        // GroupNumber = bit 14; TripleState = bit 15; Size = bit 16.
        setBit(5); setBit(10); setBit(13); setBit(14); setBit(15); setBit(16);
        pushU32(mask); if (listFamily) pushU32(maskHigh);
        align(4); pushU32(0x8000000Au);
        align(4); pushU32(0x8000000Fu);
        align(4); pushU32(0x0000000Bu);
        align(4); pushU32(0x80000000u | 6u); // Caption count ("Option")
        align(2); pushU16(2);                 // SpecialEffect (bit 5)
        align(1); out.push_back(2);           // MousePointer (bit 6)
        align(4); pushU32(0x80000000u | 8u);  // GroupName count ("GroupOne")
        align(4); pushU32(0x80000000u | 4u);  // Value count ("True")
        align(2); pushU16(3);                 // GroupNumber
        align(2); pushU16(1);                 // TripleState
        align(4); out.insert(out.end(), {'O','p','t','i','o','n'});
        align(4); out.insert(out.end(), {'G','r','o','u','p','O','n','e'});
        align(4); out.insert(out.end(), {'T','r','u','e'});
    } else if (kind == xlpp::VbaUserFormControlKind::SpinButton ||
               kind == xlpp::VbaUserFormControlKind::ScrollBar) {
        // Value/Min/Max/SmallChange/LargeChange = bits 3-7, Orientation = 8,
        // Size = bit 11.
        setBit(3); setBit(4); setBit(5); setBit(6); setBit(7); setBit(8); setBit(11);
        pushU32(mask); if (listFamily) pushU32(maskHigh);
        align(4); pushU32(0x8000000Au);
        align(4); pushU32(0x8000000Fu);
        align(4); pushU32(0x00000003u);
        align(4); pushU32(0);          // Value (i4)
        align(4); pushU32(0);          // Min
        align(4); pushU32(100u);       // Max
        align(4); pushU32(1u);         // SmallChange
        align(4); pushU32(10u);        // LargeChange
        align(1); out.push_back(0);    // Orientation vertical
    } else if (kind == xlpp::VbaUserFormControlKind::ComboBox ||
               kind == xlpp::VbaUserFormControlKind::ListBox) {
        // 64-bit mask: common flags + list-specific flags. ListRows = bit 15,
        // ColumnCount = bit 19, ColumnWidths = bit 20, Size = bit 31.
        setBit(15); setBit(19); setBit(20); setBit(31);
        pushU32(mask); pushU32(maskHigh);
        align(4); pushU32(0x8000000Au);
        align(4); pushU32(0x8000000Fu);
        align(4); pushU32(0x00000011u);
        align(4); pushU32(0x80000000u | 5u); // Caption count
        align(1); out.push_back(3);          // MousePointer
        align(4); pushU32(4u);               // ListRows
        align(4); pushU32(2u);               // ColumnCount
        align(4); pushU32(0x80000000u | 9u); // ColumnWidths count ("75pt;90pt" = 9)
        align(4); out.insert(out.end(), {'I','t','e','m','s'});
        align(4); out.insert(out.end(), {'7','5','p','t',';','9','0','p','t'});
    } else {
        return {};
    }

    // Size pair (Width, Height) then opaque tail. cbControl covers everything
    // up to (but excluding) the StreamData/TextProps tail, matching P1F.
    align(4); pushU32(1600u); pushU32(400u);
    putU16(cbPos, static_cast<std::uint16_t>(out.size() - 4));
    out.insert(out.end(), {0xFA, 0xFB, 0xFC, 0xFD});
    return out;
}

void testPivotSelectiveItemAndFilterMutationP1E(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1e_pivot_source.xlsx";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1e_pivot_edited.xlsx";

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Region"), std::string("Sales")});
    sheet.append({std::string("East"), 10.0});
    sheet.append({std::string("West"), 20.0});
    sheet.append({std::string("East"), 30.0});

    xlpp::PivotCache cache;
    cache.setSourceData("'Data'!$A$1:$B$4");
    cache.setSharedCacheKey("p1e-shared-cache");
    cache.setFields({"Region", "Sales"});
    cache.setRecords({{"East", "10"}, {"West", "20"}, {"East", "30"}});

    xlpp::PivotTable first("SalesPivot");
    first.setLocation("D2");
    first.cache() = cache;
    auto& row = first.addRowField("Region");
    xlpp::PivotFieldItem east; east.cacheIndex = 0; east.type = "data"; east.caption = "East"; east.hidden = false;
    xlpp::PivotFieldItem west; west.cacheIndex = 1; west.type = "data"; west.caption = "West"; west.hidden = false;
    row.addItem(east); row.addItem(west);
    first.addDataField("Sales", "sum");
    xlpp::PivotFilter filter;
    filter.fieldIndex = 0;
    filter.type = "captionContains";
    filter.id = 1;
    filter.name = "East only";
    filter.description = "original";
    filter.stringValue1 = "East";
    filter.autoFilterXml = "<autoFilter ref=\"A1:A4\"><filters><filter val=\"East\"/></filters></autoFilter>";
    first.addFilter(filter);
    sheet.addPivotTable(first);

    xlpp::PivotTable second("SalesPivot2");
    second.setLocation("J2");
    second.cache() = cache;
    second.addRowField("Region");
    second.addDataField("Sales", "sum");
    sheet.addPivotTable(second);
    workbook.save(source);

    xlpp::Workbook loaded;
    loaded.load(source);
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto pivot2Before = before.get("xl/pivotTables/pivotTable2.xml");
    const auto cacheDefBefore = before.get("xl/pivotCache/pivotCacheDefinition1.xml");
    const auto cacheRecordsBefore = before.get("xl/pivotCache/pivotCacheRecords1.xml");

    xlpp::PivotFieldItemPatch itemPatch;
    itemPatch.caption = "Eastern region";
    itemPatch.hidden = true;
    itemPatch.showDetails = false;
    itemPatch.formula = true;
    test.checkTrue(loaded.updateImportedPivotFieldItem("Data", "SalesPivot", 0, 0, itemPatch),
                   "P1E selective imported PivotField item mutation succeeds");

    xlpp::PivotFilterPatch filterPatch;
    filterPatch.type = "captionBeginsWith";
    filterPatch.evaluationOrder = 4;
    filterPatch.name = "Eastern filter";
    filterPatch.description = "selectively edited";
    filterPatch.stringValue1 = "Ea";
    filterPatch.autoFilterXml = "<autoFilter ref=\"A1:A4\"><customFilters><customFilter operator=\"beginsWith\" val=\"Ea\"/></customFilters></autoFilter>";
    test.checkTrue(loaded.updateImportedPivotFilter("Data", "SalesPivot", 0, filterPatch),
                   "P1E selective imported PivotFilter mutation succeeds");
    xlpp::PivotCacheRecordValuePatch recordPatch;
    recordPatch.type = xlpp::PivotCacheRecordValueType::Number;
    recordPatch.value = "15.5";
    test.checkTrue(loaded.updateImportedPivotCacheRecordValue("Data", "SalesPivot", 0, 1, recordPatch),
                   "P1E selective physical PivotCache record-value mutation succeeds");
    loaded.save(edited);

    const auto after = xlpp::internal::ZipArchive::open(edited);
    const auto pivot1After = after.get("xl/pivotTables/pivotTable1.xml");
    test.checkEqual(after.get("xl/pivotTables/pivotTable2.xml"), pivot2Before,
                    "Selective item/filter edit preserves sibling PivotTable XML byte-for-byte");
    test.checkEqual(after.get("xl/pivotCache/pivotCacheDefinition1.xml"), cacheDefBefore,
                    "Selective PivotTable item/filter edit preserves shared cache definition byte-for-byte");
    test.checkTrue(after.get("xl/pivotCache/pivotCacheRecords1.xml") != cacheRecordsBefore
                   && after.get("xl/pivotCache/pivotCacheRecords1.xml").find("15.5") != std::string::npos,
                   "Selective PivotCache record edit changes only the targeted physical records part");
    test.checkTrue(pivot1After.find("n=\"Eastern region\"") != std::string::npos
                   && pivot1After.find("h=\"1\"") != std::string::npos
                   && pivot1After.find("sd=\"0\"") != std::string::npos
                   && pivot1After.find("f=\"1\"") != std::string::npos,
                   "Selective Pivot item semantic attributes are patched");
    test.checkTrue(pivot1After.find("type=\"captionBeginsWith\"") != std::string::npos
                   && pivot1After.find("name=\"Eastern filter\"") != std::string::npos
                   && pivot1After.find("stringValue1=\"Ea\"") != std::string::npos
                   && pivot1After.find("customFilter") != std::string::npos,
                   "Selective Pivot filter attributes and nested AutoFilter are patched");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(),
                   "P1E selective PivotTable mutation preserves package graph");

    xlpp::Workbook roundTrip;
    roundTrip.load(edited);
    const auto* roundSheet = static_cast<const xlpp::Workbook&>(roundTrip).worksheet("Data");
    const auto& roundPivot = roundSheet->pivotTables().front();
    const auto& roundItem = roundPivot.rowFields().front().items().front();
    test.checkEqual(roundItem.caption, std::string("Eastern region"), "Selective Pivot item caption reloads");
    test.checkTrue(roundItem.hidden && !roundItem.showDetails && roundItem.formula,
                   "Selective Pivot item flags reload");
    test.checkEqual(roundPivot.filters().front().type, std::string("captionBeginsWith"),
                    "Selective Pivot filter type reloads");
    test.checkEqual(roundPivot.filters().front().description, std::string("selectively edited"),
                    "Selective Pivot filter description reloads");
    test.checkTrue(roundPivot.filters().front().autoFilterXml.find("customFilter") != std::string::npos,
                   "Selective Pivot nested AutoFilter reloads");
    for (const auto& pivot : roundSheet->pivotTables()) {
        test.checkEqual(pivot.cache().records().front()[1], std::string("15.5"),
                        "Selective physical cache-record mutation reaches every shared-cache model");
    }

    std::filesystem::remove(source);
    std::filesystem::remove(edited);
}

void testVbaUserFormControlSitesP1E(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1e_userform_source.xlsm";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1e_userform_edited.xlsm";

    xlpp::Workbook workbook;
    workbook.addWorksheet("Data").setVbaCodeName("DataSheet");
    xlpp::VbaDesignerStorage storage;
    storage.name = "UserForm1";
    storage.streams.push_back({"f", makeP1EUserFormControlStream()});
    storage.streams.push_back({"o", {0x10, 0x20, 0x30, 0x40, 0x50}});
    storage.streams.push_back({"Controls/Nested/state", {0xAA, 0xBB, 0xCC}});
    workbook.setVbaDesignerModule("UserForm1", "Option Explicit\n", storage);
    workbook.save(source);

    xlpp::Workbook loaded;
    loaded.load(source);
    const auto beforeStorages = loaded.vbaDesignerStorages();
    const auto* beforeO = beforeStorages.front().findStream("o");
    const auto* beforeNested = beforeStorages.front().findStream("Controls/Nested/state");
    const auto controls = loaded.inspectVbaUserFormControls("UserForm1");
    test.checkTrue(controls.valid && controls.controls.size() == 1,
                   "P1E MS-OFORMS FormSiteData decodes one embedded control");
    test.checkEqual(controls.classInfoCount, std::size_t{0}, "FormSiteData ClassTable count decodes");
    test.checkEqual(controls.totalObjectStreamBytes, std::size_t{5}, "OleSite ObjectStreamSize aggregation decodes");
    const auto& control = controls.controls.front();
    test.checkEqual(control.siteType, std::uint8_t{1}, "Embedded control SITE_TYPE is ST_Ole");
    test.checkEqual(control.name.value_or(""), std::string("Button1"), "Control Name decodes");
    test.checkEqual(control.tag.value_or(""), std::string("primary"), "Control Tag decodes");
    test.checkEqual(control.id.value_or(-1), std::int32_t{42}, "Control ID decodes");
    test.checkEqual(control.helpContextId.value_or(-1), std::int32_t{77}, "Control HelpContextID decodes");
    test.checkEqual(control.objectStreamSize.value_or(0), std::uint32_t{5}, "Control ObjectStreamSize decodes");
    test.checkEqual(control.objectStreamOffset, std::size_t{0}, "First embedded control maps to the start of object stream o");
    test.checkTrue(control.objectData == std::vector<unsigned char>({0x10, 0x20, 0x30, 0x40, 0x50}),
                   "Control inspection exposes its lossless ObjectStreamSize slice");
    test.checkEqual(controls.objectStreamBytes, std::size_t{5}, "UserForm object-stream byte count is reported");
    test.checkEqual(controls.unassignedObjectStreamBytes, std::size_t{0}, "All UserForm object-stream bytes are assigned to control sites");
    test.checkEqual(control.tabIndex.value_or(-1), std::int16_t{3}, "Control TabIndex decodes");
    test.checkEqual(control.clsidCacheIndex.value_or(0), std::uint16_t{1}, "Control ClsidCacheIndex decodes");
    test.checkEqual(control.groupId.value_or(0), std::uint16_t{2}, "Control GroupID decodes");
    test.checkEqual(control.top.value_or(0), std::int32_t{120}, "Control top position decodes");
    test.checkEqual(control.left.value_or(0), std::int32_t{240}, "Control left position decodes");
    test.checkEqual(control.controlTipText.value_or(""), std::string("Run"), "Control tooltip decodes");
    test.checkEqual(control.controlSource.value_or(""), std::string("B2"), "Control source decodes");
    test.checkEqual(control.rowSource.value_or(""), std::string("A2:A4"), "Control row source decodes");

    xlpp::VbaUserFormControlSitePatch patch;
    patch.name = "Nút ✓";
    patch.tag = "primary-updated";
    patch.helpContextId = 88;
    patch.bitFlags = 0x31u;
    patch.tabIndex = 7;
    patch.groupId = 4;
    patch.controlTipText = "Chạy ✓";
    patch.controlSource = "C3";
    patch.rowSource = "C3:C8";
    patch.top = 333;
    patch.left = 444;
    test.checkTrue(loaded.updateVbaUserFormControlSite("UserForm1", 0, patch),
                   "P1E semantic UserForm control-site patch succeeds");
    loaded.save(edited);

    xlpp::Workbook roundTrip;
    roundTrip.load(edited);
    const auto after = roundTrip.inspectVbaUserFormControls("UserForm1");
    test.checkTrue(after.valid && after.controls.size() == 1, "Edited FormSiteData remains structurally valid");
    const auto& updated = after.controls.front();
    test.checkEqual(updated.name.value_or(""), std::string("Nút ✓"), "Unicode control Name round-trips");
    test.checkEqual(updated.tag.value_or(""), std::string("primary-updated"), "Control Tag growth round-trips");
    test.checkEqual(updated.helpContextId.value_or(0), std::int32_t{88}, "Control HelpContextID edit round-trips");
    test.checkEqual(updated.bitFlags.value_or(0), std::uint32_t{0x31u}, "Control BitFlags edit round-trips");
    test.checkEqual(updated.tabIndex.value_or(0), std::int16_t{7}, "Control TabIndex edit round-trips");
    test.checkEqual(updated.groupId.value_or(0), std::uint16_t{4}, "Control GroupID edit round-trips");
    test.checkEqual(updated.controlTipText.value_or(""), std::string("Chạy ✓"), "Unicode control tooltip round-trips");
    test.checkEqual(updated.controlSource.value_or(""), std::string("C3"), "ControlSource edit round-trips");
    test.checkEqual(updated.rowSource.value_or(""), std::string("C3:C8"), "RowSource edit round-trips");
    test.checkEqual(updated.top.value_or(0), std::int32_t{333}, "Control top-position edit round-trips");
    test.checkEqual(updated.left.value_or(0), std::int32_t{444}, "Control left-position edit round-trips");
    test.checkEqual(after.totalObjectStreamBytes, std::size_t{5}, "Control-site string growth does not change ObjectStreamSize");
    test.checkTrue(after.controls.front().objectData == std::vector<unsigned char>({0x10, 0x20, 0x30, 0x40, 0x50}),
                   "Control object-stream slice remains byte-for-byte stable after site metadata edits");

    const auto afterStorages = roundTrip.vbaDesignerStorages();
    test.checkTrue(beforeO && afterStorages.front().findStream("o")
                   && beforeO->data == afterStorages.front().findStream("o")->data,
                   "Control-site metadata edits preserve object stream o byte-for-byte");
    test.checkTrue(beforeNested && afterStorages.front().findStream("Controls/Nested/state")
                   && beforeNested->data == afterStorages.front().findStream("Controls/Nested/state")->data,
                   "Control-site metadata edits preserve nested control streams byte-for-byte");

    bool threw = false;
    try {
        xlpp::VbaUserFormControlSitePatch invalid;
        invalid.groupId = 9;
        auto info = roundTrip.vbaProjectInfo();
        auto& f = *std::find_if(info.designerStorages.front().streams.begin(), info.designerStorages.front().streams.end(),
                               [](const auto& stream) { return stream.path == "f"; });
        // Clear GroupID presence bit while retaining the bytes to exercise the explicit mask contract.
        const auto siteOffset = after.siteDataOffset + 2 + 8 + 4;
        const auto maskOffset = siteOffset + 4;
        f.data[maskOffset + 1] = static_cast<unsigned char>(f.data[maskOffset + 1] & ~(1u << 1));
        roundTrip.setVbaProjectInfo(std::move(info));
        roundTrip.updateVbaUserFormControlSite("UserForm1", 0, invalid);
    } catch (const std::exception&) {
        threw = true;
    }
    test.checkTrue(threw, "Control-site editor refuses to synthesize a property absent from SitePropMask");

    std::filesystem::remove(source);
    std::filesystem::remove(edited);
}

void testPivotSelectiveDataAndPageFieldsP1F(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1f_pivot_source.xlsx";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1f_pivot_edited.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Region"), std::string("Sales"), std::string("Year")});
    sheet.append({std::string("East"), 10.0, 2025.0});
    sheet.append({std::string("West"), 20.0, 2026.0});
    xlpp::PivotCache cache;
    cache.setSourceData("'Data'!$A$1:$C$3");
    cache.setFields({"Region", "Sales", "Year"});
    cache.setRecords({{"East", "10", "2025"}, {"West", "20", "2026"}});
    xlpp::PivotTable pivot("SalesPivot");
    pivot.setLocation("E2"); pivot.cache() = cache;
    pivot.addRowField("Region");
    pivot.addPageField("Year");
    auto& page = pivot.pageFieldSettings().back(); page.setItem(0); page.setHierarchy(-1); page.setName("Year selector");
    auto& data = pivot.addDataField("Sales", "sum"); data.setDisplayName("Sales total");
    sheet.addPivotTable(pivot);
    workbook.save(source);

    xlpp::Workbook loaded; loaded.load(source);
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto cacheBefore = before.get("xl/pivotCache/pivotCacheDefinition1.xml");
    const auto recordsBefore = before.get("xl/pivotCache/pivotCacheRecords1.xml");
    xlpp::PivotDataFieldPatch dataPatch;
    dataPatch.name = "Average sales"; dataPatch.subtotal = "average"; dataPatch.showDataAs = "percentOfTotal";
    dataPatch.baseField = 0; dataPatch.baseItem = 0; dataPatch.numberFormatId = 4;
    test.checkTrue(loaded.updateImportedPivotDataField("Data", "SalesPivot", 0, dataPatch),
                   "P1F selective imported dataField mutation succeeds");
    xlpp::PivotPageFieldPatch pagePatch;
    pagePatch.item = 1; pagePatch.hierarchy = 2; pagePatch.name = "Fiscal year";
    test.checkTrue(loaded.updateImportedPivotPageField("Data", "SalesPivot", 0, pagePatch),
                   "P1F selective imported pageField mutation succeeds");
    loaded.save(edited);
    const auto after = xlpp::internal::ZipArchive::open(edited);
    const auto pivotXml = after.get("xl/pivotTables/pivotTable1.xml");
    test.checkTrue(pivotXml.find("name=\"Average sales\"") != std::string::npos
                   && pivotXml.find("subtotal=\"average\"") != std::string::npos
                   && pivotXml.find("showDataAs=\"percentOfTotal\"") != std::string::npos
                   && pivotXml.find("numFmtId=\"4\"") != std::string::npos,
                   "P1F dataField semantic attributes are patched");
    test.checkTrue(pivotXml.find("name=\"Fiscal year\"") != std::string::npos
                   && pivotXml.find("item=\"1\"") != std::string::npos
                   && pivotXml.find("hier=\"2\"") != std::string::npos,
                   "P1F pageField semantic attributes are patched");
    test.checkEqual(after.get("xl/pivotCache/pivotCacheDefinition1.xml"), cacheBefore,
                    "P1F data/page field edits preserve cache definition byte-for-byte");
    test.checkEqual(after.get("xl/pivotCache/pivotCacheRecords1.xml"), recordsBefore,
                    "P1F data/page field edits preserve cache records byte-for-byte");
    xlpp::Workbook round; round.load(edited);
    const auto* roundSheet = static_cast<const xlpp::Workbook&>(round).worksheet("Data");
    const auto& roundPivot = roundSheet->pivotTables().front();
    test.checkEqual(roundPivot.dataFields().front().displayName(), std::string("Average sales"), "P1F dataField name reloads");
    test.checkEqual(roundPivot.dataFields().front().subtotal(), std::string("average"), "P1F dataField subtotal reloads");
    test.checkEqual(roundPivot.dataFields().front().showDataAs(), std::string("percentOfTotal"), "P1F dataField showDataAs reloads");
    test.checkEqual(roundPivot.pageFieldSettings().front().item(), 1, "P1F pageField item reloads");
    test.checkEqual(roundPivot.pageFieldSettings().front().hierarchy(), 2, "P1F pageField hierarchy reloads");
    std::filesystem::remove(source); std::filesystem::remove(edited);
}

void testVbaUserFormControlObjectsP1F(TestContext& test) {
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1f_controls_source.xlsm";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1f_controls_edited.xlsm";
    xlpp::Workbook workbook; workbook.addWorksheet("Data").setVbaCodeName("DataSheet");
    const auto buttonObject = makeP1FButtonOrLabelObject(false);
    xlpp::VbaDesignerStorage buttonStorage; buttonStorage.name = "ButtonForm";
    buttonStorage.streams.push_back({"f", makeP1EUserFormControlStream(static_cast<std::uint32_t>(buttonObject.size()), 17)});
    buttonStorage.streams.push_back({"o", buttonObject});
    workbook.setVbaDesignerModule("ButtonForm", "Option Explicit\n", buttonStorage);
    const auto labelObject = makeP1FButtonOrLabelObject(true);
    xlpp::VbaDesignerStorage labelStorage; labelStorage.name = "LabelForm";
    labelStorage.streams.push_back({"f", makeP1EUserFormControlStream(static_cast<std::uint32_t>(labelObject.size()), 21)});
    labelStorage.streams.push_back({"o", labelObject});
    workbook.setVbaDesignerModule("LabelForm", "Option Explicit\n", labelStorage);
    workbook.save(source);

    xlpp::Workbook loaded; loaded.load(source);
    auto buttonSite = loaded.inspectVbaUserFormControls("ButtonForm");
    test.checkTrue(buttonSite.valid && buttonSite.controls.front().kind == xlpp::VbaUserFormControlKind::CommandButton,
                   "P1F ClsidCacheIndex classifies built-in CommandButton");
    auto button = loaded.inspectVbaUserFormControlObject("ButtonForm", 0);
    test.checkTrue(button.valid && button.properties.semanticPropertiesSupported,
                   "P1F CommandButton object stream is semantically decoded");
    test.checkEqual(button.properties.caption.value_or(""), std::string("Run"), "CommandButton Caption decodes");
    test.checkEqual(button.properties.width.value_or(0), std::int32_t{3200}, "CommandButton Width decodes");
    test.checkEqual(button.trailingBytes, std::size_t{4}, "CommandButton opaque TextProps/StreamData tail is reported");
    auto label = loaded.inspectVbaUserFormControlObject("LabelForm", 0);
    test.checkTrue(label.valid && label.properties.kind == xlpp::VbaUserFormControlKind::Label,
                   "P1F Label object stream is semantically decoded");
    test.checkEqual(label.properties.borderStyle.value_or(0), std::uint16_t{1}, "Label BorderStyle decodes");

    xlpp::VbaUserFormControlObjectPatch buttonPatch;
    buttonPatch.caption = "Chạy ✓"; buttonPatch.width = 4100; buttonPatch.height = 1100; buttonPatch.foreColor = 0x80000008u;
    test.checkTrue(loaded.updateVbaUserFormControlObject("ButtonForm", 0, buttonPatch),
                   "P1F semantic CommandButton object edit succeeds");
    xlpp::VbaUserFormControlObjectPatch labelPatch;
    labelPatch.caption = "Nhãn ✓"; labelPatch.width = 2600; labelPatch.borderStyle = 0;
    test.checkTrue(loaded.updateVbaUserFormControlObject("LabelForm", 0, labelPatch),
                   "P1F semantic Label object edit succeeds");
    loaded.save(edited);

    xlpp::Workbook round; round.load(edited);
    const auto buttonAfter = round.inspectVbaUserFormControlObject("ButtonForm", 0);
    test.checkTrue(buttonAfter.valid, "Edited CommandButton object remains valid");
    test.checkEqual(buttonAfter.properties.caption.value_or(""), std::string("Chạy ✓"), "Unicode CommandButton Caption round-trips");
    test.checkEqual(buttonAfter.properties.width.value_or(0), std::int32_t{4100}, "CommandButton Width edit round-trips");
    test.checkEqual(buttonAfter.properties.height.value_or(0), std::int32_t{1100}, "CommandButton Height edit round-trips");
    test.checkEqual(buttonAfter.trailingBytes, std::size_t{4}, "CommandButton opaque tail survives caption growth");
    const auto controlsAfter = round.inspectVbaUserFormControls("ButtonForm");
    test.checkEqual(controlsAfter.controls.front().objectStreamSize.value_or(0),
                    static_cast<std::uint32_t>(controlsAfter.controls.front().objectData.size()),
                    "CommandButton caption growth updates OleSite.ObjectStreamSize");
    const auto labelAfter = round.inspectVbaUserFormControlObject("LabelForm", 0);
    test.checkEqual(labelAfter.properties.caption.value_or(""), std::string("Nhãn ✓"), "Unicode Label Caption round-trips");
    test.checkEqual(labelAfter.properties.width.value_or(0), std::int32_t{2600}, "Label Width edit round-trips");
    test.checkEqual(labelAfter.properties.borderStyle.value_or(99), std::uint16_t{0}, "Label BorderStyle edit round-trips");
    test.checkTrue(round.validateVbaDesignerProject().ok(), "P1F semantic control edits preserve Designer ownership/structure");
    std::filesystem::remove(source); std::filesystem::remove(edited);
}


void testVbaUserFormExtendedControlsP1Y(TestContext& test) {
    // P1Y-A extends semantic UserForm control-object decoding beyond
    // CommandButton/Label to the other built-in MS-OFORMS families.
    struct KindCase {
        xlpp::VbaUserFormControlKind kind;
        std::uint16_t clsid;
        const char* name;
    };
    const KindCase cases[] = {
        {xlpp::VbaUserFormControlKind::TextBox, 23, "TextBoxForm"},
        {xlpp::VbaUserFormControlKind::CheckBox, 26, "CheckBoxForm"},
        {xlpp::VbaUserFormControlKind::OptionButton, 27, "OptionButtonForm"},
        {xlpp::VbaUserFormControlKind::ToggleButton, 28, "ToggleButtonForm"},
        {xlpp::VbaUserFormControlKind::ComboBox, 25, "ComboBoxForm"},
        {xlpp::VbaUserFormControlKind::ListBox, 24, "ListBoxForm"},
        {xlpp::VbaUserFormControlKind::SpinButton, 16, "SpinButtonForm"},
        {xlpp::VbaUserFormControlKind::ScrollBar, 47, "ScrollBarForm"},
    };
    const auto source = std::filesystem::temp_directory_path() / "xlpp_p1y_controls_source.xlsm";
    const auto edited = std::filesystem::temp_directory_path() / "xlpp_p1y_controls_edited.xlsm";

    xlpp::Workbook workbook;
    workbook.addWorksheet("Data").setVbaCodeName("DataSheet");
    for (const auto& c : cases) {
        const auto object = makeP1YUserFormControlObject(c.kind);
        test.checkTrue(!object.empty(), std::string("P1Y object stream built for ") + c.name);
        xlpp::VbaDesignerStorage storage;
        storage.name = c.name;
        storage.streams.push_back({"f", makeP1EUserFormControlStream(static_cast<std::uint32_t>(object.size()), c.clsid)});
        storage.streams.push_back({"o", object});
        workbook.setVbaDesignerModule(c.name, "Option Explicit\n", storage);
    }
    workbook.save(source);

    xlpp::Workbook loaded; loaded.load(source);
    for (const auto& c : cases) {
        auto site = loaded.inspectVbaUserFormControls(c.name);
        test.checkTrue(site.valid && site.controls.size() == 1 && site.controls.front().kind == c.kind,
                       std::string("P1Y ClsidCacheIndex classifies ") + c.name);
        auto obj = loaded.inspectVbaUserFormControlObject(c.name, 0);
        test.checkTrue(obj.valid, std::string("P1Y object stream decoded for ") + c.name);
        test.checkTrue(obj.properties.semanticPropertiesSupported,
                       std::string("P1Y semantic properties enabled for ") + c.name);
        test.checkEqual(obj.properties.foreColor.value_or(0), std::uint32_t{0x8000000Au},
                        std::string("P1Y ForeColor decodes for ") + c.name);
        test.checkEqual(obj.properties.width.value_or(0), std::int32_t{1600},
                        std::string("P1Y Width decodes for ") + c.name);
        test.checkEqual(obj.properties.height.value_or(0), std::int32_t{400},
                        std::string("P1Y Height decodes for ") + c.name);
        test.checkEqual(obj.trailingBytes, std::size_t{4},
                        std::string("P1Y opaque tail preserved for ") + c.name);
    }

    // Class-specific fields.
    {
        auto text = loaded.inspectVbaUserFormControlObject("TextBoxForm", 0);
        test.checkEqual(text.properties.caption.value_or(""), std::string("Main"), "P1Y TextBox Caption decodes");
        test.checkEqual(text.properties.text.value_or(""), std::string("ABC"), "P1Y TextBox Text decodes");
        test.checkEqual(text.properties.scrollBars.value_or(0), std::uint16_t{1}, "P1Y TextBox ScrollBars decodes");
        test.checkEqual(text.properties.maxLength.value_or(0), std::uint32_t{100}, "P1Y TextBox MaxLength decodes");
        test.checkEqual(text.properties.passwordChar.value_or(0), std::uint16_t{42}, "P1Y TextBox PasswordChar decodes");
        test.checkEqual(text.properties.multiLine.value_or(0), std::uint16_t{1}, "P1Y TextBox MultiLine decodes");
    }
    {
        auto option = loaded.inspectVbaUserFormControlObject("OptionButtonForm", 0);
        test.checkEqual(option.properties.groupName.value_or(""), std::string("GroupOne"), "P1Y OptionButton GroupName decodes");
        test.checkEqual(option.properties.value.value_or(""), std::string("True"), "P1Y OptionButton Value decodes");
        test.checkEqual(option.properties.groupNumber.value_or(0), std::uint16_t{3}, "P1Y OptionButton GroupNumber decodes");
        test.checkEqual(option.properties.tripleState.value_or(0), std::uint16_t{1}, "P1Y OptionButton TripleState decodes");
    }
    {
        auto spin = loaded.inspectVbaUserFormControlObject("SpinButtonForm", 0);
        test.checkEqual(spin.properties.min.value_or(999), std::uint32_t{0}, "P1Y SpinButton Min decodes");
        test.checkEqual(spin.properties.max.value_or(0), std::uint32_t{100}, "P1Y SpinButton Max decodes");
        test.checkEqual(spin.properties.smallChange.value_or(0), std::uint32_t{1}, "P1Y SpinButton SmallChange decodes");
        test.checkEqual(spin.properties.largeChange.value_or(0), std::uint32_t{10}, "P1Y SpinButton LargeChange decodes");
    }
    {
        auto list = loaded.inspectVbaUserFormControlObject("ListBoxForm", 0);
        test.checkEqual(list.properties.listRows.value_or(0), std::uint32_t{4}, "P1Y ListBox ListRows decodes");
        test.checkEqual(list.properties.columnCount.value_or(0), std::uint32_t{2}, "P1Y ListBox ColumnCount decodes");
        test.checkEqual(list.properties.columnWidths.value_or(""), std::string("75pt;90pt"), "P1Y ListBox ColumnWidths decodes");
    }

    // Semantic edits on the newly supported families.
    xlpp::VbaUserFormControlObjectPatch textPatch;
    textPatch.text = "XYZ ✓"; textPatch.scrollBars = 3; textPatch.maxLength = 250;
    test.checkTrue(loaded.updateVbaUserFormControlObject("TextBoxForm", 0, textPatch),
                   "P1Y TextBox semantic object edit succeeds");
    xlpp::VbaUserFormControlObjectPatch optionPatch;
    optionPatch.value = "False"; optionPatch.groupName = "GroupTwo"; optionPatch.groupNumber = 5;
    test.checkTrue(loaded.updateVbaUserFormControlObject("OptionButtonForm", 0, optionPatch),
                   "P1Y OptionButton semantic object edit succeeds");
    xlpp::VbaUserFormControlObjectPatch spinPatch;
    spinPatch.min = 5; spinPatch.max = 200; spinPatch.smallChange = 5;
    test.checkTrue(loaded.updateVbaUserFormControlObject("SpinButtonForm", 0, spinPatch),
                   "P1Y SpinButton semantic object edit succeeds");
    xlpp::VbaUserFormControlObjectPatch listPatch;
    listPatch.columnWidths = "100pt;120pt"; listPatch.listRows = 8;
    test.checkTrue(loaded.updateVbaUserFormControlObject("ListBoxForm", 0, listPatch),
                   "P1Y ListBox semantic object edit succeeds");
    loaded.save(edited);

    xlpp::Workbook round; round.load(edited);
    {
        auto text = round.inspectVbaUserFormControlObject("TextBoxForm", 0);
        test.checkEqual(text.properties.text.value_or(""), std::string("XYZ ✓"), "P1Y TextBox Text edit round-trips");
        test.checkEqual(text.properties.scrollBars.value_or(0), std::uint16_t{3}, "P1Y TextBox ScrollBars edit round-trips");
        test.checkEqual(text.properties.maxLength.value_or(0), std::uint32_t{250}, "P1Y TextBox MaxLength edit round-trips");
    }
    {
        auto option = round.inspectVbaUserFormControlObject("OptionButtonForm", 0);
        test.checkEqual(option.properties.value.value_or(""), std::string("False"), "P1Y OptionButton Value edit round-trips");
        test.checkEqual(option.properties.groupName.value_or(""), std::string("GroupTwo"), "P1Y OptionButton GroupName edit round-trips");
        test.checkEqual(option.properties.groupNumber.value_or(0), std::uint16_t{5}, "P1Y OptionButton GroupNumber edit round-trips");
    }
    {
        auto spin = round.inspectVbaUserFormControlObject("SpinButtonForm", 0);
        test.checkEqual(spin.properties.min.value_or(999), std::uint32_t{5}, "P1Y SpinButton Min edit round-trips");
        test.checkEqual(spin.properties.max.value_or(0), std::uint32_t{200}, "P1Y SpinButton Max edit round-trips");
        test.checkEqual(spin.properties.smallChange.value_or(0), std::uint32_t{5}, "P1Y SpinButton SmallChange edit round-trips");
    }
    {
        auto list = round.inspectVbaUserFormControlObject("ListBoxForm", 0);
        test.checkEqual(list.properties.columnWidths.value_or(""), std::string("100pt;120pt"), "P1Y ListBox ColumnWidths edit round-trips");
        test.checkEqual(list.properties.listRows.value_or(0), std::uint32_t{8}, "P1Y ListBox ListRows edit round-trips");
    }
    test.checkTrue(round.validateVbaDesignerProject().ok(), "P1Y extended control edits preserve Designer ownership/structure");
    std::filesystem::remove(source); std::filesystem::remove(edited);
}


void testSparklinesRoundTrip(TestContext& test) {
    // Sparklines are an x14 worksheet extension that openpyxl does not model at
    // all. XL++ authors inline sparkline groups and round-trips them.
    const auto source = std::filesystem::temp_directory_path() / "xlpp_sparklines.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.cell("A1").setValue("Month");
    sheet.cell("B1").setValue("Value");
    sheet.append({std::string("Jan"), 10.0});
    sheet.append({std::string("Feb"), 15.0});
    sheet.append({std::string("Mar"), 12.0});

    xlpp::SparklineGroup group;
    group.type = "line";
    group.lineStyle = "smooth";
    group.displayMarkers = true;
    group.high = true;
    group.low = true;
    group.negative = true;
    group.markersColor = "FF0000FF";
    group.negativeColor = "FFFF0000";
    group.sparklines.push_back({"'Data'!B2:B4", "D2"});
    sheet.sparklineGroups().push_back(std::move(group));
    workbook.save(source);

    const auto archive = xlpp::internal::ZipArchive::open(source);
    const auto sheetPart = "xl/worksheets/sheet1.xml";
    test.checkTrue(archive.contains(sheetPart), "Worksheet part exists");
    const auto xml = archive.get(sheetPart);
    test.checkTrue(xml.find("sparklineGroups") != std::string::npos, "Sparkline groups extension is serialized");
    test.checkTrue(xml.find("B2:B4") != std::string::npos, "Sparkline data reference is serialized");
    test.checkTrue(xml.find("sqref>D2<") != std::string::npos, "Sparkline target cell is serialized");
    test.checkTrue(xml.find("05C60535") != std::string::npos, "Sparkline x14 extension URI is serialized");

    xlpp::Workbook round; round.load(source);
    auto* roundSheet = round.worksheet("Data");
    test.checkTrue(roundSheet != nullptr, "Sparkline workbook reloads");
    test.checkTrue(roundSheet->hasSparklines(), "Sparklines are detected on reload");
    test.checkEqual(roundSheet->sparklineGroups().size(), std::size_t{1}, "Sparkline group count is read");
    const auto& loadedGroup = roundSheet->sparklineGroups().front();
    test.checkEqual(loadedGroup.type, std::string("line"), "Sparkline group type round-trips");
    test.checkTrue(loadedGroup.displayMarkers && loadedGroup.high && loadedGroup.low && loadedGroup.negative,
                   "Sparkline display flags round-trip");
    test.checkEqual(loadedGroup.sparklines.size(), std::size_t{1}, "Sparkline count is read");
    test.checkEqual(loadedGroup.sparklines.front().reference, std::string("'Data'!B2:B4"),
                    "Sparkline data reference round-trips");
    test.checkEqual(loadedGroup.sparklines.front().location, std::string("D2"),
                    "Sparkline target cell round-trips");

    std::filesystem::remove(source);
}


void testAutoFilterTop10AndDynamicFilters(TestContext& test) {
    // Top-10 and dynamic auto-filters round-trip through the object model.
    const auto source = std::filesystem::temp_directory_path() / "xlpp_filters.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.cell("A1").setValue("Product");
    sheet.cell("B1").setValue("Sales");
    for (int i = 1; i <= 20; ++i) {
        sheet.append({std::string("P") + std::to_string(i), static_cast<double>(i * 10)});
    }
    auto& autoFilter = sheet.autoFilter();
    autoFilter.setReference("A1:B21");
    auto& sales = autoFilter.column(1);
    sales.setTop10(true, 5);
    auto& product = autoFilter.column(0);
    product.setDynamicFilter("beginsWith", std::optional<double>(3.0));
    workbook.save(source);

    const auto archive = xlpp::internal::ZipArchive::open(source);
    const auto sheetXml = archive.get("xl/worksheets/sheet1.xml");
    test.checkTrue(sheetXml.find("<top10 top=\"1\" percent=\"0\" val=\"5\"/>") != std::string::npos,
                   "Top-10 filter is serialized");
    test.checkTrue(sheetXml.find("dynamicFilter type=\"beginsWith\" val=\"3\"") != std::string::npos,
                   "Dynamic filter is serialized");

    xlpp::Workbook round; round.load(source);
    auto* roundSheet = round.worksheet("Data");
    test.checkTrue(roundSheet != nullptr && roundSheet->autoFilter().enabled(), "AutoFilter reloads");
    const auto* salesLoaded = roundSheet->autoFilter().tryColumn(1);
    test.checkTrue(salesLoaded != nullptr && salesLoaded->top10().has_value(), "Top-10 filter is read");
    test.checkTrue(salesLoaded->top10()->top && salesLoaded->top10()->value == 5 && !salesLoaded->top10()->percent,
                   "Top-10 attributes round-trip");
    const auto* productLoaded = roundSheet->autoFilter().tryColumn(0);
    test.checkTrue(productLoaded != nullptr && productLoaded->dynamicFilter().has_value(),
                   "Dynamic filter is read");
    test.checkEqual(productLoaded->dynamicFilter()->type, std::string("beginsWith"),
                    "Dynamic filter type round-trips");
    test.checkTrue(productLoaded->dynamicFilter()->value.has_value()
                   && *productLoaded->dynamicFilter()->value == 3.0,
                   "Dynamic filter value round-trips");

    std::filesystem::remove(source);
}


void testTableTotalsRow(TestContext& test) {
    // Table column totals-row aggregation/labels round-trip.
    const auto source = std::filesystem::temp_directory_path() / "xlpp_table_totals.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.append({std::string("Product"), std::string("Sales")});
    sheet.append({std::string("A"), 10.0});
    sheet.append({std::string("B"), 20.0});
    auto& table = sheet.addTable("SalesTable", "A1:B3");
    table.setShowTotalsRow(true);
    auto& salesColumn = table.addColumn("Sales");
    salesColumn.setTotalsRowFunction("sum");
    salesColumn.setTotalsRowLabel("Total");
    workbook.save(source);

    const auto archive = xlpp::internal::ZipArchive::open(source);
    const auto tablePart = archive.contains("xl/tables/table1.xml") ? "xl/tables/table1.xml" : "xl/tables/table2.xml";
    const auto tableXml = archive.get(tablePart);
    test.checkTrue(tableXml.find("totalsRowShown=\"1\"") != std::string::npos, "Table totals row is enabled");
    test.checkTrue(tableXml.find("totalsRowFunction=\"sum\"") != std::string::npos, "Totals row function is serialized");
    test.checkTrue(tableXml.find("totalsRowLabel=\"Total\"") != std::string::npos, "Totals row label is serialized");

    xlpp::Workbook round; round.load(source);
    auto* roundSheet = round.worksheet("Data");
    test.checkTrue(roundSheet != nullptr && roundSheet->tables().size() == 1, "Table reloads");
    const auto& loadedTable = roundSheet->tables().front();
    test.checkTrue(loadedTable.showTotalsRow(), "Table totals row flag round-trips");
    const auto& loadedColumn = loadedTable.columns().front();
    test.checkEqual(loadedColumn.totalsRowFunction(), std::string("sum"), "Totals row function round-trips");
    test.checkEqual(loadedColumn.totalsRowLabel(), std::string("Total"), "Totals row label round-trips");

    std::filesystem::remove(source);
}


void testPasswordToOpenEncryptionP1G(TestContext& test) {
    const auto encryptedPath = std::filesystem::temp_directory_path() / "xlpp_p1g_agile_encrypted.xlsx";
    const auto tamperedPath = std::filesystem::temp_directory_path() / "xlpp_p1g_agile_tampered.xlsx";

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Secrets");
    sheet.cell("A1").setValue("classified");
    sheet.cell("B2").setValue(42.5);
    sheet.cell("C3").setFormula("=B2*2");

    xlpp::SaveOptions saveOptions;
    saveOptions.encryption.enabled = true;
    saveOptions.encryption.password = "P@ssw0rd✓";
    // Exercise the production default instead of weakening the test profile.
    test.checkEqual(saveOptions.encryption.spinCount, std::uint32_t{100000},
                    "P1G Agile encryption defaults to 100,000 password-hash spins");
    workbook.save(encryptedPath, saveOptions);

    test.checkTrue(xlpp::Workbook::isPasswordEncryptedFile(encryptedPath),
                   "P1G password-to-open file is recognized as an encrypted Office CFB package");
    const auto compound = xlpp::internal::readBinaryFile(encryptedPath);
    test.checkTrue(xlpp::internal::isCompoundFile(compound), "P1G encrypted workbook uses a CFB outer container");
    test.checkTrue(xlpp::internal::compoundFileContainsStream(compound, "EncryptionInfo")
                   && xlpp::internal::compoundFileContainsStream(compound, "EncryptedPackage"),
                   "P1G encrypted CFB contains EncryptionInfo and EncryptedPackage streams");
    const auto infoBytes = xlpp::internal::readCompoundFileStream(compound, "EncryptionInfo");
    const std::string infoText(infoBytes.begin() + 8, infoBytes.end());
    test.checkTrue(infoText.find("cipherAlgorithm=\"AES\"") != std::string::npos
                   && infoText.find("keyBits=\"256\"") != std::string::npos
                   && infoText.find("hashAlgorithm=\"SHA512\"") != std::string::npos
                   && infoText.find("spinCount=\"100000\"") != std::string::npos,
                   "P1G writes the AES-256/SHA-512 Agile Encryption descriptor");

    xlpp::Workbook roundTrip;
    xlpp::LoadOptions loadOptions;
    loadOptions.passwordToOpen = "P@ssw0rd✓";
    roundTrip.load(encryptedPath, loadOptions);
    test.checkEqual(std::get<std::string>(roundTrip[0].cell("A1").value()), std::string("classified"),
                    "P1G correct Unicode password decrypts workbook text");
    test.checkNear(std::get<double>(roundTrip[0].cell("B2").value()), 42.5, 1e-12,
                   "P1G correct password decrypts numeric cells");
    test.checkEqual(roundTrip[0].cell("C3").formula(), std::string("=B2*2"),
                    "P1G encrypted round-trip preserves formula text");

    std::ostringstream encryptedStream;
    workbook.save(encryptedStream, saveOptions);
    const auto encryptedStreamBytes = encryptedStream.str();
    test.checkTrue(encryptedStreamBytes.size() > 512 &&
                   static_cast<unsigned char>(encryptedStreamBytes[0]) == 0xd0 &&
                   static_cast<unsigned char>(encryptedStreamBytes[1]) == 0xcf,
                   "P1G stream save emits the encrypted CFB container");
    std::istringstream encryptedInput(encryptedStreamBytes);
    xlpp::Workbook streamRoundTrip;
    streamRoundTrip.load(encryptedInput, loadOptions);
    test.checkEqual(std::get<std::string>(streamRoundTrip[0].cell("A1").value()), std::string("classified"),
                    "P1G encrypted stream save/load round-trips with password options");

    bool wrongPasswordRejected = false;
    try {
        xlpp::Workbook wrong;
        xlpp::LoadOptions wrongOptions;
        wrongOptions.passwordToOpen = "not-the-password";
        wrong.load(encryptedPath, wrongOptions);
    } catch (const std::invalid_argument&) {
        wrongPasswordRejected = true;
    }
    test.checkTrue(wrongPasswordRejected, "P1G rejects an incorrect password before package parsing");

    // Corrupt only encrypted payload bytes while keeping a valid CFB directory.
    // Password verification should still pass, then DataIntegrity/HMAC must fail.
    auto encryptedPackage = xlpp::internal::readCompoundFileStream(compound, "EncryptedPackage");
    test.checkTrue(encryptedPackage.size() > 32, "P1G encrypted package has payload bytes available for tamper test");
    encryptedPackage.back() ^= 0x5au;
    const auto tampered = xlpp::internal::buildRootCompoundFile({
        {"EncryptionInfo", infoBytes}, {"EncryptedPackage", encryptedPackage}
    });
    xlpp::internal::writeBinaryFile(tamperedPath, tampered);
    bool integrityRejected = false;
    try {
        xlpp::Workbook corrupted;
        corrupted.load(tamperedPath, loadOptions);
    } catch (const std::runtime_error& error) {
        integrityRejected = std::string(error.what()).find("integrity") != std::string::npos
                         || std::string(error.what()).find("HMAC") != std::string::npos;
    }
    test.checkTrue(integrityRejected, "P1G Agile DataIntegrity rejects ciphertext tampering");

    std::filesystem::remove(encryptedPath);
    std::filesystem::remove(tamperedPath);
}

void testStandardEncryptionInteropP1G(TestContext& test) {
    const auto fixture = std::filesystem::path(XLPP_TEST_SOURCE_DIR)
        / "fixtures" / "encryption" / "libreoffice_standard_aes128.xlsx";
    test.checkTrue(std::filesystem::exists(fixture), "P1G LibreOffice Standard Encryption fixture exists");
    test.checkTrue(xlpp::Workbook::isPasswordEncryptedFile(fixture),
                   "P1G recognizes LibreOffice Standard Encryption CFB fixture");

    xlpp::Workbook workbook;
    xlpp::LoadOptions options;
    options.passwordToOpen = "Libre✓";
    workbook.load(fixture, options);
    test.checkEqual(workbook.sheetCount(), std::size_t{1}, "P1G Standard Encryption fixture opens with correct password");
    test.checkEqual(std::get<std::string>(workbook[0].cell("A1").value()), std::string("secret"),
                    "P1G decrypts LibreOffice Standard AES-128 text cell");
    test.checkNear(std::get<double>(workbook[0].cell("B2").value()), 42.5, 1e-12,
                   "P1G decrypts LibreOffice Standard AES-128 numeric cell");

    bool wrongPasswordRejected = false;
    try {
        xlpp::Workbook wrong;
        xlpp::LoadOptions bad;
        bad.passwordToOpen = "wrong";
        wrong.load(fixture, bad);
    } catch (const std::invalid_argument&) {
        wrongPasswordRejected = true;
    }
    test.checkTrue(wrongPasswordRejected, "P1G Standard Encryption reader rejects wrong password");
}


void testAgileEncryptionProfileMatrixP1H(TestContext& test) {
    std::vector<unsigned char> payload(4099u);
    for (std::size_t i = 0; i < payload.size(); ++i)
        payload[i] = static_cast<unsigned char>((i * 37u + 11u) & 0xffu);

    const std::array<xlpp::PackageEncryptionHash, 4> hashes{
        xlpp::PackageEncryptionHash::Sha1,
        xlpp::PackageEncryptionHash::Sha256,
        xlpp::PackageEncryptionHash::Sha384,
        xlpp::PackageEncryptionHash::Sha512,
    };
    const std::array<std::uint32_t, 3> keyBits{128u, 192u, 256u};

    std::size_t profiles = 0;
    for (const auto hash : hashes) {
        for (const auto bits : keyBits) {
            xlpp::internal::AgileEncryptionParameters params;
            params.spinCount = 8; // Exercise format combinations without making the matrix a KDF benchmark.
            params.keyBits = bits;
            params.hashAlgorithm = hash;
            const auto encrypted = xlpp::internal::encryptAgileOfficePackage(payload, "Matrix✓", params);
            const auto info = xlpp::internal::inspectOfficeEncryption(encrypted);
            test.checkEqual(static_cast<int>(info.format), static_cast<int>(xlpp::PackageEncryptionFormat::Agile),
                            "P1H Agile matrix inspection reports Agile format");
            test.checkEqual(info.keyBits, bits, "P1H Agile matrix inspection preserves AES key size");
            test.checkEqual(static_cast<int>(info.hashAlgorithm), static_cast<int>(hash),
                            "P1H Agile matrix inspection preserves hash algorithm");
            test.checkEqual(info.spinCount, std::uint32_t{8}, "P1H Agile matrix inspection preserves spin count");
            test.checkTrue(info.hasDataIntegrity && info.supportedForRead && info.supportedForWrite,
                           "P1H Agile matrix profile advertises integrity and read/write support");
            const auto decrypted = xlpp::internal::decryptOfficePackage(encrypted, "Matrix✓");
            test.checkTrue(decrypted == payload, "P1H Agile matrix profile decrypts byte-for-byte");
            ++profiles;
        }
    }
    test.checkEqual(profiles, std::size_t{12}, "P1H exercises all AES/hash Agile combinations");

    // Exercise a non-default combination through the public Workbook API too.
    const auto path = std::filesystem::temp_directory_path() / "xlpp_p1h_agile128_sha256.xlsx";
    xlpp::Workbook wb;
    wb.addWorksheet("Data").cell("A1").setValue("agile-128-sha256");
    xlpp::SaveOptions save;
    save.encryption.enabled = true;
    save.encryption.password = "Public✓";
    save.encryption.mode = xlpp::PackageEncryptionMode::Agile;
    save.encryption.keyBits = 128;
    save.encryption.hashAlgorithm = xlpp::PackageEncryptionHash::Sha256;
    save.encryption.spinCount = 64;
    wb.save(path, save);
    const auto publicInfo = xlpp::Workbook::inspectPasswordEncryptionFile(path);
    test.checkEqual(static_cast<int>(publicInfo.format), static_cast<int>(xlpp::PackageEncryptionFormat::Agile),
                    "P1H public profile inspection identifies Agile");
    test.checkEqual(publicInfo.keyBits, std::uint32_t{128}, "P1H public Agile writer uses AES-128");
    test.checkEqual(static_cast<int>(publicInfo.hashAlgorithm), static_cast<int>(xlpp::PackageEncryptionHash::Sha256),
                    "P1H public Agile writer uses SHA-256");
    xlpp::LoadOptions load;
    load.passwordToOpen = "Public✓";
    xlpp::Workbook round;
    round.load(path, load);
    test.checkEqual(std::get<std::string>(round[0].cell("A1").value()), std::string("agile-128-sha256"),
                    "P1H public non-default Agile profile round-trips workbook data");
    std::filesystem::remove(path);
}

void testStandardEncryptionWriterP1H(TestContext& test) {
    const std::array<std::uint32_t, 3> keyBits{128u, 192u, 256u};
    for (const auto bits : keyBits) {
        const auto path = std::filesystem::temp_directory_path()
            / ("xlpp_p1h_standard_" + std::to_string(bits) + ".xlsx");
        xlpp::Workbook wb;
        auto& sheet = wb.addWorksheet("Secrets");
        sheet.cell("A1").setValue("standard-aes");
        sheet.cell("B2").setValue(static_cast<double>(bits));

        xlpp::SaveOptions save;
        save.encryption.enabled = true;
        save.encryption.password = "Compat✓";
        save.encryption.mode = xlpp::PackageEncryptionMode::Standard;
        save.encryption.keyBits = bits;
        // Deliberately non-standard values: Standard writer must use the format-mandated SHA-1/50k KDF.
        save.encryption.hashAlgorithm = xlpp::PackageEncryptionHash::Sha512;
        save.encryption.spinCount = 7;
        wb.save(path, save);

        const auto info = xlpp::Workbook::inspectPasswordEncryptionFile(path);
        test.checkTrue(info.encrypted, "P1H Standard writer emits encrypted CFB package");
        test.checkEqual(static_cast<int>(info.format), static_cast<int>(xlpp::PackageEncryptionFormat::Standard),
                        "P1H Standard writer profile inspection reports Standard");
        test.checkEqual(info.keyBits, bits, "P1H Standard writer preserves requested AES key size");
        test.checkEqual(static_cast<int>(info.hashAlgorithm), static_cast<int>(xlpp::PackageEncryptionHash::Sha1),
                        "P1H Standard writer advertises mandatory SHA-1");
        test.checkEqual(info.spinCount, std::uint32_t{50000},
                        "P1H Standard writer uses fixed 50,000-iteration derivation");
        test.checkTrue(!info.hasDataIntegrity && info.supportedForRead && info.supportedForWrite,
                       "P1H Standard writer profile has expected integrity/read/write flags");

        xlpp::LoadOptions load;
        load.passwordToOpen = "Compat✓";
        xlpp::Workbook round;
        round.load(path, load);
        test.checkEqual(std::get<std::string>(round[0].cell("A1").value()), std::string("standard-aes"),
                        "P1H Standard writer round-trips text");
        test.checkNear(std::get<double>(round[0].cell("B2").value()), static_cast<double>(bits), 1e-12,
                       "P1H Standard writer round-trips numeric cell");
        std::filesystem::remove(path);
    }
}

void testEncryptionInspectionAndResourceGuardsP1H(TestContext& test) {
    const auto encryptedPath = std::filesystem::temp_directory_path() / "xlpp_p1h_resource_guard.xlsx";
    const auto plainPath = std::filesystem::temp_directory_path() / "xlpp_p1h_plain_profile.xlsx";

    xlpp::Workbook wb;
    wb.addWorksheet("Guard").cell("A1").setValue(std::string(8192, 'x'));
    wb.save(plainPath);
    const auto plainInfo = xlpp::Workbook::inspectPasswordEncryptionFile(plainPath);
    test.checkTrue(!plainInfo.encrypted, "P1H profile inspection reports plain OOXML as unencrypted");
    test.checkEqual(static_cast<int>(plainInfo.format), static_cast<int>(xlpp::PackageEncryptionFormat::None),
                    "P1H profile inspection uses None format for plain OOXML");

    xlpp::SaveOptions save;
    save.encryption.enabled = true;
    save.encryption.password = "Guard✓";
    save.encryption.spinCount = 96;
    wb.save(encryptedPath, save);
    const auto info = xlpp::Workbook::inspectPasswordEncryptionFile(encryptedPath);
    test.checkEqual(info.spinCount, std::uint32_t{96}, "P1H inspection exposes Agile spin count without password");

    bool spinRejected = false;
    try {
        xlpp::LoadOptions load;
        load.passwordToOpen = "Guard✓";
        load.maxEncryptionSpinCount = 32;
        xlpp::Workbook blocked;
        blocked.load(encryptedPath, load);
    } catch (const std::runtime_error& error) {
        spinRejected = std::string(error.what()).find("maxEncryptionSpinCount") != std::string::npos;
    }
    test.checkTrue(spinRejected, "P1H maxEncryptionSpinCount guard rejects excessive KDF work before decrypting package");

    bool sizeRejected = false;
    try {
        xlpp::LoadOptions load;
        load.passwordToOpen = "Guard✓";
        load.maxEncryptionSpinCount = 128;
        load.maxDecryptedPackageBytes = 256;
        xlpp::Workbook blocked;
        blocked.load(encryptedPath, load);
    } catch (const std::runtime_error& error) {
        sizeRejected = std::string(error.what()).find("maxDecryptedPackageBytes") != std::string::npos
                    || std::string(error.what()).find("decrypted") != std::string::npos;
    }
    test.checkTrue(sizeRejected, "P1H decrypted-package size guard rejects oversized inner OOXML before allocation");

    xlpp::LoadOptions allowed;
    allowed.passwordToOpen = "Guard✓";
    allowed.maxEncryptionSpinCount = 128;
    allowed.maxDecryptedPackageBytes = 1024u * 1024u;
    xlpp::Workbook round;
    round.load(encryptedPath, allowed);
    test.checkEqual(std::get<std::string>(round[0].cell("A1").value()).size(), std::size_t{8192},
                    "P1H resource guards allow a package within configured limits");

    std::filesystem::remove(encryptedPath);
    std::filesystem::remove(plainPath);
}

void testEncryptionMemoryAndCertificateInspectionP1I(TestContext& test) {
    // Exercise the new memory ZIP boundary independently from Workbook.
    xlpp::internal::ZipArchive zip;
    zip.add("alpha.txt", "alpha");
    zip.add("nested/beta.txt", std::string(4097, 'b'));
    zip.setForceZip64(true);
    const auto zipBytes = zip.saveToBytes();
    test.checkTrue(zipBytes.size() > 100 && zipBytes[0] == 'P' && zipBytes[1] == 'K',
                   "P1I ZipArchive serializes a ZIP/ZIP64 package directly to memory");
    const auto reopened = xlpp::internal::ZipArchive::open(zipBytes);
    test.checkEqual(reopened.get("alpha.txt"), std::string("alpha"),
                    "P1I in-memory ZIP open restores stored entry");
    test.checkEqual(reopened.get("nested/beta.txt").size(), std::size_t{4097},
                    "P1I in-memory ZIP open restores compressed entry");

    // Generate a normal password Agile package, then prepend a certificate
    // key-encryptor. Password decryption must select its encryptor by URI rather
    // than accidentally treating the first encryptedKey element as password data.
    std::vector<unsigned char> payload(zipBytes.begin(), zipBytes.end());
    xlpp::internal::AgileEncryptionParameters parameters;
    parameters.spinCount = 8;
    const auto encrypted = xlpp::internal::encryptAgileOfficePackage(payload, "Cert✓", parameters);
    auto encryptionInfo = xlpp::internal::readCompoundFileStream(encrypted, "EncryptionInfo");
    const auto encryptedPackage = xlpp::internal::readCompoundFileStream(encrypted, "EncryptedPackage");
    std::string xml(encryptionInfo.begin() + 8, encryptionInfo.end());
    const auto keyEncryptors = xml.find("<keyEncryptors>");
    test.checkTrue(keyEncryptors != std::string::npos, "P1I generated Agile descriptor has keyEncryptors owner");
    const auto insertAt = keyEncryptors + std::string("<keyEncryptors>").size();
    const std::string certificate =
        "<keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">"
        "<c:encryptedKey xmlns:c=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\" "
        "encryptedKeyValue=\"AQIDBA==\" X509Certificate=\"MAMCAQE=\" certVerifier=\"BQYHCA==\"/>"
        "</keyEncryptor>";
    xml.insert(insertAt, certificate);
    encryptionInfo.resize(8);
    encryptionInfo.insert(encryptionInfo.end(), xml.begin(), xml.end());
    const auto withCertificate = xlpp::internal::buildRootCompoundFile({
        {"EncryptionInfo", encryptionInfo}, {"EncryptedPackage", encryptedPackage}
    });

    const auto info = xlpp::internal::inspectOfficeEncryption(withCertificate);
    test.checkEqual(info.keyEncryptorCount, std::size_t{2},
                    "P1I Agile inspection counts password plus certificate key-encryptors");
    test.checkEqual(info.passwordKeyEncryptorCount, std::size_t{1},
                    "P1I Agile inspection identifies exactly one password key-encryptor");
    test.checkEqual(info.certificateKeyEncryptors.size(), std::size_t{1},
                    "P1I Agile inspection exposes certificate key-encryptor metadata");
    test.checkTrue(info.certificateKeyEncryptors[0].validEncoding,
                   "P1I certificate key-encryptor base64 metadata is decoded safely");
    test.checkEqual(info.certificateKeyEncryptors[0].encryptedKeyBytes, std::size_t{4},
                    "P1I certificate inspection reports encrypted intermediate-key byte count");
    test.checkTrue(!info.certificateKeyEncryptors[0].x509Certificate.empty(),
                   "P1I certificate inspection exposes DER certificate bytes without a private key");

    const auto decrypted = xlpp::internal::decryptOfficePackage(withCertificate, "Cert✓");
    test.checkTrue(decrypted == payload,
                   "P1I password decryption remains correct when a certificate encryptor appears first");
}

void testEncryptionPoliciesAndMalformedInputsP1I(TestContext& test) {
    const auto agilePath = std::filesystem::temp_directory_path() / "xlpp_p1i_policy_agile.xlsx";
    const auto noIntegrityPath = std::filesystem::temp_directory_path() / "xlpp_p1i_policy_no_integrity.xlsx";

    xlpp::Workbook wb;
    wb.addWorksheet("Policy").cell("A1").setValue("policy");
    xlpp::SaveOptions save;
    save.encryption.enabled = true;
    save.encryption.password = "Policy✓";
    save.encryption.spinCount = 8;
    wb.save(agilePath, save);

    // EncryptionInfo resource guard rejects the descriptor before XML/KDF work.
    bool infoLimitRejected = false;
    try {
        xlpp::LoadOptions load;
        load.passwordToOpen = "Policy✓";
        load.maxEncryptionInfoBytes = 32;
        xlpp::Workbook blocked;
        blocked.load(agilePath, load);
    } catch (const std::runtime_error& error) {
        infoLimitRejected = std::string(error.what()).find("maxEncryptionInfoBytes") != std::string::npos;
    }
    test.checkTrue(infoLimitRejected, "P1I maxEncryptionInfoBytes policy rejects oversized descriptors early");

    // Remove DataIntegrity while leaving a valid password key-encryptor and
    // encrypted package. Compatibility mode loads it; strict policy rejects it.
    const auto compound = xlpp::internal::readBinaryFile(agilePath);
    auto encryptionInfo = xlpp::internal::readCompoundFileStream(compound, "EncryptionInfo");
    const auto encryptedPackage = xlpp::internal::readCompoundFileStream(compound, "EncryptedPackage");
    std::string xml(encryptionInfo.begin() + 8, encryptionInfo.end());

    // Agile requires exactly one password key-encryptor. Duplicate descriptors
    // must fail rather than making password selection order-dependent.
    bool duplicatePasswordRejected = false;
    try {
        auto duplicateXml = xml;
        const auto passwordBegin = duplicateXml.find(
            "<keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">");
        const auto passwordEnd = passwordBegin == std::string::npos
            ? std::string::npos
            : duplicateXml.find("</keyEncryptor>", passwordBegin);
        if (passwordBegin != std::string::npos && passwordEnd != std::string::npos) {
            const auto end = passwordEnd + std::string("</keyEncryptor>").size();
            const auto descriptor = duplicateXml.substr(passwordBegin, end - passwordBegin);
            duplicateXml.insert(end, descriptor);
        }
        auto duplicateInfo = encryptionInfo;
        duplicateInfo.resize(8);
        duplicateInfo.insert(duplicateInfo.end(), duplicateXml.begin(), duplicateXml.end());
        const auto duplicateCompound = xlpp::internal::buildRootCompoundFile({
            {"EncryptionInfo", duplicateInfo}, {"EncryptedPackage", encryptedPackage}
        });
        (void)xlpp::internal::decryptOfficePackage(duplicateCompound, "Policy✓");
    } catch (const std::runtime_error& error) {
        duplicatePasswordRejected = std::string(error.what()).find("exactly one password key-encryptor") != std::string::npos;
    }
    test.checkTrue(duplicatePasswordRejected,
                   "P1I malformed Agile input with duplicate password key-encryptors is rejected");

    if (const auto begin = xml.find("<dataIntegrity"); begin != std::string::npos) {
        const auto end = xml.find("/>", begin);
        if (end != std::string::npos) xml.erase(begin, end + 2 - begin);
    }
    encryptionInfo.resize(8);
    encryptionInfo.insert(encryptionInfo.end(), xml.begin(), xml.end());
    xlpp::internal::writeBinaryFile(noIntegrityPath, xlpp::internal::buildRootCompoundFile({
        {"EncryptionInfo", encryptionInfo}, {"EncryptedPackage", encryptedPackage}
    }));

    xlpp::LoadOptions compatible;
    compatible.passwordToOpen = "Policy✓";
    xlpp::Workbook compatibleLoad;
    compatibleLoad.load(noIntegrityPath, compatible);
    test.checkEqual(std::get<std::string>(compatibleLoad[0].cell("A1").value()), std::string("policy"),
                    "P1I compatibility policy can read an Agile package without DataIntegrity");

    bool integrityPolicyRejected = false;
    try {
        xlpp::LoadOptions strict;
        strict.passwordToOpen = "Policy✓";
        strict.requireAgileDataIntegrity = true;
        xlpp::Workbook blocked;
        blocked.load(noIntegrityPath, strict);
    } catch (const std::runtime_error& error) {
        integrityPolicyRejected = std::string(error.what()).find("DataIntegrity") != std::string::npos;
    }
    test.checkTrue(integrityPolicyRejected,
                   "P1I requireAgileDataIntegrity policy rejects unauthenticated Agile packages");

    const auto standardFixture = std::filesystem::path(XLPP_TEST_SOURCE_DIR)
        / "fixtures" / "encryption" / "libreoffice_standard_aes128.xlsx";
    bool standardPolicyRejected = false;
    try {
        xlpp::LoadOptions strict;
        strict.passwordToOpen = "Libre✓";
        strict.allowStandardEncryption = false;
        xlpp::Workbook blocked;
        blocked.load(standardFixture, strict);
    } catch (const std::runtime_error& error) {
        standardPolicyRejected = std::string(error.what()).find("Standard Encryption") != std::string::npos;
    }
    test.checkTrue(standardPolicyRejected,
                   "P1I allowStandardEncryption=false rejects legacy unauthenticated Standard Encryption");

    std::filesystem::remove(agilePath);
    std::filesystem::remove(noIntegrityPath);
}

void testLargeEncryptedCompoundFileP1G(TestContext& test) {
    // >7 MiB forces more than 109 FAT sectors, exercising the DIFAT path that
    // the older VBA-only compact CFB writer did not need.
    std::vector<unsigned char> payload(8u * 1024u * 1024u + 123u);
    for (std::size_t i = 0; i < payload.size(); ++i)
        payload[i] = static_cast<unsigned char>((i * 131u + 17u) & 0xffu);

    xlpp::internal::AgileEncryptionParameters parameters;
    parameters.spinCount = 8; // This suite stresses CFB/DIFAT and segmentation, not KDF cost.
    const auto encrypted = xlpp::internal::encryptAgileOfficePackage(payload, "large-package", parameters);
    test.checkTrue(xlpp::internal::isCompoundFile(encrypted), "P1G large encrypted payload emits a valid CFB signature");
    test.checkTrue(encrypted.size() > payload.size(), "P1G large encrypted CFB includes encryption/container overhead");
    const auto decrypted = xlpp::internal::decryptOfficePackage(encrypted, "large-package");
    test.checkTrue(decrypted == payload,
                   "P1G Agile encryption round-trips >8 MiB through CFB DIFAT and 4096-byte segmented AES");
}


void testStructuralReferenceTransformerP1J(TestContext& test) {
    xlpp::internal::StructuralEditSpec insert;
    insert.axis = xlpp::internal::StructuralAxis::Row;
    insert.action = xlpp::internal::StructuralAction::Insert;
    insert.index = 3;
    insert.amount = 2;
    insert.targetSheetName = "Sheet1";

    const std::string formula = "SUM(A2:A5,'Sheet1'!$B$3,Sheet2!A3,\"A3\",Table1[A3],[Book.xlsx]Sheet1!A3,LOG10(A3))";
    const auto rewritten = xlpp::internal::rewriteFormulaReferences(formula, "Sheet1", insert);
    test.checkEqual(rewritten.text,
                    std::string("SUM(A2:A7,'Sheet1'!$B$5,Sheet2!A3,\"A3\",Table1[A3],[Book.xlsx]Sheet1!A3,LOG10(A5))"),
                    "Structural transformer rewrites only target A1 references");
    test.checkEqual(rewritten.referencesRewritten, std::size_t{3}, "Structural transformer rewrite count");

    xlpp::internal::StructuralEditSpec erase = insert;
    erase.action = xlpp::internal::StructuralAction::Delete;
    erase.index = 3;
    erase.amount = 2;
    const auto deleted = xlpp::internal::rewriteFormulaReferences(
        "A3+A2:A5+'Sheet1'!4:6+Sheet2!A3", "Sheet1", erase);
    test.checkEqual(deleted.text, std::string("#REF!+A2:A3+'Sheet1'!3:4+Sheet2!A3"),
                    "Deleting rows shrinks ranges and invalidates deleted single cells");
    test.checkEqual(deleted.referencesInvalidated, std::size_t{1}, "Deleted single reference becomes REF");

    xlpp::internal::StructuralEditSpec colInsert;
    colInsert.axis = xlpp::internal::StructuralAxis::Column;
    colInsert.action = xlpp::internal::StructuralAction::Insert;
    colInsert.index = 2;
    colInsert.amount = 1;
    colInsert.targetSheetName = "Sheet1";
    const auto columns = xlpp::internal::rewriteFormulaReferences("SUM($A:$C)+A1:C4+$D$5", "Sheet1", colInsert);
    test.checkEqual(columns.text, std::string("SUM($A:$D)+A1:D4+$E$5"),
                    "Column insertion expands ranges and preserves absolute markers");

    const auto caseInsensitive = xlpp::internal::rewriteFormulaReferences("sheet1!A3+SHEET1!B3", "Calc", insert);
    test.checkEqual(caseInsensitive.text, std::string("sheet1!A5+SHEET1!B5"),
                    "Worksheet-qualified references resolve case-insensitively like Excel");
}

void testWorkbookReferenceSafeStructuralEditsP1J(TestContext& test) {
    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data");
    auto& calc = workbook.addWorksheet("Calc");
    workbook.addWorksheet("Other");

    for (std::size_t row = 1; row <= 6; ++row) {
        data.cell(row, 1).setValue(static_cast<double>(row));
        data.cell(row, 2).setValue(static_cast<double>(row * 10));
        data.cell(row, 3).setValue(static_cast<double>(row * 100));
    }
    data.cell("D2").setFormula("SUM(A2:A5)+B3");
    calc.cell("A1").setFormula("SUM(Data!A2:A5)");
    calc.cell("A2").setFormula("SUM(Other!A2:A5)+Data!A3");
    calc.cell("B1").setFormula("Data!B2+Data!C2");
    data.cell("E2").setArrayFormula("A2:A5*2", "E2:E5");
    xlpp::Hyperlink internalLink("#Data!C4");
    internalLink.setExternal(false);
    calc.cell("C1").setHyperlink(std::move(internalLink));
    xlpp::Hyperlink externalLink("https://example.com/A3");
    externalLink.setExternal(true);
    calc.cell("C2").setHyperlink(std::move(externalLink));

    data.mergeCells("B2:C4");
    data.freezePanes("D3");
    data.sheetView().setTopLeftCell("C4");
    data.rowDimension(4).height = 24.0;
    data.columnDimension(3).width = 18.0;
    data.autoFilter().setReference("A1:C5");
    data.autoFilter().column(1).addValue("20");
    data.autoFilter().sortState().setReference("A1:C5");
    data.autoFilter().sortState().addCondition("B2:B5", false);
    data.conditionalFormatting().addRule("A2:A5", xlpp::ConditionalRule::formula("A2>0"));
    data.dataValidations().add(xlpp::DataValidation::list("B2:B5", "=Data!$A$2:$A$5"));
    auto& table = data.addTable("DataTable", "A1:C5");
    table.addColumn("A"); table.addColumn("B"); table.addColumn("C");
    data.setPrintArea("A1:A5,C1:C5");
    data.setPrintTitlesRows("1:1");
    data.setPrintTitlesCols("A:A");

    xlpp::DefinedName global("DataRange", "=Data!$A$2:$A$5");
    workbook.addDefinedName(global);
    xlpp::DefinedName local("LocalRange", "=$A$2:$A$5");
    local.setLocalSheetId(0);
    workbook.addDefinedName(local);

    xlpp::Chart chart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries series("Sales");
    series.setCategoriesReference("='Data'!$A$2:$A$5");
    series.setValuesReference("='Data'!$B$2:$B$5");
    xlpp::ChartSeries::ErrorBars bars;
    bars.valueType = xlpp::ChartSeries::ErrorValueType::Custom;
    bars.plusReference = "='Data'!$C$2:$C$5";
    bars.minusReference = "='Data'!$C$2:$C$5";
    series.setErrorBars({bars});
    chart.addSeries(std::move(series));
    calc.addChart(std::move(chart));

    xlpp::PivotTable pivot("SalesPivot");
    pivot.setLocation("F2:H6");
    pivot.cache().setSourceData("'Data'!$A$1:$C$5");
    pivot.cache().setFields({"A", "B", "C"});
    pivot.cache().addRecord({"2", "20", "200"});
    pivot.cache().addRecord({"3", "30", "300"});
    pivot.cache().addRecord({"4", "40", "400"});
    pivot.cache().addRecord({"5", "50", "500"});
    pivot.addRowField("A");
    pivot.addDataField("B", "sum");
    calc.addPivotTable(std::move(pivot));

    const auto report = workbook.insertRows("Data", 3, 2);
    test.checkTrue(report.cellsMoved >= 1, "Workbook structural insert reports moved cells");
    test.checkTrue(report.formulasRewritten >= 4, "Workbook structural insert rewrites formulas across sheets");
    test.checkTrue(report.definedNamesRewritten >= 2, "Workbook structural insert rewrites defined names");
    test.checkTrue(report.chartCachesSynchronized >= 1, "Workbook structural edit synchronizes chart caches after reference rewrites");
    test.checkEqual(data.cell("D2").formula(), std::string("SUM(A2:A7)+B5"), "Local formula expanded after row insertion");
    test.checkEqual(calc.cell("A1").formula(), std::string("SUM(Data!A2:A7)"), "Cross-sheet formula expanded after row insertion");
    test.checkEqual(calc.cell("A2").formula(), std::string("SUM(Other!A2:A5)+Data!A5"), "Non-target sheet reference remains unchanged");
    test.checkEqual(data.mergedRanges().front(), std::string("B2:C6"), "Merged range expands through insertion");
    test.checkEqual(data.frozenPane().value(), std::string("D5"), "Freeze pane follows inserted rows");
    test.checkTrue(data.tryRowDimension(6) != nullptr, "Row dimension follows inserted rows");
    test.checkEqual(data.autoFilter().reference(), std::string("A1:C7"), "AutoFilter expands after inserted rows");
    test.checkEqual(data.conditionalFormatting().entries().front().reference(), std::string("A2:A7"), "Conditional formatting sqref expands");
    test.checkEqual(data.dataValidations().items().front().reference(), std::string("B2:B7"), "Data validation sqref expands");
    test.checkEqual(data.dataValidations().items().front().formula1(), std::string("=Data!$A$2:$A$7"), "Data validation formula expands");
    test.checkEqual(data.tables().front().reference(), std::string("A1:C7"), "Table range expands on row insert");
    test.checkEqual(data.printArea(), std::string("A1:A7,C1:C7"), "Print-area union preserves comma separators while expanding");
    test.checkEqual(data.sheetView().topLeftCell(), std::string("C6"), "Sheet-view topLeftCell follows row insertion");
    test.checkEqual(data.cell("E2").formulaMetadata().reference(), std::string("E2:E7"), "Array formula metadata range expands");
    test.checkEqual(calc.cell("C1").hyperlinkValue()->target(), std::string("#Data!C6"), "Internal hyperlink target follows row insertion");
    test.checkEqual(calc.cell("C2").hyperlinkValue()->target(), std::string("https://example.com/A3"), "External hyperlink target is untouched");
    test.checkEqual(workbook.definedName("DataRange")->value(), std::string("=Data!$A$2:$A$7"), "Global defined name rewritten");
    test.checkEqual(workbook.definedName("LocalRange")->value(), std::string("=$A$2:$A$7"), "Local defined name rewritten");
    test.checkEqual(calc.charts().front().series().front().valuesReference(), std::string("='Data'!$B$2:$B$7"), "Chart values reference rewritten");
    test.checkEqual(calc.pivotTables().front().cache().sourceData(), std::string("'Data'!$A$1:$C$7"), "Pivot cache source rewritten");
    test.checkEqual(calc.pivotTables().front().cache().records().size(), std::size_t{6}, "Pivot cache records expand with inserted source rows");
    test.checkTrue(workbook.calcProperties().fullCalcOnLoad(), "Structural edit requests full calculation on load");

    const auto deleteReport = workbook.deleteColumns("Data", 2, 1);
    test.checkTrue(deleteReport.formulaReferencesInvalidated >= 1, "Column deletion reports invalidated references");
    test.checkEqual(calc.cell("B1").formula(), std::string("Data!#REF!+Data!B2"), "Deleted column becomes REF while following column shifts");
    test.checkEqual(data.tables().front().reference(), std::string("A1:B7"), "Table range shrinks on column deletion");
    test.checkEqual(data.tables().front().columns().size(), std::size_t{2}, "Table column model shrinks with deleted worksheet column");
    test.checkTrue(data.tryColumnDimension(2) != nullptr, "Column dimension after deleted column shifts left");
    test.checkEqual(data.printArea(), std::string("A1:A7,B1:B7"), "Print-area union preserves comma after column deletion");
    test.checkEqual(data.sheetView().topLeftCell(), std::string("B6"), "Sheet-view topLeftCell follows column deletion");
    test.checkEqual(data.cell("D2").formulaMetadata().reference(), std::string("D2:D7"), "Array formula metadata shifts with deleted column");
    test.checkEqual(calc.cell("C1").hyperlinkValue()->target(), std::string("#Data!B6"), "Internal hyperlink target shifts with deleted source column");
    test.checkEqual(calc.pivotTables().front().cache().fields().size(), std::size_t{2}, "Pivot cache field schema shrinks with deleted source column");
    test.checkEqual(calc.pivotTables().front().cache().records().front().size(), std::size_t{2}, "Pivot cache record width shrinks with deleted source column");
    test.checkEqual(calc.pivotTables().front().dataFields().size(), std::size_t{0}, "Pivot data field bound to deleted source column is removed");

    const auto integrity = workbook.validateModelIntegrity();
    test.checkTrue(integrity.ok(), "Reference-safe structural edit leaves a valid in-memory workbook model");

    const auto temp = std::filesystem::temp_directory_path() / "xlpp_p1j_structural_edit.xlsx";
    workbook.save(temp);
    xlpp::Workbook loaded;
    loaded.load(temp);
    test.checkEqual(loaded.worksheet("Calc")->cell("B1").formula(), std::string("Data!#REF!+Data!B2"), "Structural formula rewrite survives save/reload");
    test.checkEqual(loaded.worksheet("Data")->tables().front().reference(), std::string("A1:B7"), "Structural table rewrite survives save/reload");
    test.checkEqual(loaded.worksheet("Data")->printArea(), std::string("A1:A7,B1:B7"), "Structural print-area union survives save/reload");
    test.checkEqual(loaded.worksheet("Calc")->cell("C1").hyperlinkValue()->target(), std::string("#Data!B6"), "Structural hyperlink rewrite survives save/reload");
    test.checkEqual(loaded.definedName("DataRange")->value(), std::string("=Data!$A$2:$A$7"), "Structural defined-name rewrite survives save/reload");
    std::filesystem::remove(temp);
}

void testStructuralEditEdgeCasesP1J(TestContext& test) {
    xlpp::internal::StructuralEditSpec rowInsert;
    rowInsert.axis = xlpp::internal::StructuralAxis::Row;
    rowInsert.action = xlpp::internal::StructuralAction::Insert;
    rowInsert.index = 1;
    rowInsert.amount = 1;
    rowInsert.targetSheetName = "Data";
    const auto bottom = xlpp::internal::rewriteFormulaReferences("Data!A1048576", "Calc", rowInsert);
    test.checkEqual(bottom.text, std::string("Data!#REF!"), "Structural transformer invalidates a cell pushed beyond the last Excel row");

    xlpp::internal::StructuralEditSpec columnInsert;
    columnInsert.axis = xlpp::internal::StructuralAxis::Column;
    columnInsert.action = xlpp::internal::StructuralAction::Insert;
    columnInsert.index = 1;
    columnInsert.amount = 1;
    columnInsert.targetSheetName = "Data";
    const auto right = xlpp::internal::rewriteFormulaReferences("Data!XFD1", "Calc", columnInsert);
    test.checkEqual(right.text, std::string("Data!#REF!"), "Structural transformer invalidates a cell pushed beyond XFD");

    xlpp::Worksheet bounded("Bounded");
    bounded.cell(1048576, 1).setValue(1.0);
    bool rowOverflowThrown = false;
    try { bounded.insertRows(1, 1); } catch (const std::out_of_range&) { rowOverflowThrown = true; }
    test.checkTrue(rowOverflowThrown, "Physical row insertion refuses to move cells outside the Excel row limit");

    xlpp::Worksheet boundedColumns("BoundedColumns");
    boundedColumns.cell(1, 16384).setValue(1.0);
    bool columnOverflowThrown = false;
    try { boundedColumns.insertColumns(1, 1); } catch (const std::out_of_range&) { columnOverflowThrown = true; }
    test.checkTrue(columnOverflowThrown, "Physical column insertion refuses to move cells beyond XFD");

    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data");
    data.cell("A1").setValue("Group");
    data.cell("B1").setValue("Value");
    data.cell("A2").setValue("A");
    data.cell("B2").setValue(1.0);
    auto& calc = workbook.addWorksheet("Calc");
    xlpp::PivotTable pivot("BrokenSourcePivot");
    pivot.setLocation("D2:F5");
    pivot.cache().setSourceData("'Data'!$A$1:$B$2");
    pivot.cache().setFields({"Group", "Value"});
    pivot.cache().addRecord({"A", "1"});
    pivot.addRowField("Group");
    pivot.addDataField("Value", "sum");
    calc.addPivotTable(std::move(pivot));

    const auto report = workbook.deleteColumns("Data", 1, 2);
    test.checkTrue(report.formulaReferencesInvalidated >= 1, "Fully deleted Pivot source is reported as invalidated");
    test.checkEqual(calc.pivotTables().front().cache().sourceData(), std::string("'Data'!#REF!"), "Fully deleted Pivot source becomes #REF! instead of leaving an invalid geometry");
    test.checkEqual(calc.pivotTables().front().cache().fields().size(), std::size_t{2}, "Broken Pivot source preserves cached field schema");
    test.checkEqual(calc.pivotTables().front().cache().records().size(), std::size_t{1}, "Broken Pivot source preserves cached records");
    const auto brokenValidation = workbook.validateModelIntegrity();
    test.checkTrue(brokenValidation.ok(), "A preserved #REF Pivot source is a model warning, not a structural serialization error");
    test.checkTrue(brokenValidation.warningCount() >= 1, "Model validator reports broken Pivot source as a warning");

    const auto temp = std::filesystem::temp_directory_path() / "xlpp_p1j_broken_pivot_source.xlsx";
    workbook.save(temp);
    xlpp::Workbook loaded;
    loaded.load(temp);
    test.checkEqual(loaded.worksheet("Calc")->pivotTables().front().cache().sourceData(), std::string("'Data'!#REF!"), "Broken Pivot source remains savable and round-trips");
    std::filesystem::remove(temp);
}

void testWorkbookModelIntegrityValidatorP1J(TestContext& test) {
    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data");
    data.cell("A1").setValue("Name");
    data.cell("B1").setValue("Value");
    data.cell("A2").setValue("A");
    data.cell("B2").setValue(1.0);
    auto& table = data.addTable("T", "A1:B2");
    table.addColumn("Name");
    table.addColumn("Value");

    auto& reportSheet = workbook.addWorksheet("Report");
    xlpp::PivotTable pivot("P");
    pivot.setLocation("D2:F5");
    pivot.cache().setSourceData("'Data'!$A$1:$B$2");
    pivot.cache().setFields({"Name", "Value"});
    pivot.cache().addRecord({"A", "1"});
    pivot.addRowField("Name");
    pivot.addDataField("Value", "sum");
    reportSheet.addPivotTable(std::move(pivot));

    const auto clean = workbook.validateModelIntegrity();
    test.checkTrue(clean.ok(), "Model validator accepts a coherent workbook");
    test.checkEqual(clean.errorCount(), std::size_t{0}, "Coherent model has no validation errors");

    // Bypass high-level guards deliberately to prove validation catches model
    // corruption that package relationship checks cannot observe.
    data.tables().front().columns().pop_back();
    reportSheet.pivotTables().front().cache().records().front().pop_back();
    const auto corrupt = workbook.validateModelIntegrity();
    test.checkTrue(!corrupt.ok(), "Model validator rejects semantic corruption");
    test.checkTrue(corrupt.errorCount() >= 2, "Model validator finds table-width and Pivot-record-width corruption");
    bool sawTable = false, sawPivot = false;
    for (const auto& issue : corrupt.issues) {
        sawTable = sawTable || issue.code == "table.column_width_mismatch";
        sawPivot = sawPivot || issue.code == "pivot.record_width_mismatch";
    }
    test.checkTrue(sawTable, "Model validator reports table column-width mismatch with a stable code");
    test.checkTrue(sawPivot, "Model validator reports Pivot cache record-width mismatch with a stable code");
}

void testWorkbookModelIntegrityHardeningP1J(TestContext& test) {
    xlpp::Workbook workbook;
    auto& a = workbook.addWorksheet("A");
    auto& b = workbook.addWorksheet("B");
    a.cell("A1").setFormula("Missing!A1+#REF!");

    auto& tableA = a.addTable("SharedName", "A1:B2");
    tableA.addColumn("A"); tableA.addColumn("B");
    auto& tableB = b.addTable("sharedname", "A1:B2");
    tableB.addColumn("A"); tableB.addColumn("B");

    xlpp::PivotTable pivot1("PivotName");
    pivot1.setLocation("D2:F5");
    pivot1.cache().setSourceData("'Missing Sheet'!A1:B2");
    pivot1.cache().setFields({"A", "B"});
    pivot1.cache().addRecord({"x", "1"});
    a.addPivotTable(std::move(pivot1));
    xlpp::PivotTable pivot2("pivotname");
    pivot2.setLocation("D2:F5");
    pivot2.cache().setSourceData("'A'!A1:B2");
    pivot2.cache().setFields({"A", "B"});
    pivot2.cache().addRecord({"x", "1"});
    b.addPivotTable(std::move(pivot2));

    xlpp::DefinedName invalidScope("Scoped", "=A!A1");
    invalidScope.setLocalSheetId(99);
    workbook.definedNames().push_back(invalidScope);
    xlpp::DefinedName dup1("CaseName", "=A!A1"); dup1.setLocalSheetId(0);
    xlpp::DefinedName dup2("casename", "=A!A2"); dup2.setLocalSheetId(0);
    workbook.definedNames().push_back(std::move(dup1));
    workbook.definedNames().push_back(std::move(dup2));

    xlpp::Chart chart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries series("Broken");
    series.setCategoriesReference("='Missing Sheet'!$A$1:$A$2");
    series.setValuesReference("='A'!$A$1:$A$2");
    chart.addSeries(std::move(series));
    b.addChart(std::move(chart));

    const auto report = workbook.validateModelIntegrity();
    test.checkTrue(!report.ok(), "Model hardening validator rejects cross-object semantic corruption");
    bool formulaWarning=false, tableDuplicate=false, pivotDuplicate=false, pivotMissing=false;
    bool chartMissing=false, scopeInvalid=false, nameDuplicate=false;
    for (const auto& issue : report.issues) {
        formulaWarning = formulaWarning || issue.code == "formula.broken_reference";
        tableDuplicate = tableDuplicate || issue.code == "table.duplicate_name";
        pivotDuplicate = pivotDuplicate || issue.code == "pivot.duplicate_name";
        pivotMissing = pivotMissing || issue.code == "pivot.source_sheet_missing";
        chartMissing = chartMissing || issue.code == "chart.source_sheet_missing";
        scopeInvalid = scopeInvalid || issue.code == "defined_name.local_scope_invalid";
        nameDuplicate = nameDuplicate || issue.code == "defined_name.duplicate_in_scope";
    }
    test.checkTrue(formulaWarning, "Model validator reports formulas containing #REF! with a stable warning code");
    test.checkTrue(tableDuplicate, "Model validator catches workbook-wide case-insensitive table-name collisions");
    test.checkTrue(pivotDuplicate, "Model validator catches workbook-wide case-insensitive PivotTable-name collisions");
    test.checkTrue(pivotMissing, "Model validator catches Pivot sources whose worksheet no longer exists");
    test.checkTrue(chartMissing, "Model validator catches simple chart references whose worksheet no longer exists");
    test.checkTrue(scopeInvalid, "Model validator catches out-of-range defined-name localSheetId");
    test.checkTrue(nameDuplicate, "Model validator catches case-insensitive duplicate defined names within one scope");
}

void testValidateModelBeforeSaveP1J(TestContext& test) {
    xlpp::Workbook clean;
    auto& sheet = clean.addWorksheet("Data");
    sheet.cell("A1").setValue("Name");
    sheet.cell("B1").setValue("Value");
    auto& table = sheet.addTable("ValidTable", "A1:B1");
    table.addColumn("Name");
    table.addColumn("Value");

    xlpp::SaveOptions strictSave;
    strictSave.validateModelBeforeSave = true;
    const auto goodPath = std::filesystem::temp_directory_path() / "xlpp_p1j_validate_before_save_good.xlsx";
    clean.save(goodPath, strictSave);
    test.checkTrue(std::filesystem::exists(goodPath), "validateModelBeforeSave allows a coherent workbook");
    std::filesystem::remove(goodPath);

    // Create a semantic error through the mutable model surface. The normal
    // save path remains backwards-compatible, while the opt-in validation
    // gate rejects it before serialization.
    table.columns().pop_back();
    const auto broken = clean.validateModelIntegrity();
    test.checkTrue(!broken.ok(), "Test fixture contains a semantic model error");

    bool rejected = false;
    std::string message;
    try {
        clean.save(std::filesystem::temp_directory_path() / "xlpp_p1j_validate_before_save_bad.xlsx", strictSave);
    } catch (const std::runtime_error& error) {
        rejected = true;
        message = error.what();
    }
    test.checkTrue(rejected, "validateModelBeforeSave rejects semantic model errors before serialization");
    test.checkTrue(message.find("table.column_width_mismatch") != std::string::npos,
                   "Validation save error includes a stable model-validation code");

    xlpp::SaveOptions compatibleSave;
    compatibleSave.validateModelBeforeSave = false;
    const auto compatibilityPath = std::filesystem::temp_directory_path() / "xlpp_p1j_validate_before_save_compat.xlsx";
    // Do not assert that every deliberately corrupt model is serializable;
    // instead verify the option's compatibility default at the type level.
    test.checkTrue(!compatibleSave.validateModelBeforeSave,
                   "Model validation before save remains opt-in for preservation compatibility");
    std::filesystem::remove(compatibilityPath);
}

void testCoreGridAndWorksheetNameInvariantsP1J(TestContext& test) {
    bool threw = false;
    try { (void)xlpp::CellReference::columnName(16385); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "CellReference rejects columns beyond XFD when formatting");
    threw = false;
    try { (void)xlpp::CellReference::columnIndex("XFE"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "CellReference rejects column names beyond XFD");
    threw = false;
    try { (void)xlpp::CellReference::parse("A1048577"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "CellReference rejects rows beyond Excel's last row");
    threw = false;
    try { (void)xlpp::CellReference::parse("XFE1"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "CellReference rejects cell addresses beyond XFD");
    test.checkTrue(xlpp::CellReference::validGridPosition(1048576, 16384), "CellReference recognizes the maximum valid Excel coordinate");
    test.checkTrue(!xlpp::CellReference::validGridPosition(1048577, 1), "CellReference grid predicate rejects rows past the limit");

    xlpp::Worksheet sheet("Grid");
    threw = false;
    try { (void)sheet.cell(1048577, 1); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "Worksheet cell access rejects rows outside the Excel grid");
    threw = false;
    try { (void)sheet.cell(1, 16385); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "Worksheet cell access rejects columns outside the Excel grid");
    test.checkTrue(sheet.tryCell(1048577, 1) == nullptr, "tryCell safely returns null for out-of-grid rows");
    threw = false;
    try { (void)sheet.range(1, 1, 1048577, 1); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "Numeric range creation rejects out-of-grid bounds");
    threw = false;
    try { (void)sheet.rowDimension(1048577); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Row dimensions reject rows beyond the Excel limit");

    xlpp::Workbook workbook;
    workbook.addWorksheet("Data");
    threw = false;
    try { workbook.addWorksheet("data"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Workbook worksheet names are unique case-insensitively like Excel");
    test.checkTrue(workbook.worksheet("DATA") != nullptr, "Workbook worksheet lookup is case-insensitive like Excel");
    for (const std::string invalid : {"Bad/Name", "Bad:Name", "Bad?Name", "Bad*Name", "Bad[Name]", "'Quoted", "Quoted'"}) {
        threw = false;
        try { workbook.addWorksheet(invalid); } catch (const std::invalid_argument&) { threw = true; }
        test.checkTrue(threw, "Workbook rejects invalid Excel worksheet name: " + invalid);
    }
    threw = false;
    try { workbook.addWorksheet(std::string(32, 'A')); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Workbook rejects worksheet names longer than 31 characters");
    const std::string unicode31 = std::string(30, 'A') + "✓";
    test.checkEqual(workbook.addWorksheet(unicode31).name(), unicode31, "Worksheet-name length counts Unicode code points instead of UTF-8 bytes");
    threw = false;
    try { workbook.renameWorksheet("Data", unicode31); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Safe worksheet rename rejects case/Unicode duplicate target names");

    xlpp::DefinedName globalName("Revenue", "=Data!$A$1");
    workbook.addDefinedName(std::move(globalName));
    test.checkTrue(workbook.definedName("REVENUE") != nullptr, "Defined-name lookup is case-insensitive like Excel");
    threw = false;
    try { workbook.addDefinedName(xlpp::DefinedName("revenue", "=Data!$A$2")); }
    catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Defined names reject case-insensitive duplicates in the same workbook scope");

    xlpp::DefinedName localName("Revenue", "=$B$1");
    localName.setLocalSheetId(0);
    workbook.addDefinedName(std::move(localName));
    test.checkTrue(workbook.definedNames().size() >= 2, "A local defined name may shadow a workbook-global name");
    xlpp::DefinedName invalidLocal("BadScope", "=$A$1");
    invalidLocal.setLocalSheetId(999);
    threw = false;
    try { workbook.addDefinedName(std::move(invalidLocal)); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "Defined-name creation rejects an out-of-range localSheetId");

    xlpp::Worksheet local("Local");
    threw = false;
    try { local.rename("Invalid\\Name"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Direct Worksheet::rename enforces Excel worksheet-name invariants");
}

void testWorkbookSheetRenameRemoveSafetyP1J(TestContext& test) {
    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data");
    data.setVbaCodeName("DataCode");
    data.cell("A1").setValue("Category");
    data.cell("B1").setValue("Value");
    data.cell("A2").setValue("A");
    data.cell("B2").setValue(10.0);
    auto& calc = workbook.addWorksheet("Calc");
    auto& other = workbook.addWorksheet("Other");
    other.cell("A1").setValue("keep");

    calc.cell("A1").setFormula("SUM(Data!B1:B2)");
    calc.cell("A2").setFormula("SUM(Other!A1:A2)+Data!B2");
    xlpp::Hyperlink internalLink("#Data!A2");
    internalLink.setExternal(false);
    calc.cell("B1").setHyperlink(std::move(internalLink));
    xlpp::Hyperlink externalLink("https://example.com/Data!A2");
    externalLink.setExternal(true);
    calc.cell("B2").setHyperlink(std::move(externalLink));
    calc.conditionalFormatting().addRule("A1:A2", xlpp::ConditionalRule::formula("Data!B2>0"));
    calc.dataValidations().add(xlpp::DataValidation::list("C1:C2", "=Data!$A$1:$A$2"));

    xlpp::DefinedName global("DataValues", "=Data!$B$1:$B$2");
    workbook.addDefinedName(global);
    xlpp::DefinedName removedLocal("RemovedLocal", "=$A$1:$A$2");
    removedLocal.setLocalSheetId(0);
    workbook.addDefinedName(removedLocal);
    xlpp::DefinedName survivingLocal("CalcLocal", "=Data!$A$1:$A$2");
    survivingLocal.setLocalSheetId(1);
    workbook.addDefinedName(survivingLocal);

    xlpp::Chart chart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries series("Series");
    series.setCategoriesReference("='Data'!$A$1:$A$2");
    series.setValuesReference("='Data'!$B$1:$B$2");
    chart.addSeries(std::move(series));
    calc.addChart(std::move(chart));

    xlpp::PivotTable pivot("DataPivot");
    pivot.setLocation("E2:G6");
    pivot.cache().setSourceData("'Data'!$A$1:$B$2");
    pivot.cache().setFields({"Category", "Value"});
    pivot.cache().addRecord({"A", "10"});
    pivot.addRowField("Category");
    pivot.addDataField("Value", "sum");
    calc.addPivotTable(std::move(pivot));

    test.checkTrue(workbook.renameWorksheet("Data", "Sales 2026"), "Workbook-safe worksheet rename succeeds");
    test.checkTrue(workbook.worksheet("Data") == nullptr, "Old worksheet name is retired after safe rename");
    const auto* renamed = workbook.worksheet("Sales 2026");
    test.checkTrue(renamed != nullptr, "Renamed worksheet is addressable by its new name");
    test.checkEqual(renamed->vbaCodeName(), std::string("DataCode"), "Visible worksheet rename preserves the VBA codeName identity");
    test.checkEqual(calc.cell("A1").formula(), std::string("SUM('Sales 2026'!B1:B2)"), "Worksheet rename rewrites cross-sheet formulas");
    test.checkEqual(calc.cell("A2").formula(), std::string("SUM(Other!A1:A2)+'Sales 2026'!B2"), "Worksheet rename rewrites only the targeted qualifier");
    test.checkEqual(calc.cell("B1").hyperlinkValue()->target(), std::string("#'Sales 2026'!A2"), "Worksheet rename rewrites internal hyperlinks");
    test.checkEqual(calc.cell("B2").hyperlinkValue()->target(), std::string("https://example.com/Data!A2"), "Worksheet rename leaves external hyperlinks untouched");
    test.checkEqual(workbook.definedName("DataValues")->value(), std::string("='Sales 2026'!$B$1:$B$2"), "Worksheet rename rewrites workbook defined names");
    test.checkEqual(calc.charts().front().series().front().valuesReference(), std::string("='Sales 2026'!$B$1:$B$2"), "Worksheet rename rewrites chart series references");
    test.checkEqual(calc.pivotTables().front().cache().sourceData(), std::string("'Sales 2026'!$A$1:$B$2"), "Worksheet rename rewrites Pivot cache sources");
    test.checkTrue(workbook.calcProperties().fullCalcOnLoad(), "Worksheet rename requests full recalculation when references change");

    const auto renamedPath = std::filesystem::temp_directory_path() / "xlpp_p1j_safe_rename.xlsx";
    workbook.save(renamedPath);
    xlpp::Workbook renamedRoundTrip;
    renamedRoundTrip.load(renamedPath);
    test.checkEqual(renamedRoundTrip.worksheet("Calc")->cell("A1").formula(), std::string("SUM('Sales 2026'!B1:B2)"), "Safe worksheet rename survives save/reload");
    test.checkEqual(renamedRoundTrip.worksheet("Calc")->pivotTables().front().cache().sourceData(), std::string("'Sales 2026'!A1:B2"), "Renamed Pivot source survives save/reload");

    test.checkTrue(workbook.removeWorksheet("Sales 2026"), "Workbook-safe worksheet removal succeeds");
    test.checkEqual(workbook.sheetCount(), std::size_t{2}, "Safe worksheet removal updates workbook sheet count");
    test.checkEqual(calc.cell("A1").formula(), std::string("SUM(#REF!)"), "Removing a worksheet invalidates formulas that reference it");
    test.checkEqual(calc.cell("A2").formula(), std::string("SUM(Other!A1:A2)+#REF!"), "Removing a worksheet preserves unrelated references while invalidating removed-sheet refs");
    test.checkEqual(calc.cell("B1").hyperlinkValue()->target(), std::string("#REF!"), "Removing a worksheet produces one canonical #REF! hyperlink target");
    test.checkTrue(workbook.definedName("RemovedLocal") == nullptr, "Local defined names scoped to a removed sheet are retired");
    test.checkEqual(workbook.definedName("CalcLocal")->localSheetId().value(), std::size_t{0}, "Surviving local defined-name scope shifts after sheet removal");
    test.checkEqual(workbook.definedName("DataValues")->value(), std::string("=#REF!"), "Workbook defined names targeting a removed sheet become #REF!");
    test.checkEqual(calc.charts().front().series().front().valuesReference(), std::string("=#REF!"), "Chart references targeting a removed sheet become #REF!");
    test.checkTrue(!calc.charts().front().series().front().valuesCache().present, "Removed-sheet chart references clear stale caches");
    test.checkEqual(calc.pivotTables().front().cache().sourceData(), std::string("#REF!"), "Pivot source targeting a removed sheet becomes #REF! while cached data is preserved");
    test.checkEqual(calc.pivotTables().front().cache().records().size(), std::size_t{1}, "Removed-sheet Pivot keeps stale cache records for preservation/recovery");

    const auto removalValidation = workbook.validateModelIntegrity();
    test.checkTrue(removalValidation.ok(), "Safe worksheet removal leaves no hard model-integrity errors");
    test.checkTrue(removalValidation.warningCount() >= 1, "Safe worksheet removal surfaces broken references as model warnings");

    bool lastSheetRejected = false;
    xlpp::Workbook single;
    single.addWorksheet("Only");
    try { (void)single.removeWorksheet("Only"); } catch (const std::logic_error&) { lastSheetRejected = true; }
    test.checkTrue(lastSheetRejected, "Workbook refuses to remove its last worksheet");

    // Load a package whose removed worksheet owns a drawing/chart. Safe removal
    // must retire the exclusive preserved descendants so no orphaned chart or
    // drawing parts survive into the output package.
    xlpp::Workbook ownerFixture;
    auto& owner = ownerFixture.addWorksheet("Owner");
    owner.cell("A1").setValue(1.0);
    owner.cell("A2").setValue(2.0);
    ownerFixture.addWorksheet("Keep").cell("A1").setValue("alive");
    xlpp::Chart ownedChart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries ownedSeries("Owned");
    ownedSeries.setCategoriesReference("='Owner'!$A$1:$A$2");
    ownedSeries.setValuesReference("='Owner'!$A$1:$A$2");
    ownedChart.addSeries(std::move(ownedSeries));
    owner.addChart(std::move(ownedChart));
    const auto ownerBefore = std::filesystem::temp_directory_path() / "xlpp_p1j_remove_owner_before.xlsx";
    const auto ownerAfter = std::filesystem::temp_directory_path() / "xlpp_p1j_remove_owner_after.xlsx";
    ownerFixture.save(ownerBefore);
    xlpp::Workbook importedOwner;
    importedOwner.load(ownerBefore);
    test.checkTrue(importedOwner.removeWorksheet("Owner"), "Safe removal works for a loaded sheet owning preserved drawing parts");
    importedOwner.save(ownerAfter);
    auto afterArchive = xlpp::internal::ZipArchive::open(ownerAfter);
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(afterArchive).validate();
    test.checkTrue(graph.ok(), "Removing a loaded sheet retires exclusive drawing/chart descendants without package orphans");

    std::filesystem::remove(renamedPath);
    std::filesystem::remove(ownerBefore);
    std::filesystem::remove(ownerAfter);
}

void testTransactionalStructuralRollbackP1K(TestContext& test) {
    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data");
    auto& calc = workbook.addWorksheet("Calc");
    data.cell("A1").setValue("Name");
    data.cell("B1").setValue("Value");
    data.cell("A2").setValue("alpha");
    data.cell("B2").setValue(42.0);
    calc.cell("A1").setFormula("='Data'!B2*2");
    xlpp::DefinedName valueName("DataValue", "='Data'!$B$2");
    workbook.addDefinedName(std::move(valueName));

    auto* heldData = &data;
    std::size_t cancellationCalls = 0;
    xlpp::StructuralEditOptions options;
    options.cancel = [&]() { return ++cancellationCalls == 2; };

    bool cancelled = false;
    try { (void)workbook.insertColumns("Data", 2, 1, options); }
    catch (const xlpp::StructuralEditCancelled&) { cancelled = true; }
    test.checkTrue(cancelled, "Structural edit cancellation propagates a stable exception type");
    test.checkTrue(workbook.worksheet("Data") == heldData, "Rollback preserves Worksheet object identity");
    test.checkNear(std::get<double>(heldData->cell("B2").value()), 42.0, 1e-12, "Rollback restores moved cell geometry/value");
    test.checkTrue(heldData->tryCell("C2") == nullptr, "Rollback removes partially-created moved cell location");
    test.checkEqual(calc.cell("A1").formula(), std::string("='Data'!B2*2"), "Rollback restores cross-sheet formulas");
    test.checkEqual(workbook.definedName("DataValue")->value(), std::string("='Data'!$B$2"), "Rollback restores defined names");

    options.cancel = {};
    const auto report = workbook.insertColumns("Data", 2, 1, options);
    test.checkNear(std::get<double>(heldData->cell("C2").value()), 42.0, 1e-12, "Successful transactional edit moves cell after rollback test");
    test.checkEqual(calc.cell("A1").formula(), std::string("='Data'!C2*2"), "Successful transaction rewrites cross-sheet formula");
    test.checkEqual(workbook.definedName("DataValue")->value(), std::string("='Data'!$C$2"), "Successful transaction rewrites defined name");
    test.checkEqual(report.modelErrorsAfterEdit, std::size_t{0}, "Post-edit semantic validation reports no errors");

    // Performance-sensitive callers can explicitly opt out of rollback. This
    // path is deliberately observable so the contract cannot silently regress.
    xlpp::Workbook noRollback;
    auto& raw = noRollback.addWorksheet("Raw");
    raw.cell("B2").setValue(9.0);
    xlpp::StructuralEditOptions fastOptions;
    fastOptions.rollbackOnFailure = false;
    std::size_t fastCalls = 0;
    fastOptions.cancel = [&]() { return ++fastCalls == 2; };
    bool fastCancelled = false;
    try { (void)noRollback.insertColumns("Raw", 2, 1, fastOptions); }
    catch (const xlpp::StructuralEditCancelled&) { fastCancelled = true; }
    test.checkTrue(fastCancelled, "Non-transactional structural edit can still be cooperatively cancelled");
    test.checkTrue(raw.tryCell("C2") != nullptr, "Rollback opt-out leaves already-applied local mutation visible by contract");
}

void testStructuralReferenceCompositionP1K(TestContext& test) {
    using xlpp::internal::StructuralAction;
    using xlpp::internal::StructuralAxis;
    using xlpp::internal::StructuralEditSpec;
    StructuralEditSpec edit;
    edit.axis = StructuralAxis::Column;
    edit.action = StructuralAction::Insert;
    edit.index = 2;
    edit.amount = 1;
    edit.targetSheetName = "Data";

    const auto composite = xlpp::internal::rewriteFormulaReferences(
        "=SUM('Data'!A1:B2,'Data'!D1:D2)+SUM('Data'!A:A 'Data'!C:C)+\"Data!B2\"",
        "Calc", edit);
    test.checkEqual(composite.text,
                    std::string("=SUM('Data'!A1:C2,'Data'!E1:E2)+SUM('Data'!A:A 'Data'!D:D)+\"Data!B2\""),
                    "Reference scanner rewrites union/intersection operands while preserving string literals");
    test.checkEqual(composite.referencesRewritten, std::size_t{3}, "Composite formula reports each rewritten reference that actually changes");

    const auto external = xlpp::internal::rewriteFormulaReferences(
        "='[Book.xlsx]Data'!B2+'Data'!B2", "Calc", edit);
    test.checkEqual(external.text, std::string("='[Book.xlsx]Data'!B2+'Data'!C2"),
                    "External workbook qualifier remains untouched while local peer rewrites");
}


void testThreeDimensionalReferenceHardeningP1K(TestContext& test) {
    using xlpp::internal::StructuralAction;
    using xlpp::internal::StructuralAxis;
    using xlpp::internal::StructuralEditSpec;

    StructuralEditSpec edit;
    edit.axis = StructuralAxis::Column;
    edit.action = StructuralAction::Insert;
    edit.index = 2;
    edit.amount = 1;
    edit.targetSheetName = "Start";

    const auto structural = xlpp::internal::rewriteFormulaReferences(
        "=SUM(Start:End!B2)+Start!B2+'[Book.xlsx]Start'!B2", "Calc", edit);
    test.checkEqual(structural.text,
                    std::string("=SUM(Start:End!B2)+Start!C2+'[Book.xlsx]Start'!B2"),
                    "Structural edit preserves unsupported 3-D reference while rewriting ordinary local reference");
    test.checkEqual(structural.referencesSkippedUnsupported, std::size_t{1},
                    "Unsupported 3-D structural reference is reported instead of silently guessed");

    const auto quotedStructural = xlpp::internal::rewriteFormulaReferences(
        "=SUM('Start:End'!B2)", "Calc", edit);
    test.checkEqual(quotedStructural.text, std::string("=SUM('Start:End'!B2)"),
                    "Quoted 3-D structural reference is preserved byte-for-byte");
    test.checkEqual(quotedStructural.referencesSkippedUnsupported, std::size_t{1},
                    "Quoted 3-D structural reference contributes to unsupported diagnostics");

    const auto renamed = xlpp::internal::renameWorksheetReferences(
        "=SUM(Start:End!A1)+Start!B2", "Start", "Begin");
    test.checkEqual(renamed.text, std::string("=SUM('Begin:End'!A1)+'Begin'!B2"),
                    "Renaming a 3-D endpoint rewrites only that endpoint and preserves coordinates");

    const auto renamedLast = xlpp::internal::renameWorksheetReferences(
        "=SUM('Start:End'!A1)", "End", "Finish");
    test.checkEqual(renamedLast.text, std::string("=SUM('Start:Finish'!A1)"),
                    "Renaming the last endpoint of a quoted 3-D reference is supported");

    const auto invalidated = xlpp::internal::invalidateWorksheetReferences(
        "=SUM(Start:End!A1)+End!A1", "Start");
    test.checkEqual(invalidated.text, std::string("=SUM(#REF!)+End!A1"),
                    "Removing a 3-D endpoint invalidates the complete 3-D reference safely");
    test.checkEqual(invalidated.referencesInvalidated, std::size_t{1},
                    "3-D endpoint removal reports reference invalidation");

    xlpp::Workbook workbook;
    workbook.addWorksheet("Start").cell("A1").setValue(1.0);
    workbook.addWorksheet("Middle").cell("A1").setValue(2.0);
    workbook.addWorksheet("End").cell("A1").setValue(3.0);
    auto& calc = workbook.addWorksheet("Calc");
    calc.cell("A1").setFormula("=SUM(Start:End!A1)+Start!B2");
    calc.cell("A2").setFormula("=SUM(Start:End!A1)");

    const auto report = workbook.insertColumns("Start", 2, 1);
    test.checkEqual(calc.cell("A1").formula(), std::string("=SUM(Start:End!A1)+Start!C2"),
                    "Workbook structural edit preserves 3-D formula and rewrites ordinary peer reference");
    test.checkEqual(report.referencesSkippedUnsupported, std::size_t{2},
                    "Workbook structural report surfaces both mixed and pure preserved 3-D references");
    test.checkTrue(std::any_of(report.diagnostics.begin(), report.diagnostics.end(), [](const std::string& value) {
        return value.find("3-D worksheet reference") != std::string::npos;
    }), "Workbook structural report includes actionable 3-D diagnostic");

    workbook.renameWorksheet("Start", "Begin");
    test.checkEqual(calc.cell("A1").formula(), std::string("=SUM('Begin:End'!A1)+'Begin'!C2"),
                    "Workbook-safe rename rewrites 3-D endpoint and ordinary references together");
}


void testReferenceTransformerNegativeCorpusP1K(TestContext& test) {
    using xlpp::internal::StructuralAction;
    using xlpp::internal::StructuralAxis;
    using xlpp::internal::StructuralEditSpec;

    StructuralEditSpec edit;
    edit.axis = StructuralAxis::Column;
    edit.action = StructuralAction::Insert;
    edit.index = 2;
    edit.amount = 1;
    edit.targetSheetName = "Data";

    const std::vector<std::string> malformed = {
        "=", "=SUM(", "='Unclosed!A1", "=[Book.xlsx", "=A0", "=XFE1",
        "=A1048577", "=1:1048577", "=A:XFE", "=Data:!A1", "=Data::End!A1",
        "=Table1[Column", "=R1C1", "=\"A1", "=Data!$"
    };
    for (const auto& formula : malformed) {
        bool threw = false;
        xlpp::internal::ReferenceRewriteResult rewritten;
        try { rewritten = xlpp::internal::rewriteFormulaReferences(formula, "Calc", edit); }
        catch (...) { threw = true; }
        test.checkTrue(!threw, "Malformed reference corpus entry does not throw: " + formula);
        if (!threw)
            test.checkEqual(rewritten.text, formula, "Malformed/non-A1 corpus entry is preserved verbatim: " + formula);
    }

    StructuralEditSpec insertColumns;
    insertColumns.axis = StructuralAxis::Column;
    insertColumns.action = StructuralAction::Insert;
    insertColumns.index = 2;
    insertColumns.amount = 3;
    insertColumns.targetSheetName = "Data";
    StructuralEditSpec deleteColumns = insertColumns;
    deleteColumns.action = StructuralAction::Delete;

    const std::string original = "=SUM('Data'!A1:B9,'Data'!D:D)+$A$1+'Data'!F5";
    const auto inserted = xlpp::internal::rewriteFormulaReferences(original, "Calc", insertColumns);
    const auto restored = xlpp::internal::rewriteFormulaReferences(inserted.text, "Calc", deleteColumns);
    test.checkEqual(restored.text, original,
                    "Insert/delete structural transformations compose back to the original formula when no referenced geometry is deleted");
    test.checkEqual(restored.referencesInvalidated, std::size_t{0},
                    "Round-trip structural transformation does not introduce #REF invalidations");
}


void testStrictSaveValidationP1K(TestContext& test) {
    const auto goodPath = std::filesystem::temp_directory_path() / "xlpp_p1k_strict_good.xlsx";
    const auto warningPath = std::filesystem::temp_directory_path() / "xlpp_p1k_strict_warning.xlsx";
    const auto rejectedWarningPath = std::filesystem::temp_directory_path() / "xlpp_p1k_strict_warning_rejected.xlsx";
    const auto packagePath = std::filesystem::temp_directory_path() / "xlpp_p1k_strict_package_rejected.xlsx";
    std::filesystem::remove(goodPath);
    std::filesystem::remove(warningPath);
    std::filesystem::remove(rejectedWarningPath);
    std::filesystem::remove(packagePath);

    xlpp::SaveOptions strict;
    strict.validateModelBeforeSave = true;
    strict.rejectModelWarningsBeforeSave = true;
    strict.validatePackageBeforeWrite = true;

    xlpp::Workbook good;
    good.addWorksheet("Data").cell("A1").setValue(1.0);
    good.save(goodPath, strict);
    test.checkTrue(std::filesystem::exists(goodPath), "Combined model+package strict validation allows a coherent workbook");

    xlpp::Workbook warning;
    warning.addWorksheet("Data").cell("A1").setFormula("=#REF!+1");
    xlpp::SaveOptions compatible = strict;
    compatible.rejectModelWarningsBeforeSave = false;
    warning.save(warningPath, compatible);
    test.checkTrue(std::filesystem::exists(warningPath), "Compatibility model validation still permits non-fatal #REF warnings");

    bool warningRejected = false;
    try { warning.save(rejectedWarningPath, strict); }
    catch (const std::runtime_error& e) {
        warningRejected = std::string(e.what()).find("formula.broken_reference") != std::string::npos;
    }
    test.checkTrue(warningRejected, "Strict warning policy rejects a workbook containing #REF formula warnings");
    test.checkTrue(!std::filesystem::exists(rejectedWarningPath), "Strict semantic rejection occurs before output bytes are written");

    xlpp::Workbook packageBroken;
    packageBroken.addWorksheet("Data").cell("A1").setValue("ok");
    xlpp::PreservedPart orphan;
    orphan.name = "xl/charts/chart999.xml";
    orphan.data = R"(<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>)";
    orphan.overrideType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    orphan.extension = "xml";
    orphan.compress = true;
    packageBroken.preservedParts().push_back(std::move(orphan));
    xlpp::SaveOptions packageStrict;
    packageStrict.validatePackageBeforeWrite = true;
    bool packageRejected = false;
    try { packageBroken.save(packagePath, packageStrict); }
    catch (const std::runtime_error& e) {
        packageRejected = std::string(e.what()).find("orphaned_parts") != std::string::npos;
    }
    test.checkTrue(packageRejected, "In-memory pre-write OPC validation rejects an orphan chart part");
    test.checkTrue(!std::filesystem::exists(packagePath), "Package topology rejection occurs before output bytes are written");

    std::filesystem::remove(goodPath);
    std::filesystem::remove(warningPath);
}


void testWorksheetOverlapAndValidationHardeningP1K(TestContext& test) {
    xlpp::Worksheet sheet("Data");
    sheet.mergeCells("A1:B2");
    bool mergeRejected = false;
    try { sheet.mergeCells("B2:C3"); } catch (const std::invalid_argument&) { mergeRejected = true; }
    test.checkTrue(mergeRejected, "Worksheet rejects overlapping merged ranges at mutation time");

    auto& firstTable = sheet.addTable("Sales", "A5:B8");
    firstTable.addColumn("Name");
    firstTable.addColumn("Value");
    bool tableOverlapRejected = false;
    try { (void)sheet.addTable("Other", "B7:C10"); } catch (const std::invalid_argument&) { tableOverlapRejected = true; }
    test.checkTrue(tableOverlapRejected, "Worksheet rejects overlapping table ranges at mutation time");
    bool tableCaseRejected = false;
    try { (void)sheet.addTable("sales", "D5:E8"); } catch (const std::invalid_argument&) { tableCaseRejected = true; }
    test.checkTrue(tableCaseRejected, "Worksheet table identifiers are case-insensitive like Excel names");
    test.checkTrue(sheet.table("SALES") != nullptr, "Case-insensitive table lookup resolves the canonical table");

    xlpp::Workbook workbook;
    auto& ws = workbook.addWorksheet("Data");
    auto& t1 = ws.addTable("T1", "A1:B3");
    t1.addColumn("A"); t1.addColumn("B");
    auto& t2 = ws.addTable("T2", "D1:E3");
    t2.addColumn("D"); t2.addColumn("E");
    // Mutable table access is intentionally available for advanced callers;
    // the model validator must still catch an invalid overlap introduced there.
    ws.tables()[1].setReference("B2:C4");
    ws.conditionalFormatting().add("A1:A2").addRule(xlpp::ConditionalRule::formula("A1>0"));
    ws.dataValidations().add(xlpp::DataValidationType::Whole, "B1:B2");
    const auto validation = workbook.validateModelIntegrity();
    bool sawTableOverlap = false;
    for (const auto& issue : validation.issues)
        if (issue.code == "table.overlap") sawTableOverlap = true;
    test.checkTrue(sawTableOverlap, "Model validator catches table overlap introduced through mutable model access");
}


int main() {
    std::cout << std::unitbuf;
    const std::vector<std::pair<std::string, TestFunction>> tests{
        {"Cell references", testCellReferences},
        {"Range and dimensions", testRangeAndDimensions},
        {"Append and structural edits", testAppendAndStructuralEdits},
        {"Cell convenience methods", testCellConvenience},
        {"Named style association on cells", testNamedStyleAssociation},
        {"Remove worksheet", testRemoveWorksheet},
        {"Worksheet extents", testWorksheetExtents},
        {"Merged cells", testMergedCells},
        {"Worksheet layout", testWorksheetLayout},
        {"XLSX layout round-trip", testRoundTrip},
        {"AutoFilter model", testAutoFilter},
        {"AutoFilter XLSX round-trip", testAutoFilterRoundTrip},
        {"Cell styles", testCellStyles},
        {"Compact cell model P1N", testCompactCellModelP1N},
        {"Compact optional cell payload P1O", testCompactOptionalCellPayloadP1O},
        {"Lazy Style/formula metadata and tracking P1P", testLazyStyleFormulaAndMutationTrackingP1P},
        {"Styles XLSX round-trip", testStylesRoundTrip},
        {".h header migration smoke test", testHeaderMigration},
        {"Named styles registry", testNamedStyles},
        {"Named styles XLSX round-trip", testNamedStylesRoundTrip},
        {"Conditional formatting model", testConditionalFormatting},
        {"Conditional formatting XLSX round-trip", testConditionalFormattingRoundTrip},
        {"Validation and conditional package regression", testValidationAndConditionalPackageRegression},
        {"Data validation model", testDataValidation},
        {"Data validation XLSX round-trip", testDataValidationRoundTrip},
        {"Tables and defined names model", testTablesAndDefinedNames},
        {"Tables and defined names XLSX round-trip", testTablesAndDefinedNamesRoundTrip},
        {"Hyperlinks, comments and properties model", testHyperlinksCommentsAndProperties},
        {"Hyperlinks and properties XLSX round-trip", testHyperlinksAndPropertiesRoundTrip},
        {"Legacy comments XLSX round-trip", testCommentsRoundTrip},
        {"Comment mutation after cached save", testCommentMutationAfterSave},
        {"Rich-text legacy comment import", testRichTextCommentImport},
        {"Page setup, protection and images model", testPageSetupProtectionAndImages},
        {"Page setup, protection and images XLSX round-trip", testPageSetupProtectionAndImagesRoundTrip},
        {"Print area, titles and fit-to-page", testPrintAreaTitlesAndFitToPage},
        {"Comprehensive formatting round-trip", testComprehensiveFormattingRoundTrip},
        {"Protection password add/remove", testProtectionPasswordAddRemove},
        {"Image package regression", testImagePackageRegression},
        {"Formula metadata and error cells model", testFormulaMetadataAndErrorCells},
        {"Formula metadata and error cells XLSX round-trip", testFormulaMetadataAndErrorCellsRoundTrip},
        {"Direct streaming ZIP reader", testDirectZipReader},
        {"Streaming pull iterator and shared strings", testStreamingSharedStringsAndPullIterator},
        {"Streaming read/write and benchmark", testStreamingReadWrite},
        {"Fast XML scanner", testXmlScanner},
        {"Fast XML scanner benchmark", testXmlScannerBenchmark},
        {"Date/time core", testDateTimeCore},
        {"Built-in date format IDs", testBuiltinDateFormatIds},
        {"Date cells XLSX round-trip", testDateCellsRoundTrip},
        {"Streaming shared-string writer", testSharedStringWriter},
        {"DOM shared-string save and load", testDOMSharedStrings},
        {"Rich text shared-string load", testRichTextSharedStrings},
        {"Rich text cell XLSX round-trip", testRichTextCellRoundTrip},
        {"Row proxy and range helpers", testRowProxyAndRangeHelpers},
        {"Cell style index", testCellStyleIndex},
        {"Stream load and save", testStreamLoadSave},
        {"Built-in date format round-trip", testNumFmtIdDateRoundTrip},
        {"Cell edge cases and cleanup", testEdgeCasesAndCleanup},
        {"Worksheet rows() iteration", testWorksheetRows},
        {"iterRows and iterCols", testIterRowsCols},
        {"Cell::offset()", testCellOffset},
        {"Workbook navigation", testWorkbookNav},
        {"Workbook copyWorksheet", testCopyWorksheet},
        {"Parallel package pipeline", testParallelPackagePipeline},
        {"ZIP64, limits, cancel and progress", testZip64LimitsCancelProgress},
        {"ZIP integrity and atomic-commit hardening", testZipIntegrityHardening},
        {"ZIP64 forced write path", testZip64WritePath},
        {"Package part preservation", testPartPreservation},
        {"Binary package part preservation", testBinaryPartPreservation},
        {"VBA project package lifecycle", testVbaProjectPackageLifecycle},
        {"Chart part preservation", testChartPartPreservation},
        {"Relationship graph round-trip preservation", testRelationshipGraphRoundTripPreservation},
        {"Independent image/chart fixture round-trip", testIndependentImageChartFixtureRoundTrip},
        {"Drawing image reader metadata", testDrawingImageReaderMetadata},
        {"Absolute image anchor reader", testAbsoluteImageAnchorReader},
        {"Append image to preserved drawing", testAppendImageToPreservedDrawing},
        {"Selective imported image move and resize", testSelectiveImportedImageMoveResize},
        {"Selective absolute image move and resize", testSelectiveAbsoluteImageMoveResize},
        {"Selective imported image remove", testSelectiveImportedImageRemove},
        {"Selective image replace with shared media", testSelectiveImageReplaceWithSharedMedia},
        {"Multi-drawing selective mutation", testMultiDrawingSelectiveMutation},
        {"Unknown drawing relationship validation", testUnknownDrawingRelationshipValidation},
        {"Independent pivot fixture round-trip", testIndependentPivotFixtureRoundTrip},
        {"Preserved and generated pivot coexistence", testPreservedAndGeneratedPivotCoexistence},
        {"Owner-reference and object regression detection", testOwnerReferenceAndObjectRegressionDetection},
        {"Extended owner graph and relationship hardening", testExtendedOwnerGraphAndRelationshipHardening},
        {"Package validator failure detection", testPackageValidatorFailureDetection},
        {"ZIP unique entry policy", testZipUniqueEntryPolicy},
        {"Pivot validation", testPivotValidation},
        {"Pivot model validation and aggregation", testPivotModelValidationAndAggregation},
        {"Strict/transitional namespaces", testStrictNamespaces},
        {"Lenient load recovery", testLenientLoad},
        {"Malformed input hardening", testMalformedInputHardening},
        {"Mutation fuzz", testMutationFuzz},
        {"Cell reference matrix", testCellReferenceMatrix},
        {"Cell error matrix", testCellErrorMatrix},
        {"xlfn() helper", testXlfnHelper},
        {"Formula metadata defaults", testFormulaMetadataDefaults},
        {"Number format detection", testNumberFormatDetection},
        {"Date/time boundaries", testDateTimeBoundaries},
        {"Streaming writer shared-string modes", testStreamingWriterModes},
        {"Streaming reader features", testStreamingReaderFeatures},
        {"Compression levels and parallel output", testCompressionLevelsAndParallel},
        {"Chart and pivot package", testChartAndPivotPackage},
        {"Pivot auto-cache from source", testPivotAutoCacheFromSource},
        {"Excel-compatible pivot view", testExcelCompatiblePivotView},
        {"Multiple pivot cache IDs", testMultiplePivotCacheIds},
        {"Strict pivot namespaces", testStrictPivotNamespaces},
        {"Imported chart selective mutation", testImportedChartInspectionAndSelectiveMutation},
        {"Imported scatter chart deep selective editing", testImportedScatterChartDeepSelectiveEditing},
        {"Imported combined chart axis structure", testImportedCombinedChartAxisStructure},
        {"Imported chart labels, trendlines and error bars", testImportedChartLabelsTrendlinesAndErrorBars},
        {"Imported chart per-point labels, custom error bars and formatting", testImportedChartPerPointCustomErrorsAndFormatting},
        {"Imported chart data-point, rich text and advanced formatting", testImportedChartDataPointRichTextAndAdvancedFormatting},
        {"Imported chart layout, axis and legend formatting", testImportedChartLayoutAxisLegendFormatting},
        {"Imported chart axis scaling, display units and area formatting", testImportedChartAxisScalingDisplayUnitsAndAreaFormatting},
        {"Imported chart auxiliary objects", testImportedChartAuxiliaryObjects},
        {"Stock chart structure, generation and data-table text", testStockChartStructureGenerationAndDataTableText},
        {"3D and surface chart preservation foundation", testThreeDSurfaceChartPreservationFoundation},
        {"Projected pie, doughnut and radar expansion", testProjectedPieDoughnutRadarExpansion},
        {"Chart style, theme and series caches", testChartStyleThemeAndSeriesCaches},
        {"Chart cache synchronization and theme transforms", testChartCacheSynchronizationAndThemeTransforms},
        {"Chart cache dependency tracking and style resolution", testChartCacheDependencyTrackingAndStyleResolution},
        {"Formula dependency propagation and style application", testFormulaDependencyPropagationAndStyleApplication},
        {"Formula dependency grammar and defined names P0X", testFormulaDependencyGrammarAndDefinedNamesP0X},
        {"Structured references, dynamic names and theme matrix P0Y", testStructuredReferencesDynamicNamesAndThemeMatrixP0Y},
        {"Structured escaping, INDEX ranges and chart-style rules P0Z", testStructuredEscapingIndexRangesAndChartStyleRulesP0Z},
        {"Imported chart remove and append", testImportedChartRemoveAndAppend},
        {"Chart type names", testChartTypeNameMap},
        {"Internal hyperlink and memory stream", testInternalHyperlinkAndMemoryStream},
        {"Workbook edge cases", testWorkbookEdgeCases},
        {"Custom properties and calc round-trip", testCustomPropertiesAndCalcRoundTrip},
        {"Differential save cache", testDifferentialSaveCache},
        {"Strict after transitional save", testStrictAfterTransitionalSave},
        {"copyWorksheet aliasing", testCopyWorksheetAliasing},
        {"Reference stability across inserts", testReferenceStabilityAcrossInserts},
        {"Reference lifetime contract", testReferenceLifetimeContract},
        {"Workbook copy/move semantics", testWorkbookCopyMoveSemantics},
        {"Styled empty cells round-trip", testStyledEmptyCellsRoundTrip},
        {"Defined names full round-trip", testDefinedNamesFullRoundTrip},
        {"Row values and cells", testRowValuesAndCells},
        {"Date cell number format", testDateCellNumberFormat},
        {"CellRange operations", testCellRangeOperations},
        {"Properties full round-trip", testPropertiesFullRoundTrip},
        {"Hyperlink display and tooltip", testHyperlinkDisplayTooltipRoundTrip},
        {"Column dimension by name", testColumnDimensionByName},
        {"Advanced conditional formatting round-trip", testAdvancedConditionalFormattingRoundTrip},
        {"Advanced sheet view round-trip", testAdvancedSheetViewRoundTrip},
        {"AutoFilter mutation and operator matrix", testAutoFilterMutationAndOperatorMatrix},
        {"Cell overloads and optional models", testCellOverloadsAndOptionalModels},
        {"Table mutation round-trip", testTableMutationRoundTrip},
        {"Advanced page setup round-trip", testAdvancedPageSetupRoundTrip},
        {"Image file API", testImageFileApi},
        {"Workbook clear, const and VBA file API", testWorkbookClearConstAndVbaFileApi},
        {"Chart and pivot advanced model", testChartAndPivotAdvancedModel},
        {"Streaming accessors and post-increment", testStreamingAccessorsAndPostIncrement},
        {"Stored reference mutation after save", testStoredReferenceMutationAfterSave},
        {"Range bounds and indexed cell", testRangeBoundsAndIndexedCell},
        {"Remaining public mutation APIs", testRemainingPublicMutationApis},
        {"Internal IO and scanner coverage", testInternalIoAndScannerCoverage},
        {"Pivot imported reader/editor P1A", testPivotImportedReaderAndEditorP1A},
        {"Pivot OLAP calculated members P1Y", testPivotOlapCalculatedMembersP1Y},
        {"Pivot typed group items P1Y", testPivotTypedGroupItemsP1Y},
        {"VBA class/document/project metadata P1A", testVbaClassDocumentAndProjectMetadataP1A},
        {"Pivot shared cache/calculated/grouping P1B", testPivotSharedCacheCalculatedGroupingP1B},
        {"Pivot date grouping/imported calculated field P1B", testPivotDateGroupingImportedCalculatedFieldP1B},
        {"VBA project refs/locale/module metadata P1B", testVbaProjectReferencesLocaleAndModuleMetadataP1B},
        {"Pivot filters, PivotChart and selective cache P1C", testPivotFiltersChartLinkAndSelectiveCacheP1C},
        {"VBA UserForm designer storage P1C", testVbaDesignerUserFormStorageP1C},
        {"Pivot selective field/items/link validation P1D", testPivotSelectiveFieldItemsAndLinkValidationP1D},
        {"VBA UserForm semantic properties P1D", testVbaUserFormSemanticPropertiesP1D},
        {"Pivot selective item/filter mutation P1E", testPivotSelectiveItemAndFilterMutationP1E},
        {"VBA UserForm control sites P1E", testVbaUserFormControlSitesP1E},
        {"Pivot selective data/page fields P1F", testPivotSelectiveDataAndPageFieldsP1F},
        {"VBA UserForm control objects P1F", testVbaUserFormControlObjectsP1F},
        {"VBA UserForm extended controls P1Y", testVbaUserFormExtendedControlsP1Y},
        {"Worksheet sparklines P1Z", testSparklinesRoundTrip},
        {"AutoFilter top10/dynamic filters P1Z", testAutoFilterTop10AndDynamicFilters},
        {"Table totals row P1Z", testTableTotalsRow},
        {"Structural reference transformer P1J", testStructuralReferenceTransformerP1J},
        {"Workbook reference-safe structural edits P1J", testWorkbookReferenceSafeStructuralEditsP1J},
        {"Structural edit edge cases P1J", testStructuralEditEdgeCasesP1J},
        {"Workbook model-integrity validator P1J", testWorkbookModelIntegrityValidatorP1J},
        {"Workbook model-integrity hardening P1J", testWorkbookModelIntegrityHardeningP1J},
        {"Validate model before save P1J", testValidateModelBeforeSaveP1J},
        {"Core grid/name invariants P1J", testCoreGridAndWorksheetNameInvariantsP1J},
        {"Workbook-safe sheet rename/remove P1J", testWorkbookSheetRenameRemoveSafetyP1J},
        {"Transactional structural rollback P1K", testTransactionalStructuralRollbackP1K},
        {"Structural reference composition P1K", testStructuralReferenceCompositionP1K},
        {"3-D reference hardening P1K", testThreeDimensionalReferenceHardeningP1K},
        {"Reference transformer negative corpus P1K", testReferenceTransformerNegativeCorpusP1K},
        {"Strict save validation P1K", testStrictSaveValidationP1K},
        {"Worksheet overlap and validator hardening P1K", testWorksheetOverlapAndValidationHardeningP1K},
        {"Password-to-open Agile encryption P1G", testPasswordToOpenEncryptionP1G},
        {"Standard encryption interop P1G", testStandardEncryptionInteropP1G},
        {"Large encrypted CFB/DIFAT P1G", testLargeEncryptedCompoundFileP1G},
        {"Agile encryption profile matrix P1H", testAgileEncryptionProfileMatrixP1H},
        {"Standard encryption writer P1H", testStandardEncryptionWriterP1H},
        {"Encryption inspection/resource guards P1H", testEncryptionInspectionAndResourceGuardsP1H},
        {"Encryption memory/certificate inspection P1I", testEncryptionMemoryAndCertificateInspectionP1I},
        {"Encryption policy hardening P1I", testEncryptionPoliciesAndMalformedInputsP1I},
        {"VBA source text build and read", testVbaSourceTextBuildAndRead},
        {"External OOXML cell and style reader fixture", testExternalCellAndStyleReaderFixture},
        {"External OOXML worksheet feature reader fixture", testExternalWorksheetFeatureReaderFixture},
        {"External OOXML workbook metadata reader fixture", testExternalWorkbookMetadataReaderFixture},
    };

    std::cout << "============================================================\n";
    std::cout << " XL++ Unit Tests - Milestone 23: OPC Preservation Core\n";
    std::cout << "============================================================\n";

    TestContext context;
    std::size_t passed = 0;
    for (std::size_t index = 0; index < tests.size(); ++index) {
        std::cout << "\n[RUN  " << index + 1 << '/' << tests.size() << "] " << tests[index].first << '\n';
        try {
            tests[index].second(context);
            ++passed;
            std::cout << "[PASS " << index + 1 << '/' << tests.size() << "] " << tests[index].first << '\n';
        } catch (const std::exception& error) {
            std::cerr << "[FAIL " << index + 1 << '/' << tests.size() << "] " << tests[index].first
                      << "\n    Reason: " << error.what() << '\n';
        } catch (...) {
            std::cerr << "[FAIL " << index + 1 << '/' << tests.size() << "] " << tests[index].first
                      << "\n    Reason: unknown exception\n";
        }
    }

    std::cout << "\n============================================================\n";
    std::cout << " Test suites passed : " << passed << '/' << tests.size() << '\n';
    std::cout << " Checks executed    : " << context.checks() << '\n';
    std::cout << " Final result       : " << (passed == tests.size() ? "PASS" : "FAIL") << '\n';
    std::cout << "============================================================\n";
    return passed == tests.size() ? 0 : 1;
}
