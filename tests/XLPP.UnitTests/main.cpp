#include <XLPP/XLPP.h>
#include "Packaging/ZipArchive.h"
#include "Packaging/ZipArchiveReader.h"
#include "XML/XmlScanner.h"
#include "XML/XmlUtilities.h"
#include <array>
#include <chrono>
#include <cmath>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iostream>
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
        z.add("xl/worksheets/sheet1.xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>)");
        // Rich text shared string: two <r> elements that should be concatenated
        z.add("xl/sharedStrings.xml",
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><r><t>Hello </t></r><r><t>World</t></r></si><si><t>plain</t></si></sst>)");
        z.save(path);
    }
    xlpp::Workbook loaded;
    loaded.load(path);
    test.checkEqual(std::get<std::string>(loaded.worksheet("Rich")->cell("A1").value()), std::string("Hello World"), "Rich text concatenated on DOM load");
    test.checkEqual(std::get<std::string>(loaded.worksheet("Rich")->cell("A2").value()), std::string("plain"), "Plain shared string next to rich text");
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
        z.add("xl/worksheets/sheet1.xml",
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
        z.add("[Content_Types].xml", ct);
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
        z.add("xl/worksheets/sheet1.xml",
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
        z.add("xl/worksheets/sheet1.xml",
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
        z.add("xl/worksheets/sheet1.xml",
              sheetXml("<cols><col min=\"5\" max=\"2\" width=\"8\"/></cols>"), false);
        z.save(crafted);
        bool threw = false;
        try { loads(false); } catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Descending column range rejected");
    }
    // Bad numeric cell values (number without t, overflow, non-finite) never crash.
    {
        auto z = xlpp::internal::ZipArchive::open(wbPath);
        z.add("xl/worksheets/sheet1.xml",
              sheetXml("<sheetData><row r=\"1\"><c r=\"A1\"><v>1e400</v></c><c r=\"B1\"><v>99999999999999999999</v></c></row></sheetData>"), false);
        z.save(crafted);
        auto w = loads(true);
        test.checkTrue(w.diagnostics().hadErrors(), "Numeric overflow recorded in lenient mode");
        test.checkEqual(w.worksheets().size(), std::size_t{1}, "Sheet slot survives numeric overflow");
    }
    // Unclosed/invalid XML structures are tolerated (empty result), not a crash.
    {
        auto z = xlpp::internal::ZipArchive::open(wbPath);
        z.add("xl/worksheets/sheet1.xml",
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
        reader.forEachRow("Data", [&](std::size_t row, const xlpp::StreamingRow& cells) {
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

void testChartAndPivotPackage(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m21_chart_pivot.xlsx";
    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("Charts");
    sheet.append({std::string("Q1"), 10.0});
    sheet.append({std::string("Q2"), 20.0});

    xlpp::Chart chart(xlpp::Chart::Type::Bar);
    chart.setTitle("Sales");
    chart.setXAxisTitle("Quarter");
    chart.setYAxisTitle("Units");
    chart.setShowLegend(true);
    chart.setLegendPosition("b");
    auto& series = chart.addSeries(xlpp::ChartSeries("Units"));
    series.reference("Charts", "$B$2:$B$3");
    series.categories("Charts", "$A$2:$A$3");
    sheet.addChart(chart);

    xlpp::PivotTable pivot("SalesPivot");
    pivot.setLocation("D1");
    pivot.cache().setCacheId(1);
    pivot.cache().setSourceData("'Charts'!$A$1:$B$3");
    pivot.addRowField("Quarter");
    pivot.addColumnField("Units");
    pivot.addDataField();
    sheet.addPivotTable(std::move(pivot));

    wb.save(path);
    test.checkTrue(std::filesystem::exists(path), "Chart/pivot workbook saved");

    auto z = xlpp::internal::ZipArchive::open(path);
    test.checkTrue(z.contains("xl/charts/chart1.xml"), "Chart part written");
    const auto chartXml = z.get("xl/charts/chart1.xml");
    test.checkTrue(chartXml.find("barChart") != std::string::npos, "Bar chart type in part");
    test.checkTrue(chartXml.find("Sales") != std::string::npos, "Chart title in part");
    test.checkTrue(z.contains("xl/pivotTables/pivotTable1.xml"), "Pivot part written");
    test.checkTrue(z.contains("xl/pivotCache/pivotCacheDefinition1.xml"), "Pivot cache written");
    test.checkTrue(z.contains("xl/drawings/drawing1.xml"), "Drawing part written for chart");

    {
        xlpp::Workbook loaded;
        loaded.load(path);
        test.checkTrue(loaded.worksheet("Charts") != nullptr, "Chart workbook reloads");
    }
    std::filesystem::remove(path);
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
        {"Styles XLSX round-trip", testStylesRoundTrip},
        {".h header migration smoke test", testHeaderMigration},
        {"Named styles registry", testNamedStyles},
        {"Named styles XLSX round-trip", testNamedStylesRoundTrip},
        {"Conditional formatting model", testConditionalFormatting},
        {"Conditional formatting XLSX round-trip", testConditionalFormattingRoundTrip},
        {"Data validation model", testDataValidation},
        {"Data validation XLSX round-trip", testDataValidationRoundTrip},
        {"Tables and defined names model", testTablesAndDefinedNames},
        {"Tables and defined names XLSX round-trip", testTablesAndDefinedNamesRoundTrip},
        {"Hyperlinks, comments and properties model", testHyperlinksCommentsAndProperties},
        {"Hyperlinks and properties XLSX round-trip", testHyperlinksAndPropertiesRoundTrip},
        {"Legacy comments XLSX round-trip", testCommentsRoundTrip},
        {"Page setup, protection and images model", testPageSetupProtectionAndImages},
        {"Page setup, protection and images XLSX round-trip", testPageSetupProtectionAndImagesRoundTrip},
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
        {"ZIP64 forced write path", testZip64WritePath},
        {"Package part preservation", testPartPreservation},
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
    };

    std::cout << "============================================================\n";
    std::cout << " XL++ Unit Tests - Milestone 20: Core Compatibility Completion\n";
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
