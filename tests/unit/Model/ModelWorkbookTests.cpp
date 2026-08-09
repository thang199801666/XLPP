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

namespace {
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
    auto& first = wb.addWorksheet("First");
    wb.addWorksheet("Second");
    auto& third = wb.addWorksheet("Third");
    first.cell("A1").setFormula("=Second!A1");
    third.cell("A1").setFormula("='Second'!B2+\"Second!B2\"");
    third.cell("A2").setFormula("='[Other.xlsx]Second'!A1");
    wb.addDefinedName(xlpp::DefinedName("RemovedSheetRange", "Second!$A$1:$B$2"));
    test.checkEqual(wb.worksheets().size(), std::size_t{3}, "Three sheets added");
    test.checkTrue(wb.removeWorksheet("Second"), "removeWorksheet returns true for existing sheet");
    test.checkEqual(wb.worksheets().size(), std::size_t{2}, "Two sheets remain");
    test.checkTrue(wb.worksheet("First") != nullptr, "First sheet still present");
    test.checkTrue(wb.worksheet("Second") == nullptr, "Second sheet removed");
    auto* firstAfterRemoval = wb.worksheet("First");
    auto* thirdAfterRemoval = wb.worksheet("Third");
    test.checkEqual(firstAfterRemoval->cell("A1").formula(), std::string("=#REF!A1"),
                    "Removing a sheet invalidates surviving cross-sheet formulas");
    test.checkEqual(thirdAfterRemoval->cell("A1").formula(), std::string("=#REF!B2+\"Second!B2\""),
                    "Sheet removal preserves string literals while invalidating qualifiers");
    test.checkEqual(thirdAfterRemoval->cell("A2").formula(), std::string("='[Other.xlsx]Second'!A1"),
                    "Sheet removal preserves external-workbook references");
    test.checkEqual(wb.definedName("RemovedSheetRange")->value(), std::string("#REF!$A$1:$B$2"),
                    "Sheet removal invalidates workbook defined-name dependencies");
    test.checkTrue(wb.calcProperties().fullCalcOnLoad(), "Sheet removal requests full host recalculation");
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


void testStableModelHandles(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Stable");

    auto& namedStyle = workbook.addNamedStyle(xlpp::NamedStyle("Primary"));
    auto* namedStyleAddress = &namedStyle;
    for (int i = 0; i < 256; ++i)
        workbook.addNamedStyle(xlpp::NamedStyle("Style" + std::to_string(i)));
    test.checkTrue(workbook.namedStyle("Primary") == namedStyleAddress,
                   "Named-style handle survives registry growth");
    namedStyle.style().font().setBold(true);
    test.checkTrue(workbook.namedStyle("Primary")->style().font().bold(),
                   "Named-style handle remains writable after registry growth");

    auto& definedName = workbook.addDefinedName(xlpp::DefinedName("RootName", "Stable!$A$1"));
    auto* definedNameAddress = &definedName;
    for (int i = 0; i < 256; ++i)
        workbook.addDefinedName(xlpp::DefinedName("Name" + std::to_string(i), "Stable!$A$1"));
    test.checkTrue(workbook.definedName("RootName") == definedNameAddress,
                   "Defined-name handle survives registry growth");
    definedName.setComment("still-live");
    test.checkEqual(workbook.definedName("RootName")->comment(), std::string("still-live"),
                    "Defined-name handle remains writable after registry growth");

    auto& table = sheet.addTable("RootTable", "A1:B2");
    auto* tableAddress = &table;
    for (int i = 0; i < 128; ++i)
        sheet.addTable("Table" + std::to_string(i), "A1:B2");
    test.checkTrue(&sheet.tables().front() == tableAddress, "Table handle survives worksheet table growth");

    auto& column = table.addColumn("First");
    auto* columnAddress = &column;
    for (int i = 0; i < 128; ++i) table.addColumn("C" + std::to_string(i));
    test.checkTrue(&table.columns().front() == columnAddress, "Table-column handle survives column growth");

    auto& chart = sheet.addChart(xlpp::Chart(xlpp::Chart::Type::Line));
    auto* chartAddress = &chart;
    for (int i = 0; i < 128; ++i) sheet.addChart(xlpp::Chart(xlpp::Chart::Type::Line));
    test.checkTrue(&sheet.charts().front() == chartAddress, "Chart handle survives worksheet chart growth");

    auto& series = chart.addSeries(xlpp::ChartSeries{});
    auto* seriesAddress = &series;
    for (int i = 0; i < 128; ++i) chart.addSeries(xlpp::ChartSeries{});
    test.checkTrue(&chart.series().front() == seriesAddress, "Chart-series handle survives series growth");

    auto& pivot = sheet.addPivotTable(xlpp::PivotTable("RootPivot"));
    auto* pivotAddress = &pivot;
    for (int i = 0; i < 128; ++i) sheet.addPivotTable(xlpp::PivotTable("Pivot" + std::to_string(i)));
    test.checkTrue(&sheet.pivotTables().front() == pivotAddress, "Pivot handle survives worksheet pivot growth");

    auto& image = sheet.addImage(xlpp::Image("A1", {1, 2, 3}, "png"));
    auto* imageAddress = &image;
    for (int i = 0; i < 128; ++i)
        sheet.addImage(xlpp::Image("A1", {1, 2, 3}, "png"));
    test.checkTrue(&sheet.images().front() == imageAddress, "Image handle survives worksheet image growth");

    auto& cfEntry = sheet.conditionalFormatting().add("A1:A2");
    auto* cfEntryAddress = &cfEntry;
    for (int i = 0; i < 128; ++i)
        sheet.conditionalFormatting().add("B" + std::to_string(i + 1));
    test.checkTrue(&sheet.conditionalFormatting().entries().front() == cfEntryAddress,
                   "Conditional-format entry handle survives collection growth");

    auto& cfRule = cfEntry.addRule(xlpp::ConditionalRule::formula("A1>0"));
    auto* cfRuleAddress = &cfRule;
    for (int i = 0; i < 128; ++i)
        cfEntry.addRule(xlpp::ConditionalRule::formula("A1>" + std::to_string(i)));
    test.checkTrue(&cfEntry.rules().front() == cfRuleAddress,
                   "Conditional-format rule handle survives rule growth");

    auto& validation = sheet.dataValidations().add(xlpp::DataValidationType::Whole, "C1");
    auto* validationAddress = &validation;
    for (int i = 0; i < 128; ++i)
        sheet.dataValidations().add(xlpp::DataValidationType::Whole, "C" + std::to_string(i + 2));
    test.checkTrue(&sheet.dataValidations().items().front() == validationAddress,
                   "Data-validation handle survives collection growth");

    xlpp::RichText rich;
    rich.addRun(xlpp::RichTextRun("root"));
    auto* runAddress = &rich.runs().front();
    for (int i = 0; i < 128; ++i) rich.addRun(xlpp::RichTextRun("run"));
    test.checkTrue(&rich.runs().front() == runAddress, "Rich-text run handle survives run growth");

    workbook.customProperties().add(xlpp::CustomProperty("RootProperty", "root"));
    auto* propertyAddress = &workbook.customProperties().items().front();
    for (int i = 0; i < 128; ++i)
        workbook.customProperties().add(xlpp::CustomProperty("P" + std::to_string(i), i));
    test.checkTrue(&workbook.customProperties().items().front() == propertyAddress,
                   "Custom-property handle survives property growth");
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


} // namespace

void registerModelWorkbookTests(std::vector<TestCase>& tests) {
    tests.push_back({"Cell references", testCellReferences});
    tests.push_back({"Range and dimensions", testRangeAndDimensions});
    tests.push_back({"Append and structural edits", testAppendAndStructuralEdits});
    tests.push_back({"Cell convenience methods", testCellConvenience});
    tests.push_back({"Named style association on cells", testNamedStyleAssociation});
    tests.push_back({"Remove worksheet", testRemoveWorksheet});
    tests.push_back({"Worksheet extents", testWorksheetExtents});
    tests.push_back({"DOM shared-string save and load", testDOMSharedStrings});
    tests.push_back({"Rich text shared-string load", testRichTextSharedStrings});
    tests.push_back({"Rich text cell XLSX round-trip", testRichTextCellRoundTrip});
    tests.push_back({"Built-in date format IDs", testBuiltinDateFormatIds});
    tests.push_back({"Row proxy and range helpers", testRowProxyAndRangeHelpers});
    tests.push_back({"Cell style index", testCellStyleIndex});
    tests.push_back({"Stream load and save", testStreamLoadSave});
    tests.push_back({"Built-in date format round-trip", testNumFmtIdDateRoundTrip});
    tests.push_back({"Cell edge cases and cleanup", testEdgeCasesAndCleanup});
    tests.push_back({"Worksheet rows() iteration", testWorksheetRows});
    tests.push_back({"iterRows and iterCols", testIterRowsCols});
    tests.push_back({"Cell::offset()", testCellOffset});
    tests.push_back({"Workbook navigation", testWorkbookNav});
    tests.push_back({"Workbook copyWorksheet", testCopyWorksheet});
    tests.push_back({"Merged cells", testMergedCells});
    tests.push_back({"Worksheet layout", testWorksheetLayout});
    tests.push_back({"XLSX layout round-trip", testRoundTrip});
    tests.push_back({"AutoFilter model", testAutoFilter});
    tests.push_back({"AutoFilter XLSX round-trip", testAutoFilterRoundTrip});
    tests.push_back({"Cell styles", testCellStyles});
    tests.push_back({"Styles XLSX round-trip", testStylesRoundTrip});
    tests.push_back({".h header migration smoke test", testHeaderMigration});
    tests.push_back({"Named styles registry", testNamedStyles});
    tests.push_back({"Named styles XLSX round-trip", testNamedStylesRoundTrip});
    tests.push_back({"Conditional formatting model", testConditionalFormatting});
    tests.push_back({"Conditional formatting XLSX round-trip", testConditionalFormattingRoundTrip});
    tests.push_back({"Data validation model", testDataValidation});
    tests.push_back({"Data validation XLSX round-trip", testDataValidationRoundTrip});
    tests.push_back({"Tables and defined names model", testTablesAndDefinedNames});
    tests.push_back({"Tables and defined names XLSX round-trip", testTablesAndDefinedNamesRoundTrip});
    tests.push_back({"Hyperlinks, comments and properties model", testHyperlinksCommentsAndProperties});
    tests.push_back({"Hyperlinks and properties XLSX round-trip", testHyperlinksAndPropertiesRoundTrip});
    tests.push_back({"Legacy comments XLSX round-trip", testCommentsRoundTrip});
    tests.push_back({"Page setup, protection and images model", testPageSetupProtectionAndImages});
    tests.push_back({"Page setup, protection and images XLSX round-trip", testPageSetupProtectionAndImagesRoundTrip});
    tests.push_back({"Formula metadata and error cells model", testFormulaMetadataAndErrorCells});
    tests.push_back({"Formula metadata and error cells XLSX round-trip", testFormulaMetadataAndErrorCellsRoundTrip});
    tests.push_back({"Stable model child handles", testStableModelHandles});
    tests.push_back({"Direct streaming ZIP reader", testDirectZipReader});
}
