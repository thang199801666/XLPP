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

void testCoreSafetyAndValidation(TestContext& test) {
    // Excel coordinate invariants are enforced at every public entry point.
    const auto max = xlpp::CellReference::parse("XFD1048576");
    test.checkEqual(max.row, xlpp::MaxExcelRows, "Maximum Excel row parses");
    test.checkEqual(max.column, xlpp::MaxExcelColumns, "Maximum Excel column parses");
    for (const auto& bad : {std::string("XFE1"), std::string("A1048577")}) {
        bool threw = false;
        try { (void)xlpp::CellReference::parse(bad); } catch (const std::out_of_range&) { threw = true; }
        test.checkTrue(threw, "Out-of-bounds A1 reference rejected: " + bad);
    }
    for (const auto& bad : {std::string("A4$2"), std::string("A$4$2"), std::string("A$$1"),
                             std::string("$$A1"), std::string("$A1$"), std::string("A$")}) {
        bool malformed = false;
        try { (void)xlpp::CellReference::parse(bad); } catch (const std::invalid_argument&) { malformed = true; }
        test.checkTrue(malformed, "Malformed absolute A1 reference rejected: " + bad);
    }
    const auto absolute = xlpp::CellReference::parse("$XFD$1048576");
    test.checkEqual(absolute.row, xlpp::MaxExcelRows, "Absolute maximum Excel row parses");
    test.checkEqual(absolute.column, xlpp::MaxExcelColumns, "Absolute maximum Excel column parses");
    bool threw = false;
    try { (void)xlpp::CellReference::columnName(xlpp::MaxExcelColumns + 1); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "columnName rejects XFE and beyond");
    threw = false;
    try { (void)xlpp::CellReference::columnIndex("XFE"); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "columnIndex rejects XFE and beyond");
    xlpp::Worksheet bounded("Bounds");
    threw = false;
    try { bounded.cell(xlpp::MaxExcelRows + 1, 1); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "Worksheet::cell rejects row overflow");
    threw = false;
    try { bounded.cell(1, xlpp::MaxExcelColumns + 1); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "Worksheet::cell rejects column overflow");
    threw = false;
    try { bounded.rowDimension(xlpp::MaxExcelRows + 1); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "Row dimensions reject row overflow");
    bounded.cell(xlpp::MaxExcelRows, 1).setValue(1.0);
    threw = false;
    try { bounded.insertRows(xlpp::MaxExcelRows, 1); } catch (const std::out_of_range&) { threw = true; }
    test.checkTrue(threw, "Worksheet structural insert rejects coordinate overflow");

    // Worksheet names use Excel's 31-character/forbidden-character contract.
    xlpp::Workbook names;
    names.addWorksheet(std::string(31, 'A'));
    threw = false;
    try { names.addWorksheet(std::string(32, 'B')); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Worksheet names longer than 31 characters are rejected");
    threw = false;
    try { names.addWorksheet("Bad/Name"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Worksheet forbidden characters are rejected");
    xlpp::Workbook caseNames;
    caseNames.addWorksheet("Data");
    test.checkTrue(caseNames.worksheet("data") != nullptr, "Worksheet lookup follows Excel case-insensitive semantics");
    threw = false;
    try { caseNames.addWorksheet("DATA"); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Case-insensitive duplicate worksheet names are rejected");

    xlpp::Workbook scopedNames;
    scopedNames.addWorksheet("First");
    scopedNames.addWorksheet("Second");
    xlpp::DefinedName local0("ScopedRate", "First!$A$1");
    local0.setLocalSheetId(0);
    scopedNames.addDefinedName(local0);
    xlpp::DefinedName local1("scopedrate", "Second!$A$1");
    local1.setLocalSheetId(1);
    scopedNames.addDefinedName(local1);
    test.checkTrue(scopedNames.definedName("SCOPEDRATE", std::size_t{0}) != nullptr,
                   "Defined names allow the same identifier in different local scopes");
    test.checkTrue(scopedNames.definedName("SCOPEDRATE", std::size_t{1}) != nullptr,
                   "Scoped defined-name lookup is case-insensitive");
    xlpp::DefinedName duplicateLocal("SCOPEDRATE", "Second!$B$1");
    duplicateLocal.setLocalSheetId(1);
    threw = false;
    try { scopedNames.addDefinedName(duplicateLocal); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Defined-name duplicate in the same scope is rejected case-insensitively");
    scopedNames.removeWorksheet("First");
    test.checkTrue(scopedNames.definedNames().size() == 1 &&
                   scopedNames.definedName("ScopedRate", std::size_t{0}) != nullptr,
                   "Removing a sheet removes owned local names and compacts later localSheetId values");

    // Explicit validation catches model corruption reachable through mutable APIs.
    xlpp::Workbook invalid;
    invalid.addWorksheet("First");
    invalid.addWorksheet("Second");
    invalid[1].rename("FIRST"); // syntactically valid, but workbook-wide duplicate
    auto validation = invalid.validate();
    test.checkTrue(!validation.ok() && validation.errorCount >= 1, "Workbook validation catches duplicate worksheet names after mutation");
    test.checkTrue(std::any_of(validation.issues.begin(), validation.issues.end(), [](const auto& issue) { return issue.code == "WS002"; }),
                   "Workbook validation emits stable WS002 diagnostic code");

    xlpp::Workbook duplicateTables;
    auto& ta = duplicateTables.addWorksheet("A");
    auto& tb = duplicateTables.addWorksheet("B");
    ta.addTable("Sales", "A1:B2");
    tb.addTable("sales", "A1:B2");
    validation = duplicateTables.validate();
    test.checkTrue(std::any_of(validation.issues.begin(), validation.issues.end(), [](const auto& issue) { return issue.code == "TB002"; }),
                   "Workbook validation catches workbook-wide duplicate table names");

    xlpp::Workbook badScope;
    badScope.addWorksheet("Only");
    auto& scoped = badScope.addDefinedName(xlpp::DefinedName("Scoped", "Only!$A$1"));
    scoped.setLocalSheetId(99);
    validation = badScope.validate();
    test.checkTrue(std::any_of(validation.issues.begin(), validation.issues.end(), [](const auto& issue) { return issue.code == "DN002"; }),
                   "Workbook validation catches missing defined-name scope");

    xlpp::Workbook invalidPivot;
    auto& pivotSheet = invalidPivot.addWorksheet("Pivot");
    pivotSheet.append({std::string("Key"), std::string("Value")});
    pivotSheet.append({std::string("A"), 1.0});
    xlpp::PivotTable badPivot("BadPivot");
    badPivot.cache().setSourceData("Missing!A1:B2");
    badPivot.addDataField("NotAField");
    pivotSheet.addPivotTable(std::move(badPivot));
    validation = invalidPivot.validate();
    test.checkTrue(std::any_of(validation.issues.begin(), validation.issues.end(), [](const auto& issue) { return issue.code == "PV007"; }),
                   "Workbook validation catches Pivot cache source targeting a missing worksheet");
    test.checkTrue(std::any_of(validation.issues.begin(), validation.issues.end(), [](const auto& issue) { return issue.code == "PV020"; }),
                   "Workbook validation catches unresolved Pivot data fields");

    xlpp::Workbook invalidChart;
    auto& chartSheet = invalidChart.addWorksheet("Chart");
    xlpp::Chart badChart(xlpp::Chart::Type::Bar);
    badChart.setWidth(0);
    auto& badSeries = badChart.addSeries(xlpp::ChartSeries("Missing"));
    badSeries.reference("Missing", "$B$2:$B$4");
    chartSheet.addChart(std::move(badChart));
    validation = invalidChart.validate();
    test.checkTrue(std::any_of(validation.issues.begin(), validation.issues.end(), [](const auto& issue) { return issue.code == "CH001"; }),
                   "Workbook validation catches non-positive chart dimensions");
    test.checkTrue(std::any_of(validation.issues.begin(), validation.issues.end(), [](const auto& issue) { return issue.code == "CH011"; }),
                   "Workbook validation catches chart references targeting a missing worksheet");

    test.checkTrue(xlpp::SaveOptions{}.durableWrite,
                   "Path saves enable stable-storage durability by default");

    const auto dir = std::filesystem::temp_directory_path();
    const auto stablePath = dir / "xlpp_core_atomic_stable.xlsx";
    const auto corruptPath = dir / "xlpp_core_load_corrupt.xlsx";
    const auto zipPath = dir / "xlpp_core_stream_move.zip";
    const auto crcPath = dir / "xlpp_core_crc.zip";
    const auto duplicatePath = dir / "xlpp_core_duplicate.zip";
    const auto readBytes = [](const std::filesystem::path& p) {
        std::ifstream f(p, std::ios::binary);
        return std::string((std::istreambuf_iterator<char>(f)), {});
    };
    const auto writeBytes = [](const std::filesystem::path& p, const std::string& bytes) {
        std::ofstream f(p, std::ios::binary | std::ios::trunc);
        f.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    };

    // load() has a strong exception guarantee: a failed parse does not clear the live workbook.
    xlpp::Workbook live;
    live.addWorksheet("Keep").cell("A1").setValue("original");
    writeBytes(corruptPath, "not a zip");
    threw = false;
    try { live.load(corruptPath); } catch (const std::exception&) { threw = true; }
    test.checkTrue(threw, "Corrupt load fails cleanly");
    test.checkEqual(live.sheetCount(), std::size_t{1}, "Failed load preserves existing sheet count");
    test.checkEqual(std::get<std::string>(live.worksheet("Keep")->cell("A1").value()), std::string("original"),
                    "Failed load preserves existing cell state");

    // Path saves stage in the destination directory and atomically replace only on success.
    xlpp::Workbook first;
    first.addWorksheet("Data").cell("A1").setValue("old");
    first.save(stablePath);
    xlpp::Workbook second;
    second.addWorksheet("Data").cell("A1").setValue("new");
    second.save(stablePath);
    xlpp::Workbook reloaded;
    reloaded.load(stablePath);
    test.checkEqual(std::get<std::string>(reloaded.worksheet("Data")->cell("A1").value()), std::string("new"),
                    "Atomic save replaces an existing workbook with complete new content");
    invalid = xlpp::Workbook{};
    invalid.addWorksheet("One");
    invalid.addWorksheet("Two");
    invalid[1].rename("ONE");
    threw = false;
    try { invalid.save(stablePath); } catch (const std::invalid_argument&) { threw = true; }
    test.checkTrue(threw, "Validation failure aborts staged atomic save");
    reloaded.load(stablePath);
    test.checkEqual(std::get<std::string>(reloaded.worksheet("Data")->cell("A1").value()), std::string("new"),
                    "Failed atomic save leaves previous destination byte-valid and unchanged semantically");
    std::size_t leakedStagingFiles = 0;
    const auto prefix = "." + stablePath.filename().string() + ".xlpp_";
    for (const auto& entry : std::filesystem::directory_iterator(dir))
        if (entry.path().filename().string().rfind(prefix, 0) == 0) ++leakedStagingFiles;
    test.checkEqual(leakedStagingFiles, std::size_t{0}, "Atomic save cleans all staging files on success/failure");

    const auto lowLatencyPath = dir / "xlpp_core_nondurable_save.xlsx";
    xlpp::SaveOptions lowLatency;
    lowLatency.durableWrite = false;
    second.save(lowLatencyPath, lowLatency);
    xlpp::Workbook lowLatencyReload;
    lowLatencyReload.load(lowLatencyPath);
    test.checkEqual(lowLatencyReload.worksheet("Data")->cell("A1").stringValueOr(""), std::string("new"),
                    "Opting out of fsync durability preserves normal atomic save semantics");
    std::filesystem::remove(lowLatencyPath);

    // Moving an already-open deflate source transfers zlib ownership and rebases next_in.
    const std::string payload(200000, 'x');
    {
        xlpp::internal::ZipArchive zip;
        zip.add("payload.txt", payload, true);
        zip.save(zipPath);
    }
    xlpp::internal::ZipArchiveReader reader(zipPath);
    auto source = reader.openEntry("payload.txt");
    std::array<unsigned char, 37> firstChunk{};
    const auto firstCount = source.read(firstChunk.data(), firstChunk.size());
    xlpp::internal::ZipEntrySource moved(std::move(source));
    std::string recovered(reinterpret_cast<const char*>(firstChunk.data()), firstCount);
    std::array<unsigned char, 4096> buffer{};
    while (const auto count = moved.read(buffer.data(), buffer.size()))
        recovered.append(reinterpret_cast<const char*>(buffer.data()), count);
    test.checkEqual(recovered.size(), payload.size(), "Moved active ZIP source produces complete payload");
    test.checkEqual(recovered, payload, "Moved active ZIP source preserves exact decompressed bytes");
    test.checkTrue(moved.complete(), "Moved active ZIP source reaches verified CRC-complete state");

    {
        xlpp::internal::ZipArchiveReaderLimits limits;
        limits.maxEntryBytes = 1024;
        threw = false;
        try { xlpp::internal::ZipArchiveReader limited(zipPath, limits); } catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Streaming ZIP metadata limits reject oversized declared entries before inflate");

        auto bytes = readBytes(zipPath);
        const std::string centralSig("PK\x01\x02", 4);
        const auto central = bytes.find(centralSig);
        test.checkTrue(central != std::string::npos, "Declared-size fixture central record located");
        if (central != std::string::npos && central + 28 <= bytes.size()) {
            // Lie about the uncompressed size while leaving the valid deflate
            // stream untouched. The pull reader must stop at the declared
            // output budget rather than expanding the whole payload.
            bytes[central + 24] = 1;
            bytes[central + 25] = bytes[central + 26] = bytes[central + 27] = 0;
            const auto budgetPath = dir / "xlpp_core_output_budget.zip";
            writeBytes(budgetPath, bytes);
            threw = false;
            try {
                xlpp::internal::ZipArchiveReader budgeted(budgetPath);
                (void)budgeted.readEntry("payload.txt");
            } catch (const std::exception&) { threw = true; }
            test.checkTrue(threw, "Streaming inflater enforces declared uncompressed output budget");
            std::filesystem::remove(budgetPath);
        }
    }

    // CRC corruption is rejected by both materializing and streaming ZIP paths.
    {
        xlpp::internal::ZipArchive zip;
        zip.add("crc.txt", "abcdef", false);
        zip.save(crcPath);
        auto bytes = readBytes(crcPath);
        const auto namePos = bytes.find("crc.txt");
        test.checkTrue(namePos != std::string::npos, "CRC fixture local name located");
        if (namePos != std::string::npos && namePos + 7 < bytes.size()) bytes[namePos + 7] ^= 0x01;
        writeBytes(crcPath, bytes);
        threw = false;
        try { (void)xlpp::internal::ZipArchive::open(crcPath); } catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Materializing ZIP reader rejects CRC corruption");
        threw = false;
        try {
            xlpp::internal::ZipArchiveReader direct(crcPath);
            (void)direct.readEntry("crc.txt");
        } catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Direct streaming ZIP reader rejects CRC corruption");
    }

    // Duplicate central-directory names are rejected rather than silently using last-wins semantics.
    {
        xlpp::internal::ZipArchive zip;
        zip.add("a.txt", "a", false);
        zip.add("b.txt", "b", false);
        zip.save(duplicatePath);
        auto bytes = readBytes(duplicatePath);
        const std::string centralSig("PK\x01\x02", 4);
        const auto firstCentral = bytes.find(centralSig);
        const auto secondCentral = firstCentral == std::string::npos ? std::string::npos : bytes.find(centralSig, firstCentral + 4);
        test.checkTrue(secondCentral != std::string::npos, "Duplicate-entry fixture has second central record");
        if (secondCentral != std::string::npos) {
            const auto nameOffset = secondCentral + 46;
            if (nameOffset < bytes.size()) bytes[nameOffset] = 'a'; // b.txt -> a.txt in central directory only
        }
        writeBytes(duplicatePath, bytes);
        threw = false;
        try { xlpp::internal::ZipArchiveReader direct(duplicatePath); } catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Direct ZIP reader rejects duplicate central-directory entry names");
        threw = false;
        try { (void)xlpp::internal::ZipArchive::open(duplicatePath); } catch (const std::exception&) { threw = true; }
        test.checkTrue(threw, "Materializing ZIP reader rejects ambiguous duplicate/mismatched entries");
    }

    for (const auto& path : {stablePath, corruptPath, zipPath, crcPath, duplicatePath}) {
        std::error_code ignored;
        std::filesystem::remove(path, ignored);
    }
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
    // internal worksheet deque must not invalidate the source mid-copy, and a
    // sheet clone must carry its owned dependency topology instead of retaining
    // accidental references/global identifiers from the source sheet.
    xlpp::Workbook wb;
    auto& src = wb.addWorksheet("Source");
    auto& other = wb.addWorksheet("Other");
    src.append({std::string("Category"), std::string("Value")});
    src.append({std::string("A"), 3.25});
    src.append({std::string("B"), 4.5});
    src.cell("C2").setFormula("'Source'!$B$2+'Other'!$A$1");
    src.mergeCells("A4:B4");
    src.rowDimension(1).height = 22.0;
    other.cell("A1").setValue(10.0);

    auto& table = src.addTable("SalesTable", "A1:B3");
    table.addColumn("Category");
    table.addColumn("Value");

    xlpp::Chart chart(xlpp::Chart::Type::Bar);
    auto& series = chart.addSeries(xlpp::ChartSeries("Value"));
    series.categories("Source", "$A$2:$A$3");
    series.reference("Source", "$B$2:$B$3");
    src.addChart(std::move(chart));

    xlpp::PivotTable pivot("SalesPivot");
    pivot.setLocation("E2");
    pivot.cache().setSourceData("'Source'!$A$1:$B$3");
    pivot.cache().setFields({"Category", "Value"});
    pivot.addRowField("Category");
    pivot.addDataField("Value");
    src.addPivotTable(std::move(pivot));

    xlpp::DefinedName localName("LocalMetric", "'Source'!$B$2");
    localName.setLocalSheetId(0);
    wb.addDefinedName(std::move(localName));

    auto& copy = wb.copyWorksheet(src, "Copy");
    test.checkEqual(copy.name(), std::string("Copy"), "Copy named correctly");
    test.checkNear(copy.cell("B2").numericValueOr(0.0), 3.25, 1e-12, "Copy keeps number");
    test.checkEqual(copy.mergedRanges().size(), std::size_t{1}, "Copy keeps merges");
    test.checkNear(copy.tryRowDimension(1)->height.value_or(0.0), 22.0, 1e-12, "Copy keeps dimensions");
    test.checkEqual(copy.cell("C2").formula(), std::string("'Copy'!$B$2+'Other'!$A$1"),
                    "Copied cell self-reference follows clone while cross-sheet reference stays intact");
    test.checkEqual(copy.tables().front().name(), std::string("SalesTable_2"),
                    "Copied table receives workbook-unique identifier");
    test.checkEqual(copy.tables().front().displayName(), std::string("SalesTable_2"),
                    "Copied table display name follows unique identifier");
    test.checkEqual(copy.pivotTables().front().name(), std::string("SalesPivot_2"),
                    "Copied pivot receives workbook-unique identifier");
    test.checkEqual(copy.pivotTables().front().cache().sourceData(), std::string("'Copy'!$A$1:$B$3"),
                    "Copied pivot cache source follows clone");
    test.checkEqual(copy.charts().front().series().front().valuesReference(), std::string("='Copy'!$B$2:$B$3"),
                    "Copied chart values reference follows clone");
    test.checkEqual(copy.charts().front().series().front().categoriesReference(), std::string("='Copy'!$A$2:$A$3"),
                    "Copied chart category reference follows clone");

    const auto* copiedLocal = wb.definedName("LocalMetric", std::size_t{2});
    test.checkTrue(copiedLocal != nullptr, "Sheet-local defined name is cloned into copied sheet scope");
    if (copiedLocal)
        test.checkEqual(copiedLocal->value(), std::string("'Copy'!$B$2"),
                        "Copied sheet-local defined name follows clone");

    const auto validation = wb.validate();
    test.checkTrue(validation.ok(), "Copied sheet topology remains workbook-valid");

    const auto path = std::filesystem::temp_directory_path() / "xlpp_copy_sheet_topology.xlsx";
    wb.save(path);
    xlpp::Workbook loaded;
    loaded.load(path);
    auto* loadedCopy = loaded.worksheet("Copy");
    test.checkTrue(loadedCopy != nullptr, "Copied sheet round-trips through package");
    if (loadedCopy) {
        test.checkEqual(loadedCopy->cell("C2").formula(), std::string("'Copy'!$B$2+'Other'!$A$1"),
                        "Copied formula topology survives reload");
        test.checkTrue(loadedCopy->table("SalesTable_2") != nullptr, "Copied table unique name survives reload");
        test.checkTrue(!loadedCopy->pivotTables().empty() && loadedCopy->pivotTables().front().name() == "SalesPivot_2",
                       "Copied pivot unique name survives reload");
        test.checkTrue(!loadedCopy->pivotTables().empty() &&
                           loadedCopy->pivotTables().front().cache().sourceData().find("Copy") != std::string::npos &&
                           loadedCopy->pivotTables().front().cache().sourceData().find("Source") == std::string::npos,
                       "Copied pivot source survives reload and remains attached to clone");
        test.checkTrue(!loadedCopy->charts().empty() &&
                           loadedCopy->charts().front().series().front().valuesReference().find("Copy") != std::string::npos &&
                           loadedCopy->charts().front().series().front().valuesReference().find("Source") == std::string::npos,
                       "Copied chart source survives reload and remains attached to clone");
    }
    std::filesystem::remove(path);

    for (std::size_t i = 0; i < 40; ++i) wb.addWorksheet("Fill" + std::to_string(i));
    auto& late = wb.copyWorksheet(wb[0], "LateClone");
    test.checkNear(late.cell("B2").numericValueOr(0.0), 3.25, 1e-12,
                   "Copy after many inserts keeps source data");
    test.checkEqual(wb.sheetCount(), std::size_t{44}, "Workbook grew as expected");
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
