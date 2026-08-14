#include "Packaging/ZipArchive.h"
#include "XML/XmlUtilities.h"
#include <XLPP/XLPP.h>

#include <chrono>
#include <cstdlib>
#include <filesystem>
#include <iostream>
#include <string>

namespace {
int failures = 0;
void check(bool ok, const char* message) {
    if (!ok) { ++failures; std::cerr << "FAIL: " << message << '\n'; }
}
template <class T, class U>
void checkEqual(const T& actual, const U& expected, const char* message) { check(actual == expected, message); }

std::filesystem::path tempPath(const char* stem, const char* ext) {
    const auto tick = std::chrono::high_resolution_clock::now().time_since_epoch().count();
    return std::filesystem::temp_directory_path() / (std::string(stem) + "_" + std::to_string(tick) + ext);
}
void cleanup(const std::filesystem::path& p) {
    if (std::getenv("XLPP_KEEP_P1V_ARTIFACTS")) std::cout << "P1V_ARTIFACT=" << p.string() << '\n';
    else std::filesystem::remove(p);
}

xlpp::Chart makeChart() {
    xlpp::Chart chart(xlpp::Chart::Type::Line);
    chart.setTitle("P1V Chartsheet");
    xlpp::ChartSeries series("Values");
    series.setCategoriesReference("'Data'!$A$2:$A$4");
    series.setValuesReference("'Data'!$B$2:$B$4");
    chart.addSeries(std::move(series));
    return chart;
}

void populate(xlpp::Workbook& wb) {
    auto& ws = wb.addWorksheet("Data");
    ws.append({std::string("Category"), std::string("Value")});
    ws.append({std::string("A"), 10.0});
    ws.append({std::string("B"), 20.0});
    ws.append({std::string("C"), 30.0});
}

void testAdvancedPageSetupProtectionAndCustomViews() {
    const auto path = tempPath("xlpp_p1v_advanced", ".xltx");
    xlpp::Workbook wb;
    wb.setTemplate(true);
    populate(wb);
    auto& cs = wb.addChartsheet("Dashboard", makeChart());
    wb.setActiveWorkbookSheet("Dashboard");

    auto& setup = cs.pageSetup();
    setup.setOrientation(xlpp::PageOrientation::Landscape);
    setup.setPaperHeight("210mm");
    setup.setPaperWidth("297mm");
    setup.setPageOrder(xlpp::PageOrder::OverThenDown);
    setup.setUsePrinterDefaults(false);
    setup.setCellComments(xlpp::PageCellComments::AtEnd);
    setup.setErrors(xlpp::PageErrorDisplay::Dash);
    setup.setHorizontalDpi(600);
    setup.setVerticalDpi(600);
    setup.setCopies(3);

    auto& protection = cs.protection();
    protection.setContent(true);
    protection.setObjects(true);
    protection.setAlgorithmName("SHA-512");
    protection.setHashValue("AQIDBA==");
    protection.setSaltValue("BQYHCA==");
    protection.setSpinCount(100000);

    auto& custom = cs.customViews().emplace_back();
    custom.setGuid("{11111111-2222-3333-4444-555555555555}");
    custom.setScale(90);
    custom.setState(xlpp::CustomChartsheetViewState::Hidden);
    custom.setZoomToFit(false);
    custom.pageMargins().setLeft(0.25);
    custom.pageMargins().setRight(0.25);
    custom.pageSetup().setHorizontalDpi(300);
    custom.pageSetup().setVerticalDpi(300);
    custom.pageSetup().setCopies(2);
    custom.headerFooter().setOddHeader("&CCustom View");

    wb.save(path);
    const auto zip = xlpp::internal::ZipArchive::open(path);
    const auto xml = zip.get("xl/chartsheets/sheet1.xml");
    check(xml.find("paperHeight=\"210mm\"") != std::string::npos, "Chartsheet paperHeight serializes");
    check(xml.find("horizontalDpi=\"600\"") != std::string::npos && xml.find("copies=\"3\"") != std::string::npos,
          "Chartsheet DPI/copies serialize");
    check(xml.find("algorithmName=\"SHA-512\"") != std::string::npos && xml.find("spinCount=\"100000\"") != std::string::npos,
          "Modern Chartsheet protection serializes");
    check(xml.find("<customSheetViews>") != std::string::npos && xml.find("state=\"hidden\"") != std::string::npos,
          "Custom Chartsheet view serializes");

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedCs = static_cast<const xlpp::Workbook&>(loaded).chartsheet("Dashboard");
    check(loadedCs != nullptr, "Advanced Chartsheet reloads");
    if (loadedCs) {
        const auto& ps = loadedCs->pageSetup();
        check(ps.paperHeight() && *ps.paperHeight() == "210mm" && ps.paperWidth() && *ps.paperWidth() == "297mm",
              "Chartsheet paper dimensions round-trip");
        check(ps.pageOrder() == xlpp::PageOrder::OverThenDown && ps.usePrinterDefaults() && !*ps.usePrinterDefaults(),
              "Chartsheet printer/page-order state round-trips");
        check(ps.cellComments() == xlpp::PageCellComments::AtEnd && ps.errors() == xlpp::PageErrorDisplay::Dash,
              "Chartsheet comment/error print behavior round-trips");
        check(ps.horizontalDpi() && *ps.horizontalDpi() == 600 && ps.verticalDpi() && *ps.verticalDpi() == 600
              && ps.copies() && *ps.copies() == 3, "Chartsheet DPI/copies round-trip");
        check(loadedCs->protection().algorithmName() && *loadedCs->protection().algorithmName() == "SHA-512"
              && loadedCs->protection().spinCount() && *loadedCs->protection().spinCount() == 100000,
              "Modern Chartsheet protection metadata round-trips");
        checkEqual(loadedCs->customViews().size(), std::size_t{1}, "Custom Chartsheet view reloads");
        if (!loadedCs->customViews().empty()) {
            const auto& cv = loadedCs->customViews().front();
            check(cv.guid() == "{11111111-2222-3333-4444-555555555555}" && cv.scale() && *cv.scale() == 90,
                  "Custom Chartsheet view identity/scale round-trips");
            check(cv.state() == xlpp::CustomChartsheetViewState::Hidden && cv.zoomToFit() && !*cv.zoomToFit(),
                  "Custom Chartsheet view state round-trips");
            check(cv.pageSetupValue() && cv.pageSetupValue()->horizontalDpi() && *cv.pageSetupValue()->horizontalDpi() == 300,
                  "Custom Chartsheet nested page setup round-trips");
            check(cv.headerFooterValue() && cv.headerFooterValue()->oddHeader() == "&CCustom View",
                  "Custom Chartsheet nested header/footer round-trips");
        }
    }
    cleanup(path);
}

void testCustomViewPreservationThenSelectiveMutation() {
    const auto source = tempPath("xlpp_p1v_custom_source", ".xltx");
    const auto metadataSave = tempPath("xlpp_p1v_custom_metadata", ".xltx");
    const auto customSave = tempPath("xlpp_p1v_custom_mutated", ".xltx");

    xlpp::Workbook seed;
    seed.setTemplate(true);
    populate(seed);
    auto& cs = seed.addChartsheet("Dashboard", makeChart());
    auto& custom = cs.customViews().emplace_back();
    custom.setGuid("{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}");
    custom.setScale(80);
    custom.headerFooter().setOddFooter("&RRaw-Preserve");
    seed.save(source);

    const auto sourceZip = xlpp::internal::ZipArchive::open(source);
    const auto sourceXml = sourceZip.get("xl/chartsheets/sheet1.xml");
    const auto sourceBlocks = xlpp::internal::tags(sourceXml, "customSheetViews");
    check(sourceBlocks.size() == 1, "Seed customSheetViews block exists");

    xlpp::Workbook wb;
    wb.load(source);
    auto* imported = wb.chartsheet("Dashboard");
    check(imported && imported->imported(), "Seed Chartsheet reloads preservation-backed");
    if (!imported) return;

    imported->view().setZoomScale(135); // unrelated metadata edit
    wb.save(metadataSave);
    const auto metadataXml = xlpp::internal::ZipArchive::open(metadataSave).get("xl/chartsheets/sheet1.xml");
    const auto metadataBlocks = xlpp::internal::tags(metadataXml, "customSheetViews");
    check(sourceBlocks.size() == 1 && metadataBlocks.size() == 1 && sourceBlocks.front() == metadataBlocks.front(),
          "Untouched customSheetViews remains byte-preserved during unrelated Chartsheet patch");

    imported->customViews().front().setScale(95);
    wb.save(customSave);
    const auto customXml = xlpp::internal::ZipArchive::open(customSave).get("xl/chartsheets/sheet1.xml");
    check(customXml.find("scale=\"95\"") != std::string::npos,
          "Explicit custom view mutation opts into semantic regeneration");

    xlpp::Workbook reloaded;
    reloaded.load(customSave);
    const auto* finalCs = static_cast<const xlpp::Workbook&>(reloaded).chartsheet("Dashboard");
    check(finalCs && !finalCs->customViews().empty() && finalCs->customViews().front().scale()
          && *finalCs->customViews().front().scale() == 95,
          "Regenerated custom view survives reload");

    cleanup(source); cleanup(metadataSave); cleanup(customSave);
}


void testCustomViewValidation() {
    xlpp::Workbook wb;
    populate(wb);
    auto& cs = wb.addChartsheet("Dashboard", makeChart());
    auto& a = cs.customViews().emplace_back();
    a.setGuid("{ABCDEFAB-1111-2222-3333-444444444444}");
    auto& b = cs.customViews().emplace_back();
    b.setGuid("{abcdefab-1111-2222-3333-444444444444}");
    const auto report = wb.validateModelIntegrity();
    bool duplicateIssue = false;
    for (const auto& issue : report.issues)
        if (issue.code == "chartsheet.custom_view_duplicate_guid") duplicateIssue = true;
    check(duplicateIssue, "Duplicate custom Chartsheet view GUID is diagnosed case-insensitively");

    const auto path = tempPath("xlpp_p1v_duplicate_custom_view", ".xlsx");
    bool rejected = false;
    try { wb.save(path); } catch (const std::invalid_argument&) { rejected = true; }
    check(rejected, "Duplicate custom Chartsheet view GUID is rejected by serializer");
    cleanup(path);
}

void testWorksheetExtendedPageSetupRoundTrip() {
    const auto path = tempPath("xlpp_p1v_worksheet_pagesetup", ".xlsx");
    xlpp::Workbook wb;
    auto& ws = wb.addWorksheet("Sheet1");
    ws.cell("A1").setValue("page setup");
    auto& setup = ws.pageSetup();
    setup.setPaperHeight("11in");
    setup.setPaperWidth("8.5in");
    setup.setPageOrder(xlpp::PageOrder::DownThenOver);
    setup.setUsePrinterDefaults(true);
    setup.setCellComments(xlpp::PageCellComments::AsDisplayed);
    setup.setErrors(xlpp::PageErrorDisplay::Blank);
    setup.setHorizontalDpi(1200);
    setup.setVerticalDpi(1200);
    setup.setCopies(4);
    wb.save(path);

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto& ps = loaded[0].pageSetup();
    check(ps.paperHeight() && *ps.paperHeight() == "11in" && ps.paperWidth() && *ps.paperWidth() == "8.5in",
          "Worksheet extended paper dimensions round-trip");
    check(ps.horizontalDpi() && *ps.horizontalDpi() == 1200 && ps.copies() && *ps.copies() == 4,
          "Worksheet extended DPI/copies round-trip");
    cleanup(path);
}
} // namespace

int main() {
    testAdvancedPageSetupProtectionAndCustomViews();
    testCustomViewPreservationThenSelectiveMutation();
    testWorksheetExtendedPageSetupRoundTrip();
    testCustomViewValidation();
    if (failures == 0) { std::cout << "P1V Chartsheet depth regression: PASS\n"; return 0; }
    std::cerr << failures << " P1V Chartsheet-depth check(s) failed\n";
    return 1;
}
