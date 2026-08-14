#include "Packaging/ZipArchive.h"
#include <XLPP/XLPP.h>

#include <chrono>
#include <cstdlib>
#include <filesystem>
#include <iostream>
#include <string>
#include <vector>

#ifndef XLPP_TEST_SOURCE_DIR
#define XLPP_TEST_SOURCE_DIR "."
#endif

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
    if (std::getenv("XLPP_KEEP_P1U_ARTIFACTS")) std::cout << "P1U_ARTIFACT=" << p.string() << '\n';
    else std::filesystem::remove(p);
}

xlpp::Chart makeChart(std::string title = "Dashboard") {
    xlpp::Chart chart(xlpp::Chart::Type::Bar);
    chart.setTitle(std::move(title));
    xlpp::ChartSeries series("Values");
    series.setCategoriesReference("'Data'!$A$2:$A$4");
    series.setValuesReference("'Data'!$B$2:$B$4");
    chart.addSeries(std::move(series));
    return chart;
}

void populateData(xlpp::Worksheet& ws) {
    ws.append({std::string("Category"), std::string("Value")});
    ws.append({std::string("A"), 10.0});
    ws.append({std::string("B"), 20.0});
    ws.append({std::string("C"), 30.0});
}

void testTemplateMixedVisibilityAndActiveChartsheet() {
    const auto path = tempPath("xlpp_p1u_mixed_template", ".xltx");
    xlpp::Workbook wb;
    wb.setTemplate(true);
    populateData(wb.addWorksheet("Data"));
    auto& dashboard = wb.addChartsheet("Dashboard", makeChart());
    wb.addWorksheet("Hidden Data").cell("A1").setValue("hidden");
    wb.addChartsheet("Secret Chart", makeChart("Secret"));

    dashboard.properties().setCodeName("ChartDashboard");
    dashboard.properties().setPublished(true);
    dashboard.properties().setTabColor("FF336699");
    dashboard.view().setZoomScale(125);
    dashboard.view().setZoomToFit(false);
    dashboard.protection().setContent(true);
    dashboard.protection().setObjects(true);
    dashboard.protection().setPassword("pw");
    auto& margins = dashboard.pageMargins();
    margins.setLeft(0.4); margins.setRight(0.4); margins.setTop(0.5); margins.setBottom(0.5); margins.setHeader(0.2); margins.setFooter(0.2);
    auto& setup = dashboard.pageSetup();
    setup.setOrientation(xlpp::PageOrientation::Landscape);
    setup.setPaperSize(xlpp::PaperSize::A4);
    setup.setScale(85);
    setup.setFirstPageNumber(2);
    setup.setUseFirstPageNumber(true);
    setup.setBlackAndWhite(true);
    dashboard.headerFooter().setOddHeader("&CAdvanced Header");
    dashboard.headerFooter().setOddFooter("&RPage &P");

    wb.setWorkbookSheetVisibility(2, xlpp::WorkbookSheetVisibility::Hidden);
    wb.setWorkbookSheetVisibility(3, xlpp::WorkbookSheetVisibility::VeryHidden);
    wb.setActiveWorkbookSheetIndex(1);
    wb.save(path);

    const auto zip = xlpp::internal::ZipArchive::open(path);
    const auto workbookXml = zip.get("xl/workbook.xml");
    check(workbookXml.find("activeTab=\"1\"") != std::string::npos, "Chartsheet can be the active workbook tab");
    check(workbookXml.find("name=\"Hidden Data\"") != std::string::npos && workbookXml.find("state=\"hidden\"") != std::string::npos,
          "Mixed workbook serializes hidden state");
    check(workbookXml.find("name=\"Secret Chart\"") != std::string::npos && workbookXml.find("state=\"veryHidden\"") != std::string::npos,
          "Mixed workbook serializes veryHidden state");
    check(zip.get("[Content_Types].xml").find("spreadsheetml.template.main+xml") != std::string::npos,
          "Mixed Chartsheet workbook retains XLTX template content type");

    const auto csXml = zip.get("xl/chartsheets/sheet1.xml");
    check(csXml.find("codeName=\"ChartDashboard\"") != std::string::npos, "Chartsheet codeName serializes");
    check(csXml.find("rgb=\"FF336699\"") != std::string::npos, "Chartsheet tab color serializes");
    check(csXml.find("zoomScale=\"125\"") != std::string::npos && csXml.find("zoomToFit=\"0\"") != std::string::npos,
          "Chartsheet view zoom serializes");
    check(csXml.find("<sheetProtection") != std::string::npos && csXml.find("password=") != std::string::npos,
          "Chartsheet protection serializes");
    check(csXml.find("orientation=\"landscape\"") != std::string::npos && csXml.find("paperSize=\"9\"") != std::string::npos,
          "Chartsheet page setup serializes");
    check(csXml.find("Advanced Header") != std::string::npos, "Chartsheet header/footer serializes");

    xlpp::Workbook loaded;
    loaded.load(path);
    check(loaded.isTemplate(), "XLTX template identity reloads");
    checkEqual(loaded.activeWorkbookSheetIndex(), std::size_t{1}, "Active Chartsheet index reloads");
    checkEqual(loaded.workbookSheetVisibility(2), xlpp::WorkbookSheetVisibility::Hidden, "Hidden worksheet state reloads");
    checkEqual(loaded.workbookSheetVisibility(3), xlpp::WorkbookSheetVisibility::VeryHidden, "VeryHidden chartsheet state reloads");
    const auto descriptors = loaded.workbookSheets();
    check(descriptors.size() == 4 && descriptors[1].active && descriptors[1].kind == xlpp::WorkbookSheetKind::Chartsheet,
          "Workbook descriptors expose active mixed-sheet identity");
    const auto* loadedDashboard = static_cast<const xlpp::Workbook&>(loaded).chartsheet("Dashboard");
    check(loadedDashboard && loadedDashboard->properties().codeName() == "ChartDashboard", "Chartsheet properties reload");
    check(loadedDashboard && loadedDashboard->view().zoomScale() && *loadedDashboard->view().zoomScale() == 125,
          "Chartsheet zoom reloads");
    check(loadedDashboard && loadedDashboard->protection().enabled() && loadedDashboard->protection().content(),
          "Chartsheet protection reloads");
    check(loadedDashboard && loadedDashboard->hasPageSetup() && loadedDashboard->pageSetup().orientation() == xlpp::PageOrientation::Landscape,
          "Chartsheet page setup reloads");
    check(loadedDashboard && loadedDashboard->hasHeaderFooter() && loadedDashboard->headerFooter().oddHeader() == "&CAdvanced Header",
          "Chartsheet header reloads");

    bool rejectedLastVisible = false;
    try { loaded.setWorkbookSheetVisibility(0, xlpp::WorkbookSheetVisibility::Hidden); loaded.setWorkbookSheetVisibility(1, xlpp::WorkbookSheetVisibility::Hidden); }
    catch (const std::logic_error&) { rejectedLastVisible = true; }
    check(rejectedLastVisible, "Workbook refuses to hide its final visible sheet");

    cleanup(path);
}

void testImportedChartsheetMetadataPatchPreservesChartTree() {
    const auto fixture = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl_advanced_chartsheet_template.xltx";
    const auto path = tempPath("xlpp_p1u_imported_chartsheet", ".xltx");
    const auto secondPath = tempPath("xlpp_p1u_imported_chartsheet_second", ".xltx");
    check(std::filesystem::exists(fixture), "Advanced openpyxl Chartsheet fixture is present");
    if (!std::filesystem::exists(fixture)) return;

    const auto sourceZip = xlpp::internal::ZipArchive::open(fixture);
    xlpp::Workbook wb;
    wb.load(fixture);
    check(wb.isTemplate(), "Imported openpyxl XLTX is recognized as template");
    checkEqual(wb.activeWorkbookSheetIndex(), std::size_t{1}, "Imported active Chartsheet is recognized");
    auto* cs = wb.chartsheet("Dashboard");
    check(cs && cs->imported(), "Advanced Chartsheet imports preservation-backed");
    check(cs && cs->properties().codeName() == "ChartDashboard", "Imported Chartsheet codeName parses");
    check(cs && cs->properties().tabColor() && *cs->properties().tabColor() == "FF336699", "Imported Chartsheet tabColor parses");
    check(cs && cs->view().zoomScale() && *cs->view().zoomScale() == 125 && !cs->view().zoomToFit(), "Imported Chartsheet view parses");
    check(cs && cs->protection().enabled() && cs->protection().passwordHash() == "CF75", "Imported Chartsheet protection parses");
    check(cs && cs->hasPageMargins() && cs->pageMargins().left() == 0.4, "Imported Chartsheet margins parse");
    check(cs && cs->hasPageSetup() && cs->pageSetup().scale() == 85, "Imported Chartsheet page setup parses");
    check(cs && cs->hasHeaderFooter() && cs->headerFooter().oddHeader() == "&CAdvanced Header", "Imported Chartsheet header/footer parses");
    check(cs && cs->headerFooter().firstHeader() == "&LFirst Header" && cs->headerFooter().firstFooter() == "&LFirst Footer",
          "Imported first-page Chartsheet header/footer parses");

    // Sheet-metadata mutation must not opt the imported chart/drawing subtree into regeneration.
    cs->view().setZoomScale(140);
    cs->pageSetup().setScale(90);
    cs->headerFooter().setOddHeader("&CPatched Header");
    wb.save(path);

    const auto savedZip = xlpp::internal::ZipArchive::open(path);
    check(savedZip.get("xl/charts/chart1.xml") == sourceZip.get("xl/charts/chart1.xml"),
          "Chartsheet metadata edit preserves imported chart bytes");
    check(savedZip.get("xl/drawings/drawing1.xml") == sourceZip.get("xl/drawings/drawing1.xml"),
          "Chartsheet metadata edit preserves imported drawing bytes");
    const auto patchedCsXml = savedZip.get("xl/chartsheets/sheet1.xml");
    check(patchedCsXml.find("zoomScale=\"140\"") != std::string::npos && patchedCsXml.find("scale=\"90\"") != std::string::npos,
          "Chartsheet metadata patch reaches the chartsheet part");
    check(patchedCsXml.find("firstHeader") != std::string::npos && patchedCsXml.find("firstFooter") != std::string::npos,
          "First-page Chartsheet header/footer survives metadata patch");

    xlpp::Workbook reloaded;
    reloaded.load(path);
    const auto* patched = static_cast<const xlpp::Workbook&>(reloaded).chartsheet("Dashboard");
    check(patched && patched->view().zoomScale() && *patched->view().zoomScale() == 140,
          "Patched Chartsheet zoom round-trips");
    check(patched && patched->pageSetup().scale() == 90, "Patched Chartsheet page setup round-trips");
    check(patched && patched->headerFooter().oddHeader() == "&CPatched Header", "Patched Chartsheet header round-trips");

    // Re-saving the same workbook object must not revert metadata to the original source XML.
    wb.save(secondPath);
    xlpp::Workbook secondReload;
    secondReload.load(secondPath);
    const auto* secondCs = static_cast<const xlpp::Workbook&>(secondReload).chartsheet("Dashboard");
    check(secondCs && secondCs->view().zoomScale() && *secondCs->view().zoomScale() == 140
          && secondCs->pageSetup().scale() == 90,
          "Repeated save retains imported Chartsheet metadata edits");
    cleanup(path);
    cleanup(secondPath);
}


void testImportedChartsheetChartMutationSurvivesRepeatedSave() {
    const auto fixture = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl_advanced_chartsheet_template.xltx";
    const auto path = tempPath("xlpp_p1u_imported_chart_mutation", ".xltx");
    const auto secondPath = tempPath("xlpp_p1u_imported_chart_mutation_second", ".xltx");
    if (!std::filesystem::exists(fixture)) return;

    xlpp::Workbook wb;
    wb.load(fixture);
    auto* cs = wb.chartsheet("Dashboard");
    check(cs != nullptr, "Imported Chartsheet exists for chart-mutation repeated-save regression");
    if (!cs) return;

    cs->chart().setTitle("P1U Mutated Chart");
    wb.save(path);
    wb.save(secondPath);

    xlpp::Workbook loaded;
    loaded.load(secondPath);
    const auto* loadedCs = static_cast<const xlpp::Workbook&>(loaded).chartsheet("Dashboard");
    check(loadedCs && loadedCs->chart().title() == "P1U Mutated Chart",
          "Repeated save retains regenerated imported Chartsheet chart mutation");

    const auto secondZip = xlpp::internal::ZipArchive::open(secondPath);
    bool mutatedChartBytesFound = false;
    for (const auto& name : secondZip.entryNames()) {
        if (name.rfind("xl/charts/chart", 0) == 0 && name.size() > 4 && name.substr(name.size() - 4) == ".xml"
            && secondZip.get(name).find("P1U Mutated Chart") != std::string::npos) {
            mutatedChartBytesFound = true;
            break;
        }
    }
    check(mutatedChartBytesFound,
          "Repeated save package contains regenerated chart bytes rather than original imported chart");
    cleanup(path);
    cleanup(secondPath);
}

void testXltmWithChartsheetAndVba() {
    const auto path = tempPath("xlpp_p1u_macro_chartsheet_template", ".xltm");
    xlpp::Workbook wb;
    wb.setTemplate(true);
    populateData(wb.addWorksheet("Data"));
    wb.addChartsheet("Dashboard", makeChart());
    wb.setActiveWorkbookSheet("Dashboard");
    wb.setVbaModuleText("Module1", "Sub Hello()\r\nEnd Sub\r\n");
    wb.save(path);

    const auto zip = xlpp::internal::ZipArchive::open(path);
    check(zip.contains("xl/vbaProject.bin"), "XLTM Chartsheet template contains VBA project");
    check(zip.get("[Content_Types].xml").find("application/vnd.ms-excel.template.macroEnabled.main+xml") != std::string::npos,
          "XLTM Chartsheet template uses macro-enabled template content type");
    check(zip.contains("xl/chartsheets/sheet1.xml"), "XLTM retains first-class Chartsheet part");

    xlpp::Workbook loaded;
    loaded.load(path);
    check(loaded.isTemplate() && loaded.hasVbaProject(), "XLTM Chartsheet template reloads template + VBA identity");
    check(loaded.chartsheetCount() == 1 && loaded.activeWorkbookSheetIndex() == 1,
          "XLTM active Chartsheet survives reload");
    cleanup(path);
}

void testActiveStateFollowsMoveAndRemoval() {
    xlpp::Workbook wb;
    wb.addWorksheet("A");
    wb.addChartsheet("B", makeChart());
    wb.addWorksheet("C");
    wb.setActiveWorkbookSheet("B");
    wb.moveWorkbookSheet(1, 0);
    checkEqual(wb.activeWorkbookSheetIndex(), std::size_t{0}, "Active tab follows mixed-sheet move");
    check(wb.removeChartsheet("B"), "Active Chartsheet can be removed while another sheet remains");
    check(wb.activeWorkbookSheetIndex() < wb.workbookSheetCount()
          && wb.workbookSheetVisibility(wb.activeWorkbookSheetIndex()) == xlpp::WorkbookSheetVisibility::Visible,
          "Active tab repairs to a surviving visible sheet after removal");
}
}

int main() {
    testTemplateMixedVisibilityAndActiveChartsheet();
    testImportedChartsheetMetadataPatchPreservesChartTree();
    testImportedChartsheetChartMutationSurvivesRepeatedSave();
    testXltmWithChartsheetAndVba();
    testActiveStateFollowsMoveAndRemoval();
    if (failures == 0) { std::cout << "P1U template/chartsheet regression: PASS\n"; return 0; }
    std::cerr << failures << " P1U template/chartsheet check(s) failed\n";
    return 1;
}
