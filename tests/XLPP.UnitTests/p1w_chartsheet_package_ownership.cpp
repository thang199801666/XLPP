#include "Packaging/RelationshipGraph.h"
#include "Packaging/ZipArchive.h"
#include "XML/XmlUtilities.h"
#include <XLPP/XLPP.h>

#include <algorithm>
#include <chrono>
#include <cstdlib>
#include <filesystem>
#include <iostream>
#include <string>
#include <vector>

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
    if (std::getenv("XLPP_KEEP_P1W_ARTIFACTS")) std::cout << "P1W_ARTIFACT=" << p.string() << '\n';
    else std::filesystem::remove(p);
}

xlpp::Chart makeChart(std::string title = "P1W Chartsheet") {
    xlpp::Chart chart(xlpp::Chart::Type::Line);
    chart.setTitle(std::move(title));
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

std::string firstEntry(const xlpp::internal::ZipArchive& zip, const std::string& prefix, const std::string& suffix) {
    for (const auto& name : zip.entryNames())
        if (name.rfind(prefix, 0) == 0 && name.size() >= suffix.size() && name.compare(name.size() - suffix.size(), suffix.size(), suffix) == 0)
            return name;
    return {};
}

void assertPackageValid(const xlpp::internal::ZipArchive& zip, const char* message) {
    const auto report = xlpp::internal::RelationshipGraph::fromArchive(zip).validate();
    check(report.ok(), message);
}

void testGeneratedPrinterSettingsRoundTripAndClear() {
    const auto first = tempPath("xlpp_p1w_printer", ".xltx");
    const auto second = tempPath("xlpp_p1w_printer_mutated", ".xltx");
    const auto cleared = tempPath("xlpp_p1w_printer_cleared", ".xltx");
    const std::string payload{"\x01\x00\x7F\x55\x10\x00\x22", 7};
    const std::string payload2{"\x02\x00\x6A\x33\x44\x00\x11\x7E", 8};

    xlpp::Workbook wb;
    wb.setTemplate(true);
    populate(wb);
    auto& cs = wb.addChartsheet("Dashboard", makeChart());
    cs.pageSetup().setOrientation(xlpp::PageOrientation::Landscape);
    cs.setPrinterSettingsData(payload);
    wb.save(first);

    auto zip = xlpp::internal::ZipArchive::open(first);
    const auto chartSheetPart = firstEntry(zip, "xl/chartsheets/sheet", ".xml");
    const auto printerPart = firstEntry(zip, "xl/printerSettings/printerSettings", ".bin");
    check(!chartSheetPart.empty() && !printerPart.empty(), "Generated Chartsheet owns a printerSettings part");
    check(zip.get(printerPart) == payload, "Generated printerSettings bytes are exact");
    const auto chartSheetXml = zip.get(chartSheetPart);
    check(chartSheetXml.find("<pageSetup") != std::string::npos && chartSheetXml.find("r:id=") != std::string::npos,
          "Generated Chartsheet pageSetup references printerSettings");
    const auto relsPart = xlpp::internal::RelationshipGraph::relationshipsPartForSource(chartSheetPart);
    check(zip.contains(relsPart) && zip.get(relsPart).find("/printerSettings") != std::string::npos,
          "Generated Chartsheet relationship part owns printerSettings");
    check(zip.get("[Content_Types].xml").find("spreadsheetml.printerSettings") != std::string::npos,
          "Generated printerSettings content type is emitted");
    assertPackageValid(zip, "Generated printerSettings package validates");

    xlpp::Workbook loaded;
    loaded.load(first);
    auto* imported = loaded.chartsheet("Dashboard");
    check(imported && imported->hasPrinterSettings(), "Generated printerSettings reloads into Chartsheet model");
    if (!imported) return;
    check(imported->printerSettingsData() && *imported->printerSettingsData() == payload,
          "Reloaded printerSettings payload remains byte-exact");

    imported->setPrinterSettingsData(payload2);
    loaded.save(second);
    auto secondZip = xlpp::internal::ZipArchive::open(second);
    const auto secondPrinter = firstEntry(secondZip, "xl/printerSettings/printerSettings", ".bin");
    check(!secondPrinter.empty() && secondZip.get(secondPrinter) == payload2,
          "Imported printerSettings can be replaced without regenerating chart data");
    assertPackageValid(secondZip, "Mutated printerSettings package validates");

    imported->clearPrinterSettings();
    loaded.save(cleared);
    auto clearedZip = xlpp::internal::ZipArchive::open(cleared);
    check(firstEntry(clearedZip, "xl/printerSettings/printerSettings", ".bin").empty(),
          "Clearing printerSettings retires the old binary part");
    const auto clearedSheet = firstEntry(clearedZip, "xl/chartsheets/sheet", ".xml");
    const auto clearedPageSetups = xlpp::internal::tags(clearedZip.get(clearedSheet), "pageSetup");
    check(clearedPageSetups.empty() || xlpp::internal::attribute(clearedPageSetups.front(), "r:id").empty(),
          "Clearing printerSettings removes pageSetup printer relationship reference");
    assertPackageValid(clearedZip, "Cleared printerSettings package validates");

    cleanup(first); cleanup(second); cleanup(cleared);
}

std::filesystem::path makeLegacyDrawingHFFixture() {
    const auto seedPath = tempPath("xlpp_p1w_hf_seed", ".xltx");
    const auto fixturePath = tempPath("xlpp_p1w_hf_fixture", ".xltx");
    xlpp::Workbook seed;
    seed.setTemplate(true);
    populate(seed);
    seed.addChartsheet("Dashboard", makeChart("Header Picture"));
    seed.save(seedPath);

    auto zip = xlpp::internal::ZipArchive::open(seedPath);
    const auto chartSheetPart = firstEntry(zip, "xl/chartsheets/sheet", ".xml");
    const auto relsPart = xlpp::internal::RelationshipGraph::relationshipsPartForSource(chartSheetPart);
    auto sheetXml = zip.get(chartSheetPart);
    auto relsXml = zip.get(relsPart);
    const std::string hfNode = "<legacyDrawingHF r:id=\"rIdHF\"/>";
    const auto closeSheet = sheetXml.rfind("</chartsheet>");
    check(closeSheet != std::string::npos, "Fixture chartsheet close tag exists");
    sheetXml.insert(closeSheet, hfNode);
    const auto closeRels = relsXml.rfind("</Relationships>");
    check(closeRels != std::string::npos, "Fixture relationship close tag exists");
    relsXml.insert(closeRels,
        "<Relationship Id=\"rIdHF\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawingHF1.vml\"/>");
    zip.replace(chartSheetPart, std::move(sheetXml));
    zip.replace(relsPart, std::move(relsXml));
    zip.add("xl/drawings/vmlDrawingHF1.vml",
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape id=\"HFLogo\"/></xml>");
    zip.save(fixturePath);
    cleanup(seedPath);
    return fixturePath;
}

void checkLegacyDrawingHFOwnership(const xlpp::internal::ZipArchive& zip, const char* prefix) {
    const auto sheetPart = firstEntry(zip, "xl/chartsheets/sheet", ".xml");
    check(!sheetPart.empty(), prefix);
    if (sheetPart.empty()) return;
    const auto xml = zip.get(sheetPart);
    check(xml.find("<legacyDrawingHF") != std::string::npos && xml.find("rIdHF") != std::string::npos,
          "Chartsheet legacyDrawingHF owner node is preserved");
    const auto relsPart = xlpp::internal::RelationshipGraph::relationshipsPartForSource(sheetPart);
    check(zip.contains(relsPart) && zip.get(relsPart).find("rIdHF") != std::string::npos &&
          zip.get(relsPart).find("/vmlDrawing") != std::string::npos,
          "Chartsheet legacyDrawingHF relationship is preserved");
    check(zip.contains("xl/drawings/vmlDrawingHF1.vml") && zip.get("xl/drawings/vmlDrawingHF1.vml").find("HFLogo") != std::string::npos,
          "Chartsheet legacyDrawingHF VML part is preserved byte-semantically");
    assertPackageValid(zip, "Chartsheet legacyDrawingHF package validates");
}

void testLegacyDrawingHFPreservationAcrossMetadataAndChartRegeneration() {
    const auto fixture = makeLegacyDrawingHFFixture();
    const auto metadata = tempPath("xlpp_p1w_hf_metadata", ".xltx");
    const auto regenerated = tempPath("xlpp_p1w_hf_regenerated", ".xltx");
    const auto repeated = tempPath("xlpp_p1w_hf_repeated", ".xltx");

    xlpp::Workbook wb;
    wb.load(fixture);
    auto* cs = wb.chartsheet("Dashboard");
    check(cs && cs->imported(), "legacyDrawingHF fixture loads as imported Chartsheet");
    if (!cs) return;

    cs->view().setZoomScale(125);
    wb.save(metadata);
    checkLegacyDrawingHFOwnership(xlpp::internal::ZipArchive::open(metadata),
                                  "Metadata-edited Chartsheet exists");

    cs->chart().setTitle("Regenerated Chart");
    wb.save(regenerated);
    auto regeneratedZip = xlpp::internal::ZipArchive::open(regenerated);
    checkLegacyDrawingHFOwnership(regeneratedZip, "Chart-regenerated Chartsheet exists");
    check(!regeneratedZip.contains("xl/chartsheets/sheet1.xml") ||
          firstEntry(regeneratedZip, "xl/chartsheets/sheet", ".xml") == "xl/chartsheets/sheet1.xml",
          "Old Chartsheet owner is not duplicated after chart regeneration");

    wb.save(repeated);
    auto repeatedZip = xlpp::internal::ZipArchive::open(repeated);
    checkLegacyDrawingHFOwnership(repeatedZip, "Repeated-save Chartsheet exists");
    const auto repeatedSheet = firstEntry(repeatedZip, "xl/chartsheets/sheet", ".xml");
    const auto repeatedRels = xlpp::internal::RelationshipGraph::relationshipsPartForSource(repeatedSheet);
    check(std::count(repeatedZip.get(repeatedRels).begin(), repeatedZip.get(repeatedRels).end(), '\0') == 0,
          "Repeated relationship serialization remains textual");

    cleanup(fixture); cleanup(metadata); cleanup(regenerated); cleanup(repeated);
}

void testPrinterSettingsSurviveImportedChartRegenerationAndRemoval() {
    const auto source = tempPath("xlpp_p1w_printer_chart_source", ".xltx");
    const auto regenerated = tempPath("xlpp_p1w_printer_chart_regenerated", ".xltx");
    const auto removed = tempPath("xlpp_p1w_printer_chart_removed", ".xltx");
    const std::string payload{"\x33\x00\x44\x55\x66", 5};

    xlpp::Workbook seed;
    seed.setTemplate(true);
    populate(seed);
    auto& cs = seed.addChartsheet("Dashboard", makeChart());
    cs.setPrinterSettingsData(payload);
    seed.save(source);

    xlpp::Workbook wb;
    wb.load(source);
    auto* imported = wb.chartsheet("Dashboard");
    check(imported && imported->hasPrinterSettings(), "Imported Chartsheet discovers printerSettings before chart regeneration");
    if (!imported) return;
    imported->chart().setTitle("Changed but keep printer");
    wb.save(regenerated);
    auto zip = xlpp::internal::ZipArchive::open(regenerated);
    const auto printerPart = firstEntry(zip, "xl/printerSettings/printerSettings", ".bin");
    check(!printerPart.empty() && zip.get(printerPart) == payload,
          "Chart regeneration preserves untouched printerSettings bytes");
    assertPackageValid(zip, "Chart-regenerated printerSettings package validates");

    check(wb.removeChartsheet("Dashboard"), "Imported Chartsheet can be removed after regeneration");
    wb.save(removed);
    auto removedZip = xlpp::internal::ZipArchive::open(removed);
    check(firstEntry(removedZip, "xl/printerSettings/printerSettings", ".bin").empty(),
          "Removing imported Chartsheet retires printerSettings closure");
    assertPackageValid(removedZip, "Chartsheet removal after printerSettings validates");

    cleanup(source); cleanup(regenerated); cleanup(removed);
}

} // namespace

int main() {
    testGeneratedPrinterSettingsRoundTripAndClear();
    testLegacyDrawingHFPreservationAcrossMetadataAndChartRegeneration();
    testPrinterSettingsSurviveImportedChartRegenerationAndRemoval();
    if (failures == 0) { std::cout << "P1W Chartsheet package-ownership regression: PASS\n"; return 0; }
    std::cerr << failures << " P1W Chartsheet package-ownership check(s) failed\n";
    return 1;
}
