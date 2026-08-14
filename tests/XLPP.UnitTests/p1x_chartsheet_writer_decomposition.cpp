#include "Packaging/RelationshipGraph.h"
#include "Packaging/ZipArchive.h"
#include <XLPP/XLPP.h>

#include <chrono>
#include <filesystem>
#include <iostream>
#include <string>

namespace {
int failures = 0;
void check(bool ok, const char* message) {
    if (!ok) { ++failures; std::cerr << "FAIL: " << message << '\n'; }
}
std::filesystem::path tempPath(const char* stem) {
    const auto tick = std::chrono::high_resolution_clock::now().time_since_epoch().count();
    return std::filesystem::temp_directory_path() / (std::string(stem) + "_" + std::to_string(tick) + ".xltx");
}
std::string firstEntry(const xlpp::internal::ZipArchive& zip, const std::string& prefix, const std::string& suffix) {
    for (const auto& name : zip.entryNames())
        if (name.rfind(prefix, 0) == 0 && name.size() >= suffix.size() && name.compare(name.size()-suffix.size(), suffix.size(), suffix)==0)
            return name;
    return {};
}
xlpp::Chart makeChart() {
    xlpp::Chart chart(xlpp::Chart::Type::Line);
    chart.setTitle("P1X Writer");
    xlpp::ChartSeries s("Values");
    s.setCategoriesReference("'Data'!$A$1:$A$2");
    s.setValuesReference("'Data'!$B$1:$B$2");
    chart.addSeries(std::move(s));
    return chart;
}
void populate(xlpp::Workbook& wb) {
    auto& ws = wb.addWorksheet("Data");
    ws.append({std::string("A"), 1.0});
    ws.append({std::string("B"), 2.0});
}

void testPreservedRelationshipCollisionReallocatesGeneratedDrawingId() {
    const auto seedPath = tempPath("xlpp_p1x_collision_seed");
    const auto fixturePath = tempPath("xlpp_p1x_collision_fixture");
    const auto outputPath = tempPath("xlpp_p1x_collision_output");
    const auto repeatedPath = tempPath("xlpp_p1x_collision_repeated");

    xlpp::Workbook seed;
    seed.setTemplate(true);
    populate(seed);
    seed.addChartsheet("Dashboard", makeChart());
    seed.save(seedPath);

    auto fixture = xlpp::internal::ZipArchive::open(seedPath);
    const auto sheetPart = firstEntry(fixture, "xl/chartsheets/sheet", ".xml");
    const auto relsPart = xlpp::internal::RelationshipGraph::relationshipsPartForSource(sheetPart);
    auto sheetXml = fixture.get(sheetPart);
    auto relsXml = fixture.get(relsPart);

    // Move the imported drawing away from rId1, then deliberately reserve rId1
    // for a preserved auxiliary VML relationship. On chart regeneration the
    // generated writer starts with rId1 and must relocate it while updating the
    // owner XML, otherwise two relationships would share one ID.
    const auto drawingIdPos = sheetXml.find("r:id=\"rId1\"");
    check(drawingIdPos != std::string::npos, "Seed Chartsheet has drawing rId1");
    if (drawingIdPos != std::string::npos)
        sheetXml.replace(drawingIdPos, std::string("r:id=\"rId1\"").size(), "r:id=\"rIdDrawing\"");

    const auto relIdPos = relsXml.find("Id=\"rId1\"");
    check(relIdPos != std::string::npos, "Seed relationship has rId1");
    if (relIdPos != std::string::npos)
        relsXml.replace(relIdPos, std::string("Id=\"rId1\"").size(), "Id=\"rIdDrawing\"");
    const auto close = relsXml.rfind("</Relationships>");
    relsXml.insert(close,
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawingHF1.vml\"/>");
    const auto sheetClose = sheetXml.rfind("</chartsheet>");
    sheetXml.insert(sheetClose, "<legacyDrawingHF r:id=\"rId1\"/>");
    fixture.replace(sheetPart, sheetXml);
    fixture.replace(relsPart, relsXml);
    fixture.add("xl/drawings/vmlDrawingHF1.vml", "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape id=\"HF\"/></xml>");
    fixture.save(fixturePath);

    xlpp::Workbook wb;
    wb.load(fixturePath);
    auto* cs = wb.chartsheet("Dashboard");
    check(cs != nullptr, "Collision fixture Chartsheet loads");
    if (!cs) return;
    cs->chart().setTitle("Regenerated after collision");
    wb.save(outputPath);

    auto output = xlpp::internal::ZipArchive::open(outputPath);
    const auto outSheet = firstEntry(output, "xl/chartsheets/sheet", ".xml");
    const auto outRels = xlpp::internal::RelationshipGraph::relationshipsPartForSource(outSheet);
    const auto& outSheetXml = output.get(outSheet);
    const auto& outRelsXml = output.get(outRels);
    check(outSheetXml.find("legacyDrawingHF r:id=\"rId1\"") != std::string::npos,
          "Preserved auxiliary relationship retains rId1");
    check(outSheetXml.find("<drawing r:id=\"rIdXLPP") != std::string::npos,
          "Generated drawing relationship is reallocated after collision");
    check(outRelsXml.find("Id=\"rId1\"") != std::string::npos && outRelsXml.find("/vmlDrawing") != std::string::npos,
          "Preserved VML relationship remains present");
    check(outRelsXml.find("Id=\"rIdXLPP") != std::string::npos && outRelsXml.find("/drawing") != std::string::npos,
          "Reallocated drawing relationship is serialized");
    check(output.contains("xl/drawings/vmlDrawingHF1.vml"), "Auxiliary VML part survives collision repair");
    check(xlpp::internal::RelationshipGraph::fromArchive(output).validate().ok(), "Collision-repaired package graph validates");

    wb.save(repeatedPath);
    auto repeated = xlpp::internal::ZipArchive::open(repeatedPath);
    check(xlpp::internal::RelationshipGraph::fromArchive(repeated).validate().ok(), "Repeated save after relationship collision remains valid");
    const auto repeatedSheet = firstEntry(repeated, "xl/chartsheets/sheet", ".xml");
    check(repeated.get(repeatedSheet).find("legacyDrawingHF r:id=\"rId1\"") != std::string::npos,
          "Repeated save keeps preserved VML owner binding");

    std::filesystem::remove(seedPath);
    std::filesystem::remove(fixturePath);
    std::filesystem::remove(outputPath);
    std::filesystem::remove(repeatedPath);
}
}

int main() {
    testPreservedRelationshipCollisionReallocatesGeneratedDrawingId();
    if (failures == 0) { std::cout << "P1X Chartsheet writer/decomposition regression: PASS\n"; return 0; }
    std::cerr << failures << " P1X check(s) failed\n";
    return 1;
}
