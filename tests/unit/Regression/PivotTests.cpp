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
    const auto* semanticSheet = static_cast<const xlpp::Workbook&>(workbook).worksheet("Data");
    test.checkEqual(semanticSheet->pivotTables().size(), std::size_t{1}, "Imported LibreOffice pivot is exposed through semantic model");
    const auto& importedPivot = semanticSheet->pivotTables().front();
    test.checkEqual(importedPivot.name(), std::string("SalesPivot"), "Imported pivot name is parsed");
    test.checkEqual(importedPivot.location(), std::string("E4:H9"), "Imported pivot location is parsed");
    test.checkEqual(importedPivot.cache().sourceData(), std::string("'Data'!A1:C7"), "Imported pivot worksheet source is parsed");
    test.checkEqual(importedPivot.cache().fields().size(), std::size_t{3}, "Imported pivot cache fields are parsed");
    test.checkEqual(importedPivot.rowFields().front().name(), std::string("Region"), "Imported pivot row field is parsed");
    test.checkEqual(importedPivot.columnFields().front().name(), std::string("Quarter"), "Imported pivot column field is parsed");
    test.checkEqual(importedPivot.dataFields().front().name(), std::string("Sales"), "Imported pivot data field resolves to cache field");
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

void testImportedPivotSemanticEdit(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "libreoffice" / "pivot.xlsx";
    const auto outputPath = std::filesystem::temp_directory_path() / "xlpp_imported_pivot_semantic_edit.xlsx";
    xlpp::Workbook workbook; workbook.load(sourcePath);
    auto* sheet = workbook.worksheet("Data");
    test.checkTrue(sheet != nullptr, "Semantic pivot edit fixture loads");
    auto& pivots = sheet->pivotTables(); // explicit opt-in to model regeneration
    test.checkEqual(pivots.size(), std::size_t{1}, "Imported pivot is editable through public collection");
    pivots.front().setLocation("F5");
    pivots.front().dataFields().front().setSubtotal("average");
    pivots.front().rowFields().front().setShowAll(true);
    workbook.save(outputPath);

    const auto archive = xlpp::internal::ZipArchive::open(outputPath);
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    const auto validation = graph.validate();
    test.checkTrue(validation.ok(), "Regenerated imported pivot keeps a valid package graph");
    test.checkEqual(graph.objectInventory().pivotTables, std::size_t{1}, "Semantic edit replaces original pivot table instead of duplicating it");
    const auto sheetRels = graph.relationshipsFrom("xl/worksheets/sheet1.xml");
    auto pivotRel = std::find_if(sheetRels.begin(), sheetRels.end(), [](const auto& rel) { return rel.type.find("/pivotTable") != std::string::npos; });
    test.checkTrue(pivotRel != sheetRels.end(), "Edited pivot remains worksheet-owned");
    if (pivotRel != sheetRels.end()) {
        const auto pivotPart = xlpp::internal::RelationshipGraph::resolveTarget("xl/worksheets/sheet1.xml", pivotRel->target);
        test.checkTrue(archive.contains(pivotPart), "Edited semantic pivot target exists");
        if (archive.contains(pivotPart)) {
            const auto xml = archive.get(pivotPart);
            test.checkTrue(xml.find("ref=\"F5:") != std::string::npos, "Edited pivot location is serialized");
            test.checkTrue(xml.find("subtotal=\"average\"") != std::string::npos, "Edited pivot aggregation is serialized");
            test.checkTrue(xml.find("showAll=\"1\"") != std::string::npos, "Edited pivot field flags are serialized");
        }
    }

    xlpp::Workbook reloaded; reloaded.load(outputPath);
    const auto* loadedSheet = static_cast<const xlpp::Workbook&>(reloaded).worksheet("Data");
    test.checkTrue(loadedSheet != nullptr && loadedSheet->pivotTables().size() == 1, "Regenerated pivot is semantically readable after reload");
    if (loadedSheet && !loadedSheet->pivotTables().empty()) {
        const auto& pivot = loadedSheet->pivotTables().front();
        test.checkTrue(pivot.location().rfind("F5:", 0) == 0, "Reloaded pivot retains edited location");
        test.checkEqual(pivot.dataFields().front().subtotal(), std::string("average"), "Reloaded pivot retains edited aggregation");
    }
    std::filesystem::remove(outputPath);
}

void testPivotSemanticOptionsRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_pivot_semantic_options.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("PivotOptions");
    sheet.append({std::string("Region"), std::string("Revenue"), std::string("Channel")});
    sheet.append({std::string("East"), 10.0, std::string("Retail")});
    sheet.append({std::string("West"), 20.0, std::string("Online")});

    xlpp::PivotTable pivot("OptionsPivot");
    pivot.setLocation("E3");
    pivot.setLayout(xlpp::PivotLayout::Outline);
    pivot.setRowGrandTotals(false);
    pivot.setColumnGrandTotals(true);
    pivot.setPreserveFormatting(false);
    pivot.setUseAutoFormatting(false);
    pivot.setDataCaption("Measures");
    pivot.setStyleName("PivotStyleMedium9");
    pivot.setShowRowHeaders(false);
    pivot.setShowColumnHeaders(true);
    pivot.setShowRowStripes(true);
    pivot.setShowColumnStripes(true);
    pivot.setShowLastColumn(false);
    pivot.cache().setSourceData("'PivotOptions'!$A$1:$C$3");
    pivot.cache().setRefreshOnLoad(false);
    pivot.cache().setSaveData(true);
    pivot.cache().setEnableRefresh(false);
    auto& row = pivot.addRowField("Region");
    row.setSortType(2);
    row.setSubtotalTop(false);
    row.setInsertBlankRow(true);
    row.setIncludeNewItemsInFilter(true);
    auto& page = pivot.addPageField("Channel");
    page.setSelectedItemIndex(1);
    page.setMultipleItemSelectionAllowed(true);
    auto& data = pivot.addDataField("Revenue", "average");
    data.setCaption("Average Revenue");
    data.setNumberFormatId(4);
    data.setShowDataAs("percentOfTotal");
    data.setBaseField(1);
    data.setBaseItem(2);
    sheet.addPivotTable(std::move(pivot));
    workbook.save(path);

    const auto archive = xlpp::internal::ZipArchive::open(path);
    const auto pivotXml = archive.get("xl/pivotTables/pivotTable1.xml");
    const auto cacheXml = archive.get("xl/pivotCache/pivotCacheDefinition1.xml");
    test.checkTrue(pivotXml.find("compact=\"0\" compactData=\"0\" outline=\"1\" outlineData=\"1\"") != std::string::npos,
                   "Outline pivot layout is serialized explicitly");
    test.checkTrue(pivotXml.find("rowGrandTotals=\"0\"") != std::string::npos && pivotXml.find("colGrandTotals=\"1\"") != std::string::npos,
                   "Pivot grand-total settings are serialized");
    test.checkTrue(pivotXml.find("name=\"PivotStyleMedium9\"") != std::string::npos && pivotXml.find("showRowStripes=\"1\"") != std::string::npos,
                   "Pivot style metadata is serialized");
    test.checkTrue(pivotXml.find("sortType=\"descending\"") != std::string::npos && pivotXml.find("insertBlankRow=\"1\"") != std::string::npos,
                   "Pivot field semantic options are serialized");
    test.checkTrue(pivotXml.find("<pageField fld=\"2\" hier=\"-1\" item=\"1\"/>") != std::string::npos &&
                   pivotXml.find("multipleItemSelectionAllowed=\"1\"") != std::string::npos,
                   "Pivot report-filter selection metadata is serialized");
    test.checkTrue(pivotXml.find("name=\"Average Revenue\"") != std::string::npos && pivotXml.find("numFmtId=\"4\"") != std::string::npos && pivotXml.find("showDataAs=\"percentOfTotal\"") != std::string::npos,
                   "Pivot data field formatting and show-data-as settings are serialized");
    test.checkTrue(cacheXml.find("refreshOnLoad=\"0\"") != std::string::npos && cacheXml.find("enableRefresh=\"0\"") != std::string::npos,
                   "Pivot cache lifecycle options are serialized");

    xlpp::Workbook reloaded; reloaded.load(path);
    const auto* loadedSheet = static_cast<const xlpp::Workbook&>(reloaded).worksheet("PivotOptions");
    test.checkTrue(loadedSheet && loadedSheet->pivotTables().size() == 1, "Generated pivot semantic model reloads");
    if (loadedSheet && !loadedSheet->pivotTables().empty()) {
        const auto& loaded = loadedSheet->pivotTables().front();
        test.checkTrue(loaded.layout() == xlpp::PivotLayout::Outline, "Pivot layout reloads semantically");
        test.checkTrue(!loaded.rowGrandTotals() && loaded.columnGrandTotals(), "Pivot grand totals reload semantically");
        test.checkEqual(loaded.styleName(), std::string("PivotStyleMedium9"), "Pivot style name reloads");
        test.checkTrue(loaded.showRowStripes() && loaded.showColumnStripes() && !loaded.showLastColumn(), "Pivot style flags reload");
        test.checkTrue(!loaded.cache().refreshOnLoad() && !loaded.cache().enableRefresh(), "Pivot cache options reload");
        test.checkEqual(loaded.rowFields().front().sortType(), 2, "Pivot field sort type reloads");
        test.checkTrue(!loaded.rowFields().front().subtotalTop() && loaded.rowFields().front().insertBlankRow(), "Pivot field layout flags reload");
        test.checkTrue(!loaded.pageFields().empty() && loaded.pageFields().front().selectedItemIndex() == 1, "Pivot report-filter selected item reloads");
        test.checkTrue(!loaded.pageFields().empty() && loaded.pageFields().front().multipleItemSelectionAllowed(), "Pivot report-filter multi-select flag reloads");
        test.checkEqual(loaded.dataFields().front().caption(), std::string("Average Revenue"), "Pivot data caption reloads");
        test.checkEqual(loaded.dataFields().front().numberFormatId(), std::uint32_t{4}, "Pivot data number format reloads");
        test.checkEqual(loaded.dataFields().front().showDataAs(), std::string("percentOfTotal"), "Pivot show-data-as reloads");
    }
    std::filesystem::remove(path);
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
    const auto rowFieldStart = tableXml.find("<pivotField axis=\"axisRow\"");
    const auto rowFieldEnd = rowFieldStart == std::string::npos ? std::string::npos : tableXml.find("</pivotField>", rowFieldStart);
    test.checkTrue(rowFieldStart != std::string::npos && rowFieldEnd != std::string::npos &&
                   tableXml.substr(rowFieldStart, rowFieldEnd - rowFieldStart).find("<items count=\"2\">") != std::string::npos,
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

void testCompletePivotGenerationRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_complete_pivot_roundtrip.xlsx";
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Pivot");
    xlpp::PivotTable pivot("CompletePivot");
    pivot.setLocation("H3");
    pivot.setLayout(xlpp::PivotLayout::Tabular);
    pivot.setShowEmptyRow(true); pivot.setShowEmptyColumn(true); pivot.setMultipleFieldFilters(true);
    pivot.setShowValuesRow(true); pivot.setSubtotalHiddenItems(true); pivot.setPageWrap(2); pivot.setPageOverThenDown(true);
    auto& cache = pivot.cache();
    cache.setSourceData("'Pivot'!$A$1:$E$5");
    cache.setFields({"Region", "Year", "Quarter", "Sales", "Units"});
    cache.setRefreshOnLoad(false); cache.setSaveData(true); cache.setEnableRefresh(true); cache.setMissingItemsLimit(1000);
    cache.setOptimizeMemory(true); cache.setUpgradeOnRefresh(true); cache.setRefreshedBy("XL++ regression");
    cache.addRecord({"East", "2024", "Q1", "100", "1"});
    cache.addRecord({"East", "2025", "Q2", "120", "2"});
    cache.addRecord({"West", "2024", "Q1", "90", "1"});
    cache.addRecord({"West", "2025", "Q2", "130", "3"});

    auto& region = pivot.addRowField("Region");
    region.setRepeatItemLabels(true); region.setItemHidden(1); region.addSubtotal("sum"); region.setDefaultSubtotal(false);
    auto& year = pivot.addRowField("Year");
    xlpp::PivotGrouping yearGrouping; yearGrouping.kind = xlpp::PivotGrouping::Kind::Numeric;
    yearGrouping.autoStart = false; yearGrouping.autoEnd = false; yearGrouping.start = 2024; yearGrouping.end = 2025; yearGrouping.interval = 1;
    year.setGrouping(yearGrouping);
    pivot.addColumnField("Quarter");
    auto& sales = pivot.addDataField("Sales", "sum"); sales.setCaption("Total Sales");
    auto& units = pivot.addDataField("Units", "average"); units.setCaption("Average Units"); units.setShowDataAs("rankAscending"); units.setBaseField(0);
    xlpp::PivotFilter captionFilter; captionFilter.type = "captionEqual"; captionFilter.fieldIndex = 0; captionFilter.value1 = "East"; pivot.addFilter(captionFilter);
    xlpp::PivotFilter topFilter; topFilter.type = "count"; topFilter.fieldIndex = 0; topFilter.measureFieldIndex = 0; topFilter.top10Value = 5; pivot.addFilter(topFilter);
    sheet.addPivotTable(std::move(pivot));

    workbook.save(path);
    const auto archive = xlpp::internal::ZipArchive::open(path);
    const auto pivotXml = archive.get("xl/pivotTables/pivotTable1.xml");
    const auto cacheXml = archive.get("xl/pivotCache/pivotCacheDefinition1.xml");
    const auto recordsXml = archive.get("xl/pivotCache/pivotCacheRecords1.xml");
    test.checkTrue(pivotXml.find("xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"") != std::string::npos &&
                   pivotXml.find("<x14:pivotField fillDownLabels=\"1\"/>") != std::string::npos,
                   "Repeat-item-labels uses the Excel x14 pivot-field extension");
    test.checkTrue(pivotXml.find("<item x=\"1\" h=\"1\"/>") != std::string::npos,
                   "Hidden pivot items serialize on the pivot field item list");
    test.checkTrue(pivotXml.find("<field x=\"-2\"/>") != std::string::npos,
                   "Multiple data fields add Excel's synthetic Values column field");
    test.checkTrue(pivotXml.find("{E15A36E0-9728-4E99-A89B-3F7291B0FE68}") != std::string::npos &&
                   pivotXml.find("<x14:dataField pivotShowAs=\"rankAscending\"/>") != std::string::npos &&
                   pivotXml.find("showDataAs=\"rankAscending\"") == std::string::npos,
                   "Modern Excel Show Values As modes use the x14 dataField extension");
    test.checkTrue(pivotXml.find("captionEqual") != std::string::npos && pivotXml.find("<top10 top=\"1\" percent=\"0\" val=\"5\"/>") != std::string::npos,
                   "Caption and Top-N pivot filters serialize");
    test.checkTrue(pivotXml.find("multipleFieldFilters=\"1\"") != std::string::npos && pivotXml.find("showValuesRow=\"1\"") != std::string::npos &&
                   pivotXml.find("pageWrap=\"2\"") != std::string::npos && pivotXml.find("pageOverThenDown=\"1\"") != std::string::npos,
                   "Extended pivot-table behavior flags serialize");
    test.checkTrue(cacheXml.find("missingItemsLimit=\"1000\"") != std::string::npos && cacheXml.find("optimizeMemory=\"1\"") != std::string::npos &&
                   cacheXml.find("upgradeOnRefresh=\"1\"") != std::string::npos && cacheXml.find("refreshedBy=\"XL++ regression\"") != std::string::npos,
                   "Extended pivot-cache lifecycle metadata serializes");
    test.checkTrue(cacheXml.find("<fieldGroup base=\"1\"><rangePr autoStart=\"0\" autoEnd=\"0\" startNum=\"2024\" endNum=\"2025\" groupBy=\"range\" groupInterval=\"1\"/>") != std::string::npos,
                   "Numeric pivot grouping serializes into cache fieldGroup/rangePr");
    test.checkTrue(recordsXml.find("count=\"4\"") != std::string::npos && recordsXml.find("<x v=\"0\"/>") != std::string::npos && recordsXml.find("<n v=\"100\"/>") != std::string::npos,
                   "Pivot cache records serialize shared-item indexes and typed data values");
    const auto rowItems = xlpp::internal::tags(pivotXml, "rowItems");
    test.checkTrue(!rowItems.empty(), "Pivot row items are emitted");
    if (!rowItems.empty()) {
        const auto rows = xlpp::internal::tags(rowItems.front(), "i");
        test.checkTrue(rows.size() >= 3, "Multi-field pivot emits Cartesian row tuples plus totals");
    }

    xlpp::Workbook reloaded; reloaded.load(path);
    const auto* loadedSheet = reloaded.worksheet("Pivot");
    test.checkTrue(loadedSheet && loadedSheet->pivotTables().size() == 1, "Complete generated pivot reloads");
    if (loadedSheet && !loadedSheet->pivotTables().empty()) {
        const auto& loaded = loadedSheet->pivotTables().front();
        test.checkEqual(loaded.cache().records().size(), std::size_t{4}, "Pivot cache records reload into the semantic model");
        test.checkEqual(loaded.cache().records()[0][3], std::string("100"), "Typed pivot cache numeric values reload without losing content");
        test.checkTrue(loaded.rowFields().front().repeatItemLabels() && loaded.rowFields().front().itemHidden(1), "Repeat labels and hidden item state reload");
        test.checkTrue(loaded.rowFields()[1].grouping().kind == xlpp::PivotGrouping::Kind::Numeric && loaded.rowFields()[1].grouping().interval == 1.0,
                       "Numeric pivot grouping reloads");
        test.checkEqual(loaded.filters().size(), std::size_t{2}, "Pivot filters reload");
        test.checkTrue(loaded.dataFields().size() == 2 && loaded.dataFields()[1].showDataAs() == "rankAscending",
                       "Modern x14 Show Values As mode reloads into the semantic model");
        test.checkTrue(loaded.multipleFieldFilters() && loaded.showValuesRow() && loaded.subtotalHiddenItems(), "Extended pivot behavior flags reload");
    }
    std::filesystem::remove(path);
}
