#include <unordered_set>
#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Worksheet/WorksheetName.h>
#include "Package/Xml/XmlUtilities.h"
#include "Package/Zip/ZipArchive.h"
#include "Package/Opc/RelationshipGraph.h"
#include "Core/Threading/ThreadPool.h"
#include "VBA/VbaProjectBinary.h"
#include "Encryption/OfficeEncryption.h"
#include "IO/FileTransaction.h"
#include "IO/BinaryFile.h"
#include "Preservation/XmlMergeSupport.h"
#include "Preservation/PackageClosure.h"
#include "OOXML/Common/Namespaces.h"
#include "OOXML/Common/PackageRelationships.h"
#include "OOXML/Styles/StyleCodec.h"
#include "OOXML/Drawings/WorkbookDrawingAccess.h"
#include "OOXML/Drawings/PreservedDrawingEditor.h"
#include "OOXML/Charts/ImportedChartPatcher.h"
#include "OOXML/Charts/ChartSerializer.h"
#include "OOXML/Worksheet/WorksheetWriter.h"
#include "OOXML/Pivot/PivotCodec.h"
#include "OOXML/Comments/CommentCodec.h"
#include "OOXML/Drawings/DrawingWriter.h"
#include "OOXML/Tables/TableWriter.h"
#include "OOXML/Worksheet/WorksheetBatchWriter.h"
#include "OOXML/Workbook/WorkbookReferenceSupport.h"
#include <algorithm>
#include <array>
#include <cctype>
#include <cmath>
#include <cstdlib>
#include <fstream>
#include <functional>
#include <sstream>
#include <set>
#include <stdexcept>
#include <thread>
#include <type_traits>
#include <tuple>
#include <unordered_map>
#include <vector>
#include <iomanip>
#include <limits>
#include <random>
#include <atomic>
#include <cerrno>
#include <system_error>
#ifdef _WIN32
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#ifdef DocumentProperties
#undef DocumentProperties
#endif
#endif

using xlpp::internal::xmlEscape;
using xlpp::internal::writeXmlEscaped;
using xlpp::internal::ooxml::DxfCatalog;
using xlpp::internal::ooxml::StyleCatalog;
using xlpp::internal::ooxml::nsMain;
using xlpp::internal::ooxml::nsRelsPkg;
using xlpp::internal::ooxml::nsRelsDoc;
using xlpp::internal::ooxml::nsCtPkg;
using xlpp::internal::ooxml::nsCoreProps;
using xlpp::internal::ooxml::nsExtProps;
using xlpp::internal::ooxml::nsVTypes;
using xlpp::internal::ooxml::stylesXml;
using xlpp::internal::ooxml::mergeRelationshipsXml;
using xlpp::internal::preservation::extractTagBlocks;
using xlpp::internal::preservation::eraseTagBlocks;
using xlpp::internal::preservation::joinBlocks;
using xlpp::internal::preservation::insertBefore;
using xlpp::internal::ooxml::sheetXml;
using xlpp::internal::ooxml::pivotTableXml;
using xlpp::internal::ooxml::pivotCacheXml;
using xlpp::internal::ooxml::pivotCacheRecordsXml;
using xlpp::internal::ooxml::effectivePivotTable;
using xlpp::internal::ooxml::commentsXml;
using xlpp::internal::ooxml::commentsVml;
using xlpp::internal::ooxml::drawingXml;
using xlpp::internal::ooxml::drawingRelationshipsXml;
using xlpp::internal::ooxml::tableXml;
using xlpp::internal::ooxml::qualifiedPrintReference;
using xlpp::internal::ooxml::serializeSheets;
using xlpp::internal::ooxml::resolvePackagePart;
using xlpp::internal::ooxml::relationshipKind;
using xlpp::internal::ooxml::relationshipsForSource;

namespace {

// Map public compression enums onto zlib values without exposing zlib.h.
int zlibLevel(xlpp::CompressionLevel level) {
    switch (level) {
        case xlpp::CompressionLevel::Store: return 0;
        case xlpp::CompressionLevel::Fastest: return 1;
        case xlpp::CompressionLevel::Default: return -1;
        case xlpp::CompressionLevel::Best: return 9;
    }
    return -1;
}

int zlibStrategy(xlpp::CompressionStrategy strategy) {
    switch (strategy) {
        case xlpp::CompressionStrategy::Default: return 0;
        case xlpp::CompressionStrategy::Filtered: return 1;
        case xlpp::CompressionStrategy::HuffmanOnly: return 2;
        case xlpp::CompressionStrategy::Rle: return 3;
        case xlpp::CompressionStrategy::Fixed: return 4;
    }
    return 0;
}

std::string contentTypes(std::size_t sheetCount,
                         std::size_t tableCount,
                         std::size_t commentCount,
                         const std::vector<std::size_t>& drawingIds,
                         const std::vector<xlpp::PreservedPart>& preserved,
                         bool strict,
                         bool hasSst,
                         const std::vector<std::size_t>& chartIds,
                         const std::set<std::size_t>& chartExIds,
                         const std::vector<std::size_t>& pivotIds,
                         bool hasCustomProperties,
                         bool macroEnabled) {
    std::ostringstream xml;
    std::set<std::string> emittedDefaults;
    std::set<std::string> emittedOverrides;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\""
        << nsCtPkg(strict) << "\">";

    const auto addDefault = [&](const std::string& extension, const std::string& contentType) {
        if (extension.empty() || contentType.empty() || !emittedDefaults.insert(extension).second) return;
        xml << "<Default Extension=\"" << xmlEscape(extension) << "\" ContentType=\""
            << xmlEscape(contentType) << "\"/>";
    };
    const auto addOverride = [&](const std::string& partName, const std::string& contentType) {
        if (partName.empty() || contentType.empty() || !emittedOverrides.insert(partName).second) return;
        xml << "<Override PartName=\"/" << xmlEscape(partName) << "\" ContentType=\""
            << xmlEscape(contentType) << "\"/>";
    };

    addDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
    addDefault("xml", "application/xml");
    addDefault("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing");
    addDefault("png", "image/png");
    addDefault("jpg", "image/jpeg");
    addDefault("jpeg", "image/jpeg");

    addOverride("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
    addOverride("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
    if (hasCustomProperties)
        addOverride("docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml");
    addOverride("xl/workbook.xml", macroEnabled
        ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
        : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    addOverride("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");

    for (std::size_t index = 1; index <= sheetCount; ++index)
        addOverride("xl/worksheets/sheet" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    for (std::size_t index = 1; index <= tableCount; ++index)
        addOverride("xl/tables/table" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml");
    for (std::size_t index = 1; index <= commentCount; ++index)
        addOverride("xl/comments" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");
    for (const auto index : drawingIds)
        addOverride("xl/drawings/drawing" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.drawing+xml");
    if (hasSst)
        addOverride("xl/sharedStrings.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    for (const auto index : chartIds)
        addOverride("xl/charts/chart" + std::to_string(index) + ".xml",
                    chartExIds.count(index) ? "application/vnd.ms-office.chartex+xml"
                                            : "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    for (const auto index : pivotIds) {
        addOverride("xl/pivotTables/pivotTable" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");
        addOverride("xl/pivotCache/pivotCacheDefinition" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
        addOverride("xl/pivotCache/pivotCacheRecords" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml");
    }
    for (const auto& part : preserved) {
        if (!part.overrideType.empty()) addOverride(part.name, part.overrideType);
        else if (!part.defaultType.empty()) addDefault(part.extension, part.defaultType);
    }
    xml << "</Types>";
    return xml.str();
}

// Package parts that `save` regenerates from the model; everything else is
// preserved verbatim across a load/save round trip.

// Effective content types for preserved parts: <Override> rules win; parts
// without one fall back to the matching <Default> rule for their extension.



std::string rootrels(bool strict, bool hasCustomProperties){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\""<<nsRelsPkg(strict)<<"\"><Relationship Id=\"rId1\" Type=\""<<nsRelsDoc(strict)<<"/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\""<<nsRelsPkg(strict)<<"/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\""<<nsRelsDoc(strict)<<"/extended-properties\" Target=\"docProps/app.xml\"/>";if(hasCustomProperties)x<<"<Relationship Id=\"rId4\" Type=\""<<nsRelsDoc(strict)<<"/custom-properties\" Target=\"docProps/custom.xml\"/>";x<<"</Relationships>";return x.str();}
std::string corePropertiesXml(const xlpp::DocumentProperties& p, bool strict){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\""<<nsCoreProps(strict)<<"\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"; if(!p.title().empty())x<<"<dc:title>"<<xmlEscape(p.title())<<"</dc:title>";if(!p.subject().empty())x<<"<dc:subject>"<<xmlEscape(p.subject())<<"</dc:subject>";if(!p.creator().empty())x<<"<dc:creator>"<<xmlEscape(p.creator())<<"</dc:creator>";if(!p.description().empty())x<<"<dc:description>"<<xmlEscape(p.description())<<"</dc:description>";if(!p.keywords().empty())x<<"<cp:keywords>"<<xmlEscape(p.keywords())<<"</cp:keywords>";if(!p.category().empty())x<<"<cp:category>"<<xmlEscape(p.category())<<"</cp:category>";if(!p.lastModifiedBy().empty())x<<"<cp:lastModifiedBy>"<<xmlEscape(p.lastModifiedBy())<<"</cp:lastModifiedBy>";x<<"</cp:coreProperties>";return x.str();}
std::string appPropertiesXml(bool strict){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\""<<nsExtProps(strict)<<"\" xmlns:vt=\""<<nsVTypes(strict)<<"\"><Application>XL++</Application><AppVersion>1.0</AppVersion></Properties>";return x.str();}
std::string customPropertiesXml(const xlpp::CustomProperties& props){std::ostringstream x;x<<R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">)";for(std::size_t i=0;i<props.items().size();++i){const auto&p=props.items()[i];x<<"<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\""<<(i+2)<<"\" name=\""<<xmlEscape(p.name())<<"\"><vt:"<<xmlEscape(p.type())<<">"<<xmlEscape(p.value())<<"</vt:"<<xmlEscape(p.type())<<"></property>";}x<<"</Properties>";return x.str();}

std::string mergePivotTablePartsBlocks(const std::vector<std::string>& originalBlocks,
                                      const std::vector<std::string>& generatedBlocks,
                                      bool preserveOriginal) {
    if (!preserveOriginal || originalBlocks.empty()) return joinBlocks(generatedBlocks);
    if (generatedBlocks.empty()) return joinBlocks(originalBlocks);

    std::vector<std::string> nodes;
    for (const auto& block : originalBlocks) {
        const auto items = extractTagBlocks(block, "pivotTablePart");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    for (const auto& block : generatedBlocks) {
        const auto items = extractTagBlocks(block, "pivotTablePart");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    if (nodes.empty()) return {};
    std::ostringstream xml;
    xml << "<pivotTableParts count=\"" << nodes.size() << "\">";
    for (const auto& node : nodes) xml << node;
    xml << "</pivotTableParts>";
    return xml.str();
}

std::string rebuildWorksheetTail(std::string generated,
                                 const std::string& original,
                                 bool preserveDrawing,
                                 bool preservePivot,
                                 bool preserveTables,
                                 bool preserveComments) {
    static const std::array<const char*, 10> tailTags{
        "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects",
        "controls", "webPublishItems", "tableParts", "pivotTableParts", "extLst"
    };
    std::unordered_map<std::string, std::vector<std::string>> generatedBlocks;
    std::unordered_map<std::string, std::vector<std::string>> originalBlocks;
    for (const auto* tag : tailTags) {
        generatedBlocks[tag] = extractTagBlocks(generated, tag);
        originalBlocks[tag] = extractTagBlocks(original, tag);
        eraseTagBlocks(generated, tag);
    }

    std::string tail;
    tail += joinBlocks(preserveDrawing && !originalBlocks["drawing"].empty()
        ? originalBlocks["drawing"] : generatedBlocks["drawing"]);
    tail += joinBlocks(preserveComments && !originalBlocks["legacyDrawing"].empty()
        ? originalBlocks["legacyDrawing"] : generatedBlocks["legacyDrawing"]);
    for (const auto* tag : {"legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems"})
        tail += joinBlocks(!originalBlocks[tag].empty() ? originalBlocks[tag] : generatedBlocks[tag]);
    tail += joinBlocks(preserveTables && !originalBlocks["tableParts"].empty()
        ? originalBlocks["tableParts"] : generatedBlocks["tableParts"]);
    tail += mergePivotTablePartsBlocks(originalBlocks["pivotTableParts"], generatedBlocks["pivotTableParts"], preservePivot);
    tail += joinBlocks(!originalBlocks["extLst"].empty() ? originalBlocks["extLst"] : generatedBlocks["extLst"]);
    insertBefore(generated, "</worksheet>", tail);
    return generated;
}

std::string preserveWorkbookNodes(std::string generated,
                                  const std::string& original,
                                  bool preservePivotCaches) {
    if (original.empty()) return generated;
    const auto preserveBeforeWorkbookPr = [&](const char* tag) {
        const auto blocks = extractTagBlocks(original, tag);
        if (!blocks.empty() && extractTagBlocks(generated, tag).empty())
            insertBefore(generated, "<workbookPr", joinBlocks(blocks));
    };
    preserveBeforeWorkbookPr("fileVersion");
    preserveBeforeWorkbookPr("fileSharing");

    if (extractTagBlocks(generated, "bookViews").empty())
        insertBefore(generated, "<sheets>", joinBlocks(extractTagBlocks(original, "bookViews")));

    const auto sheetsEnd = generated.find("</sheets>");
    if (sheetsEnd != std::string::npos) {
        std::string afterSheets;
        for (const auto* tag : {"functionGroups", "externalReferences"})
            if (extractTagBlocks(generated, tag).empty()) afterSheets += joinBlocks(extractTagBlocks(original, tag));
        generated.insert(sheetsEnd + std::string("</sheets>").size(), afterSheets);
    }

    std::string finalNodes;
    for (const auto* tag : {"oleSize", "customWorkbookViews"})
        if (extractTagBlocks(generated, tag).empty()) finalNodes += joinBlocks(extractTagBlocks(original, tag));
    if (preservePivotCaches && extractTagBlocks(generated, "pivotCaches").empty())
        finalNodes += joinBlocks(extractTagBlocks(original, "pivotCaches"));
    for (const auto* tag : {"webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"})
        if (extractTagBlocks(generated, tag).empty()) finalNodes += joinBlocks(extractTagBlocks(original, tag));
    insertBefore(generated, "</workbook>", finalNodes);
    return generated;
}

std::string mergeWorkbookPivotCaches(const std::string& originalWorkbookXml,
                                     const std::string& generatedPivotCachesXml) {
    const auto originalContainers = extractTagBlocks(originalWorkbookXml, "pivotCaches");
    if (originalContainers.empty()) return generatedPivotCachesXml;
    if (generatedPivotCachesXml.empty()) return joinBlocks(originalContainers);

    std::vector<std::string> nodes;
    for (const auto& block : originalContainers) {
        const auto items = extractTagBlocks(block, "pivotCache");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    const auto generatedContainers = extractTagBlocks(generatedPivotCachesXml, "pivotCaches");
    for (const auto& block : generatedContainers) {
        const auto items = extractTagBlocks(block, "pivotCache");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    if (nodes.empty()) return {};
    std::ostringstream xml;
    xml << "<pivotCaches>";
    for (const auto& node : nodes) xml << node;
    xml << "</pivotCaches>";
    return xml.str();
}

std::size_t nextAvailablePivotCacheId(const std::string& workbookXml) {
    std::size_t maximum = 0;
    for (const auto& node : extractTagBlocks(workbookXml, "pivotCache")) {
        const auto value = xlpp::internal::attribute(node, "cacheId");
        if (value.empty() || !std::all_of(value.begin(), value.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
        maximum = std::max(maximum, static_cast<std::size_t>(std::stoull(value)));
    }
    return maximum + 1;
}

std::size_t nextAvailablePartId(const std::vector<xlpp::PreservedPart>& parts,
                                const std::string& prefix,
                                const std::string& suffix) {
    std::size_t maximum = 0;
    for (const auto& part : parts) {
        if (part.name.rfind(prefix, 0) != 0 || part.name.size() <= prefix.size() + suffix.size()) continue;
        if (part.name.compare(part.name.size() - suffix.size(), suffix.size(), suffix) != 0) continue;
        const auto number = part.name.substr(prefix.size(), part.name.size() - prefix.size() - suffix.size());
        if (number.empty() || !std::all_of(number.begin(), number.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
        maximum = std::max(maximum, static_cast<std::size_t>(std::stoull(number)));
    }
    return maximum + 1;
}


std::size_t nextAvailableMediaId(const std::vector<xlpp::PreservedPart>& parts) {
    std::size_t maximum = 0;
    const std::string prefix = "xl/media/image";
    for (const auto& part : parts) {
        if (part.name.rfind(prefix, 0) != 0) continue;
        const auto dot = part.name.find('.', prefix.size());
        if (dot == std::string::npos) continue;
        const auto number = part.name.substr(prefix.size(), dot - prefix.size());
        if (number.empty() || !std::all_of(number.begin(), number.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
        maximum = std::max(maximum, static_cast<std::size_t>(std::stoull(number)));
    }
    return maximum + 1;
}



}namespace xlpp {
void Workbook::saveInPlace(const std::filesystem::path& p, const SaveOptions& options) const {
    if (sheets_.empty()) throw std::runtime_error("Workbook needs at least one worksheet");
    if (options.validateBeforeSave) {
        const auto validation = validate();
        if (!validation.ok()) {
            const auto& issue = validation.issues.front();
            throw std::invalid_argument("Workbook validation failed [" + issue.code + "]: " + issue.message +
                                        (issue.worksheet.empty() ? std::string{} : " (worksheet: " + issue.worksheet + ")"));
        }
    }
    if (!options.encryptionPassword.empty()) {
        // Serialize an ordinary OOXML ZIP first, then wrap the exact package in
        // the ECMA-376 Agile Encryption CFB envelope. Keeping the ZIP writer
        // unchanged also preserves deterministic package serialization.
        const auto temporary = internal::xlppTemporaryPath("encrypt");
        auto plainOptions = options;
        plainOptions.encryptionPassword.clear();
        try {
            saveInPlace(temporary, plainOptions);
            auto package = xlpp::internal::readBinaryFile(temporary);
            std::vector<unsigned char> encrypted;
            if (options.encryptionMode == OfficeEncryptionMode::StandardAesSha1) {
                encrypted = internal::encryptStandardOfficePackage(package, options.encryptionPassword, options.encryptionKeyBits);
            } else if (options.encryptionMode == OfficeEncryptionMode::AgileAes256Sha512) {
                encrypted = internal::encryptAgileOfficePackage(package, options.encryptionPassword, options.encryptionSpinCount);
            } else {
                throw std::invalid_argument("Unsupported encryption mode requested for save");
            }
            xlpp::internal::writeBinaryFile(p, encrypted);
            std::filesystem::remove(temporary);
        } catch (...) {
            std::error_code ignored;
            std::filesystem::remove(temporary, ignored);
            throw;
        }
        return;
    }
    if (options.calculateFormulasBeforeSave || options.synchronizeChartCaches) {
        // Keep save() logically const: prepare a private copy, then serialize it
        // with preparation flags disabled to avoid recursion. Formula calculation
        // intentionally precedes chart-cache synchronization.
        Workbook prepared = *this;
        if (options.calculateFormulasBeforeSave) prepared.calculateFormulas();
        if (options.synchronizeChartCaches) {
            ChartCacheSyncOptions syncOptions;
            syncOptions.changedReferencesOnly = options.synchronizeChangedChartCachesOnly;
            prepared.synchronizeChartCaches(syncOptions);
        }
        auto preparedOptions = options;
        preparedOptions.calculateFormulasBeforeSave = false;
        preparedOptions.synchronizeChartCaches = false;
        prepared.saveInPlace(p, preparedOptions);
        return;
    }
    const bool strict = options.strictNamespace;
    StyleCatalog styleCatalog;
    DxfCatalog dxfCatalog;
    std::size_t tableCount = 0;
    std::size_t commentCount = 0;
    std::unordered_map<std::string, std::size_t> sstIndex;
    std::vector<std::string> sstStrings;
    std::size_t sstOccurrences = 0;

    constexpr auto noSourceSheet = std::numeric_limits<std::size_t>::max();
    std::vector<std::size_t> sourceSheetIndices(sheets_.size(), noSourceSheet);
    std::vector<bool> preserveDrawing(sheets_.size(), false);
    std::vector<bool> preservePivot(sheets_.size(), false);
    // Snapshot pivot regeneration boundaries before serializeSheets()/clearDirty().
    // Imported semantic pivots opt into regeneration through pivotsDirty(); clearing
    // dirty flags must not change the package-generation plan halfway through save.
    std::vector<std::size_t> generatedPivotStarts(sheets_.size(), 0);
    for (std::size_t i = 0; i < sheets_.size(); ++i)
        generatedPivotStarts[i] = sheets_[i].generatedPivotStart();
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
        const auto sourceName = std::find(sourceSheetNames_.begin(), sourceSheetNames_.end(), sheets_[i].name());
        if (sourceName == sourceSheetNames_.end()) continue;
        const auto sourceIndex = static_cast<std::size_t>(std::distance(sourceSheetNames_.begin(), sourceName));
        if (sourceIndex >= sourceSheetXml_.size() || sourceIndex >= sourceSheetParts_.size()) continue;
        sourceSheetIndices[i] = sourceIndex;
        preserveDrawing[i] = !sheets_[i].drawingsDirty()
            && !internal::tags(sourceSheetXml_[sourceIndex], "drawing").empty();
        const auto sourceRelationships = relationshipsForSource(preservedRelationships_, sourceSheetParts_[sourceIndex]);
        const bool hasImportedPivot = !internal::tags(sourceSheetXml_[sourceIndex], "pivotTableParts").empty()
            || std::any_of(sourceRelationships.begin(), sourceRelationships.end(), [](const auto& relationship) {
                return relationshipKind(relationship) == "pivotTable";
            });
        preservePivot[i] = hasImportedPivot && !sheets_[i].pivotsDirty();
    }

    const auto firstDrawingId = nextAvailablePartId(preservedParts_, "xl/drawings/drawing", ".xml");
    const auto firstChartId = nextAvailablePartId(preservedParts_, "xl/charts/chart", ".xml");
    const auto firstPivotId = nextAvailablePartId(preservedParts_, "xl/pivotTables/pivotTable", ".xml");
    std::vector<std::size_t> generatedDrawingIds;
    std::vector<std::size_t> generatedChartIds;
    std::set<std::size_t> generatedChartExIds;
    std::vector<std::size_t> generatedPivotIds;
    std::unordered_map<std::size_t, std::size_t> generatedPivotCacheIds;
    std::size_t nextDrawingId = firstDrawingId;
    std::size_t nextChartId = firstChartId;
    std::size_t nextPivotId = firstPivotId;
    std::size_t nextPivotCacheId = nextAvailablePivotCacheId(sourceWorkbookXml_);

    for (std::size_t sheetIndex = 0; sheetIndex < sheets_.size(); ++sheetIndex) {
        const auto& sheet = sheets_[sheetIndex];
        tableCount += sheet.tables().size();
        bool hasComments = false;
        for (const auto& pair : sheet.cells()) if (pair.second.hasComment()) { hasComments = true; break; }
        if (hasComments) ++commentCount;
        if (!preserveDrawing[sheetIndex] && (!sheet.images().empty() || sheet.chartCount() > 0)) {
            generatedDrawingIds.push_back(nextDrawingId++);
            for (std::size_t chartIndex = 0; chartIndex < sheet.chartCount(); ++chartIndex) {
                const auto id = nextChartId++;
                generatedChartIds.push_back(id);
                if (sheet.chart(chartIndex).modern()) generatedChartExIds.insert(id);
            }
        } else if (preserveDrawing[sheetIndex] && sheet.appendedChartCount() > 0) {
            const auto firstAppended = sheet.loadedChartCount();
            for (std::size_t chartIndex = 0; chartIndex < sheet.appendedChartCount(); ++chartIndex) {
                const auto id = nextChartId++;
                generatedChartIds.push_back(id);
                if (sheet.chart(firstAppended + chartIndex).modern()) generatedChartExIds.insert(id);
            }
        }
        for (std::size_t pivotIndex = generatedPivotStarts[sheetIndex]; pivotIndex < sheet.pivotTables().size(); ++pivotIndex) {
            const auto pivotPartId = nextPivotId++;
            generatedPivotIds.push_back(pivotPartId);
            generatedPivotCacheIds.emplace(pivotPartId, nextPivotCacheId++);
        }
        for (const auto& entry : sheet.cells()) {
            if (!entry.second.style().isDefault())
                styleCatalog.id(entry.second.style());
            if (const auto* sv = std::get_if<std::string>(&entry.second.value())) {
                ++sstOccurrences;
                const auto [it, inserted] = sstIndex.try_emplace(*sv, sstStrings.size());
                if (inserted) sstStrings.push_back(*sv);
            }
        }
        for (const auto& formatting : sheet.conditionalFormatting().entries())
            for (const auto& rule : formatting.rules()) if (rule.hasDifferentialStyle()) dxfCatalog.id(rule.differentialStyle());
    }
    for (const auto& named : namedStyles_) styleCatalog.id(named.style());
    const bool macroEnabled = hasVbaProject();
    auto sheetXmls = serializeSheets(sheets_, styleCatalog, dxfCatalog, date1904_, strict, macroEnabled,
                                           options.parallelSheets ? options.parallelWorkers : 0,
                                           options.parallelRows, &sstIndex, &cachedSheetXml_,
                                           cachedSheetXmlStrict_, cachedSheetXmlDate1904_);
    for (auto& sheet : sheets_) sheet.clearDirty();
    internal::ZipArchive z;
    std::set<std::string> suppressedPreservedParts;
    z.setCompressionLevel(zlibLevel(options.compressionLevel));
    z.setCompressionStrategy(zlibStrategy(options.compressionStrategy));
    z.setParallelWorkers(options.parallelWorkers);
    z.add("[Content_Types].xml", contentTypes(sheets_.size(), tableCount, commentCount,
        generatedDrawingIds, preservedParts_, strict, !sstStrings.empty(), generatedChartIds,
        generatedChartExIds, generatedPivotIds, (!customProps_.empty() || customPropertiesPartPresent_), macroEnabled));
    const auto rootOriginalRelationships = relationshipsForSource(preservedRelationships_, {});
    const auto mergedRootRelationships = mergeRelationshipsXml(rootrels(strict, (!customProps_.empty() || customPropertiesPartPresent_)), rootOriginalRelationships,
        [](const PreservedRelationship& relationship) {
            const auto kind = relationshipKind(relationship);
            return kind != "officeDocument" && kind != "core-properties"
                && kind != "extended-properties" && kind != "custom-properties";
        }, strict, nullptr);
    z.add("_rels/.rels", mergedRootRelationships);
    z.add("docProps/core.xml", corePropertiesXml(properties_, strict));
    z.add("docProps/app.xml", appPropertiesXml(strict));
    if (!customProps_.empty() || customPropertiesPartPresent_) z.add("docProps/custom.xml", customPropertiesXml(customProps_));
    z.add("xl/styles.xml", stylesXml(styleCatalog, namedStyles_, dxfCatalog, strict));
    if (!sstStrings.empty()) {
        std::ostringstream sstXml;
        sstXml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"" << nsMain(strict)
               << "\" count=\"" << sstOccurrences << "\" uniqueCount=\"" << sstStrings.size() << "\">";
        for (const auto& text : sstStrings) {
            sstXml << "<si><t xml:space=\"preserve\">"; writeXmlEscaped(sstXml, text); sstXml << "</t></si>";
        }
        sstXml << "</sst>";
        z.add("xl/sharedStrings.xml", sstXml.str());
    }
    std::ostringstream wb, rels;
    wb << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"" << nsMain(strict) << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\">";
    // CT_Workbook requires workbookPr before workbookProtection and sheets.
    wb << "<workbookPr date1904=\"" << (date1904_ ? 1 : 0) << "\"";
    if (macroEnabled) wb << " codeName=\"ThisWorkbook\"";
    wb << "/>";
    const auto& wp = protection_;
    if (wp.lockStructure() || wp.lockWindows() || wp.lockRevision() || !wp.workbookPasswordHash().empty()) {
        wb << "<workbookProtection lockStructure=\"" << (wp.lockStructure()?1:0) << "\" lockWindows=\"" << (wp.lockWindows()?1:0) << "\" lockRevision=\"" << (wp.lockRevision()?1:0) << "\"";
        if (!wp.workbookPasswordHash().empty()) wb << " workbookPassword=\"" << xmlEscape(wp.workbookPasswordHash()) << "\"";
        wb << "/>";
    }
    wb << "<sheets>";
    rels << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << nsRelsPkg(strict) << "\">";
    std::size_t globalTableId = 1;
    std::size_t globalCommentId = 1;
    std::size_t globalDrawingId = firstDrawingId;
    std::size_t globalMediaId = nextAvailableMediaId(preservedParts_);
    std::size_t globalChartId = firstChartId;
    std::size_t globalPivotId = firstPivotId;
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
        const auto& sheet = sheets_[i];
        wb << "<sheet name=\"" << xmlEscape(sheet.name()) << "\" sheetId=\"" << i+1 << "\" r:id=\"rId" << i+1 << "\"/>";
        rels << "<Relationship Id=\"rId" << i+1 << "\" Type=\"" << nsRelsDoc(strict) << "/worksheet\" Target=\"worksheets/sheet" << i+1 << ".xml\"/>";
        bool hasLinks=false; for(const auto& pair:sheet.cells()) if(pair.second.hasHyperlink() && pair.second.hyperlinkValue()->external()) hasLinks=true;
        bool hasComments=false; for(const auto& pair:sheet.cells()) if(pair.second.hasComment()) { hasComments=true; break; }
        const bool hasGeneratedImages = !preserveDrawing[i] && !sheet.images().empty();
        const bool hasSheetCharts = !preserveDrawing[i] && sheet.chartCount() > 0;
        const bool hasSheetPivots = generatedPivotStarts[i] < sheet.pivotTables().size();
        const auto sourceSheetIndex = sourceSheetIndices[i];
        const std::string originalSheetPart = sourceSheetIndex != noSourceSheet
            ? sourceSheetParts_[sourceSheetIndex] : std::string{};
        const std::string originalSheetXml = sourceSheetIndex != noSourceSheet
            ? sourceSheetXml_[sourceSheetIndex] : std::string{};
        const auto originalSheetRelationships = relationshipsForSource(preservedRelationships_, originalSheetPart);
        if (!preservePivot[i] && !originalSheetPart.empty()) {
            for (const auto& relationship : originalSheetRelationships) {
                if (relationshipKind(relationship) != "pivotTable" || relationship.targetMode == "External") continue;
                const auto pivotPart = resolvePackagePart(originalSheetPart, relationship.target);
                if (!pivotPart.empty()) xlpp::internal::preservation::suppressExclusivePartClosure(pivotPart, preservedRelationships_, suppressedPreservedParts);
            }
        }
        const bool hasOriginalTableOwner = !extractTagBlocks(originalSheetXml, "tableParts").empty();
        const bool hasOriginalCommentVmlOwner = !extractTagBlocks(originalSheetXml, "legacyDrawing").empty();
        const bool hasOriginalTableRelationship = std::any_of(originalSheetRelationships.begin(), originalSheetRelationships.end(), [](const auto& relationship) {
            return relationshipKind(relationship) == "table";
        });
        const bool hasOriginalCommentsRelationship = std::any_of(originalSheetRelationships.begin(), originalSheetRelationships.end(), [](const auto& relationship) {
            return relationshipKind(relationship) == "comments";
        });
        const bool preserveTables = sheet.tables().empty() && hasOriginalTableOwner && hasOriginalTableRelationship;
        const bool preserveComments = !hasComments && hasOriginalCommentsRelationship;
        const bool hasPreservedSheetRelationships = std::any_of(originalSheetRelationships.begin(), originalSheetRelationships.end(), [&](const auto& relationship) {
            const auto kind = relationshipKind(relationship);
            if (kind == "drawing") return static_cast<bool>(preserveDrawing[i]);
            if (kind == "pivotTable") return static_cast<bool>(preservePivot[i]);
            if (kind == "table") return preserveTables;
            if (kind == "comments") return preserveComments;
            if (kind == "vmlDrawing") return preserveComments && hasOriginalCommentVmlOwner;
            return kind != "hyperlink";
        });
        if (!sheet.tables().empty() || hasLinks || hasComments || hasGeneratedImages || hasSheetCharts || hasSheetPivots || hasPreservedSheetRelationships) {
            std::ostringstream sheetRels;
            sheetRels << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << nsRelsPkg(strict) << "\">";
            for (std::size_t t = 0; t < sheet.tables().size(); ++t, ++globalTableId) {
                sheetRels << "<Relationship Id=\"rId" << t+1 << "\" Type=\"" << nsRelsDoc(strict) << "/table\" Target=\"../tables/table" << globalTableId << ".xml\"/>";
                z.add("xl/tables/table"+std::to_string(globalTableId)+".xml", tableXml(sheet.tables()[t], sheet, globalTableId, strict));
            }
            std::size_t hid=1; for(const auto& pair:sheet.cells()) if(pair.second.hasHyperlink() && pair.second.hyperlinkValue()->external()) { const auto& h=*pair.second.hyperlinkValue(); sheetRels<<"<Relationship Id=\"rIdHyperlink"<<hid++<<"\" Type=\""<<nsRelsDoc(strict)<<"/hyperlink\" Target=\""<<xmlEscape(h.target())<<"\" TargetMode=\"External\"/>"; }
            if (hasComments) {
                const auto generatedCommentsPart = "xl/comments" + std::to_string(globalCommentId) + ".xml";
                const auto generatedVmlPart = "xl/drawings/commentsDrawing" + std::to_string(globalCommentId) + ".vml";
                for (const auto& relationship : originalSheetRelationships) {
                    const auto kind = relationshipKind(relationship);
                    if (kind != "comments" && kind != "vmlDrawing") continue;
                    const auto originalPart = internal::RelationshipGraph::resolveTarget(originalSheetPart, relationship.target);
                    const auto replacementPart = kind == "comments" ? generatedCommentsPart : generatedVmlPart;
                    if (!originalPart.empty() && originalPart != replacementPart)
                        suppressedPreservedParts.insert(originalPart);
                }
                sheetRels << "<Relationship Id=\"rIdComments\" Type=\"" << nsRelsDoc(strict) << "/comments\" Target=\"../comments" << globalCommentId << ".xml\"/>";
                sheetRels << "<Relationship Id=\"rIdCommentsVml\" Type=\"" << nsRelsDoc(strict) << "/vmlDrawing\" Target=\"../drawings/commentsDrawing" << globalCommentId << ".vml\"/>";
                z.add("xl/comments" + std::to_string(globalCommentId) + ".xml", commentsXml(sheet, strict));
                z.add("xl/drawings/commentsDrawing" + std::to_string(globalCommentId) + ".vml", commentsVml(sheet));
                ++globalCommentId;
            }
            if (hasGeneratedImages || hasSheetCharts) {
                sheetRels << "<Relationship Id=\"rIdDrawing\" Type=\"" << nsRelsDoc(strict) << "/drawing\" Target=\"../drawings/drawing" << globalDrawingId << ".xml\"/>";
                const auto drawingFirstChartId = globalChartId;
                for (std::size_t ci = 0; ci < sheet.chartCount(); ++ci, ++globalChartId)
                    z.add("xl/charts/chart" + std::to_string(globalChartId) + ".xml", xlpp::internal::ooxml::serializeChart(sheet.chart(ci), strict));
                z.add("xl/drawings/drawing" + std::to_string(globalDrawingId) + ".xml", drawingXml(sheet, strict));
                z.add("xl/drawings/_rels/drawing" + std::to_string(globalDrawingId) + ".xml.rels", drawingRelationshipsXml(sheet, globalMediaId, drawingFirstChartId, strict));
                for (const auto& image : sheet.images()) {
                    const std::string bytes(reinterpret_cast<const char*>(image.bytes().data()), image.bytes().size());
                    z.add("xl/media/image" + std::to_string(globalMediaId++) + "." + image.extension(), bytes, false);
                }
                ++globalDrawingId;
            }
            if (preserveDrawing[i] && (sheet.appendedImageCount() > 0 ||
                !internal::WorkbookDrawingAccess::imageEdits(sheet).empty())) {
                if (!xlpp::internal::ooxml::applyImageChangesToPreservedDrawing(z, sheet, originalSheetPart, originalSheetXml, preservedRelationships_, preservedParts_,
                                                         globalMediaId, suppressedPreservedParts))
                    throw std::runtime_error("Cannot apply selective image mutation to preserved drawing for worksheet: " + sheet.name());
            }
            if (preserveDrawing[i] && (sheet.appendedChartCount() > 0 ||
                !internal::WorkbookDrawingAccess::chartEdits(sheet).empty())) {
                if (!xlpp::internal::ooxml::applyChartChangesToPreservedDrawing(z, sheet, originalSheetPart, originalSheetXml,
                                                         preservedRelationships_, preservedParts_, globalChartId,
                                                         suppressedPreservedParts))
                    throw std::runtime_error("Cannot apply selective chart mutation to preserved drawing for worksheet: " + sheet.name());
            }
            if (hasSheetPivots) {
                for (std::size_t pi = generatedPivotStarts[i]; pi < sheet.pivotTables().size(); ++pi, ++globalPivotId) {
                    sheetRels << "<Relationship Id=\"rIdPivot" << (pi + 1) << "\" Type=\"" << nsRelsDoc(strict) << "/pivotTable\" Target=\"../pivotTables/pivotTable" << globalPivotId << ".xml\"/>";

                    const auto cacheId = generatedPivotCacheIds.at(globalPivotId);
                    const auto effectivePivot = effectivePivotTable(sheet.pivotTables()[pi], sheets_, sheet, cacheId);
                    z.add("xl/pivotTables/pivotTable" + std::to_string(globalPivotId) + ".xml", pivotTableXml(effectivePivot, cacheId, strict));
                    z.add("xl/pivotTables/_rels/pivotTable" + std::to_string(globalPivotId) + ".xml.rels",
                          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + nsRelsPkg(strict) + "\"><Relationship Id=\"rId1\" Type=\"" + nsRelsDoc(strict) + "/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition" + std::to_string(globalPivotId) + ".xml\"/></Relationships>");
                    z.add("xl/pivotCache/pivotCacheDefinition" + std::to_string(globalPivotId) + ".xml", pivotCacheXml(effectivePivot, strict));
                    z.add("xl/pivotCache/pivotCacheRecords" + std::to_string(globalPivotId) + ".xml", pivotCacheRecordsXml(effectivePivot, strict));
                    z.add("xl/pivotCache/_rels/pivotCacheDefinition" + std::to_string(globalPivotId) + ".xml.rels",
                          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + nsRelsPkg(strict) + "\"><Relationship Id=\"rId1\" Type=\"" + nsRelsDoc(strict) + "/pivotCacheRecords\" Target=\"pivotCacheRecords" + std::to_string(globalPivotId) + ".xml\"/></Relationships>");
                }
            }
            sheetRels << "</Relationships>";
            auto generatedSheetXml = sheetXmls[i];
            const auto mergedSheetRelationships = mergeRelationshipsXml(sheetRels.str(), originalSheetRelationships,
                [&](const PreservedRelationship& relationship) {
                    const auto kind = relationshipKind(relationship);
                    if (kind == "drawing") return static_cast<bool>(preserveDrawing[i]);
                    if (kind == "pivotTable") return static_cast<bool>(preservePivot[i]);
                    if (kind == "table") return preserveTables;
                    if (kind == "comments") return preserveComments;
                    if (kind == "vmlDrawing") return preserveComments && hasOriginalCommentVmlOwner;
                    return kind != "hyperlink";
                }, strict, &generatedSheetXml);
            generatedSheetXml = rebuildWorksheetTail(std::move(generatedSheetXml), originalSheetXml,
                                                      preserveDrawing[i], preservePivot[i], preserveTables, preserveComments);
            z.add("xl/worksheets/sheet"+std::to_string(i+1)+".xml", generatedSheetXml);
            z.add("xl/worksheets/_rels/sheet"+std::to_string(i+1)+".xml.rels", mergedSheetRelationships);
        } else {
            auto generatedSheetXml = rebuildWorksheetTail(sheetXmls[i], originalSheetXml,
                                                          preserveDrawing[i], preservePivot[i], preserveTables, preserveComments);
            z.add("xl/worksheets/sheet"+std::to_string(i+1)+".xml", generatedSheetXml);
        }
    }
    wb << "</sheets>";
    std::size_t nextWorkbookRelId = sheets_.size() + 1;
    std::ostringstream generatedPivotCaches;
    if (!generatedPivotIds.empty()) {
        generatedPivotCaches << "<pivotCaches>";
        for (const auto pivotPartId : generatedPivotIds) {
            const auto cacheId = generatedPivotCacheIds.at(pivotPartId);
            generatedPivotCaches << "<pivotCache cacheId=\"" << cacheId << "\" r:id=\"rId" << nextWorkbookRelId << "\"/>";
            rels << "<Relationship Id=\"rId" << nextWorkbookRelId++ << "\" Type=\"" << nsRelsDoc(strict) << "/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition" << pivotPartId << ".xml\"/>";
        }
        generatedPivotCaches << "</pivotCaches>";
    }
    if (!sstStrings.empty())
        rels << "<Relationship Id=\"rId" << nextWorkbookRelId++ << "\" Type=\"" << nsRelsDoc(strict) << "/sharedStrings\" Target=\"sharedStrings.xml\"/>";
    rels << "<Relationship Id=\"rId" << nextWorkbookRelId++ << "\" Type=\"" << nsRelsDoc(strict) << "/styles\" Target=\"styles.xml\"/>";
    if (macroEnabled)
        rels << "<Relationship Id=\"rId" << nextWorkbookRelId++ << "\" Type=\"" << nsRelsDoc(strict) << "/vbaProject\" Target=\"vbaProject.bin\"/>";
    bool hasPrintDefinedNames = false;
    for (const auto& sheet : sheets_) {
        if (!sheet.printArea().empty() || !sheet.printTitlesRows().empty() || !sheet.printTitlesCols().empty()) {
            hasPrintDefinedNames = true;
            break;
        }
    }
    if (!definedNames_.empty() || hasPrintDefinedNames) {
        wb << "<definedNames>";
        for (const auto& item : definedNames_) {
            if (item.name() == "_xlnm.Print_Area" || item.name() == "_xlnm.Print_Titles") continue;
            wb << "<definedName name=\"" << xmlEscape(item.name()) << "\"";
            if (item.localSheetId()) wb << " localSheetId=\"" << *item.localSheetId() << "\"";
            if (item.hidden()) wb << " hidden=\"1\"";
            if (!item.comment().empty()) wb << " comment=\"" << xmlEscape(item.comment()) << "\"";
            wb << ">" << xmlEscape(item.value()) << "</definedName>";
        }
        for (std::size_t sheetIndex = 0; sheetIndex < sheets_.size(); ++sheetIndex) {
            const auto& sheet = sheets_[sheetIndex];
            if (!sheet.printArea().empty()) {
                wb << "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"" << sheetIndex << "\">"
                   << xmlEscape(qualifiedPrintReference(sheet.name(), sheet.printArea())) << "</definedName>";
            }
            if (!sheet.printTitlesRows().empty() || !sheet.printTitlesCols().empty()) {
                std::string titles;
                if (!sheet.printTitlesCols().empty()) titles = qualifiedPrintReference(sheet.name(), sheet.printTitlesCols());
                if (!sheet.printTitlesRows().empty()) {
                    if (!titles.empty()) titles += ',';
                    titles += qualifiedPrintReference(sheet.name(), sheet.printTitlesRows());
                }
                wb << "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"" << sheetIndex << "\">"
                   << xmlEscape(titles) << "</definedName>";
            }
        }
        wb << "</definedNames>";
    }
    {
        const auto& cp = calcProps_;
        if (cp.calcMode() != "auto" || cp.calcId() != 191029 || !cp.fullPrecision() || cp.iterate()
            || cp.fullCalcOnLoad() || cp.calcOnSave()) {
            wb << "<calcPr calcId=\"" << cp.calcId() << "\" calcMode=\"" << xmlEscape(cp.calcMode()) << "\""
               << " fullPrecision=\"" << (cp.fullPrecision() ? 1 : 0) << "\""
               << " iterate=\"" << (cp.iterate() ? 1 : 0) << "\""
               << " iterateCount=\"" << cp.iterateCount() << "\" iterateDelta=\"" << cp.iterateDelta() << "\"";
            if (cp.fullCalcOnLoad()) wb << " fullCalcOnLoad=\"1\"";
            if (cp.calcOnSave()) wb << " calcOnSave=\"1\"";
            wb << "/>";
        }
    }
    wb << generatedPivotCaches.str();
    rels << "</Relationships>";
    wb << "</workbook>";
    const bool preserveWorkbookPivotCaches = !internal::tags(sourceWorkbookXml_, "pivotCaches").empty();
    auto workbookXml = wb.str();
    const auto workbookOriginalRelationships = relationshipsForSource(preservedRelationships_, "xl/workbook.xml");
    const auto mergedWorkbookRelationships = mergeRelationshipsXml(rels.str(), workbookOriginalRelationships,
        [&](const PreservedRelationship& relationship) {
            const auto kind = relationshipKind(relationship);
            if (kind == "pivotCacheDefinition") return preserveWorkbookPivotCaches;
            return kind != "worksheet" && kind != "styles" && kind != "sharedStrings" && kind != "vbaProject";
        }, strict, &workbookXml);
    const auto rewrittenGeneratedPivotCaches = joinBlocks(extractTagBlocks(workbookXml, "pivotCaches"));
    eraseTagBlocks(workbookXml, "pivotCaches");
    workbookXml = preserveWorkbookNodes(std::move(workbookXml), sourceWorkbookXml_, false);
    insertBefore(workbookXml, "</workbook>", mergeWorkbookPivotCaches(sourceWorkbookXml_, rewrittenGeneratedPivotCaches));
    z.add("xl/workbook.xml", workbookXml);
    z.add("xl/_rels/workbook.xml.rels", mergedWorkbookRelationships);

    // A source-generated VBA project contains one document module per worksheet.
    // Rebuild it at save time so worksheets added or removed after
    // setVbaModuleText() cannot leave vbaProject.bin out of sync with sheetPr
    // codeName attributes. Externally attached projects remain byte-for-byte
    // preserved because their document-module mapping is owned by the caller.
    std::string generatedVbaData;
    if (generatedVbaProject_) {
        const auto modules = vbaModules();
        std::vector<std::string> codeNames;
        codeNames.reserve(sheets_.size());
        for (std::size_t i = 0; i < sheets_.size(); ++i)
            codeNames.push_back(sheets_[i].vbaCodeName().empty() ? "Sheet" + std::to_string(i + 1) : sheets_[i].vbaCodeName());
        const auto bytes = internal::buildVbaProjectBinary(modules, codeNames, vbaProjectProperties());
        generatedVbaData.assign(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    }
    for (const auto& part : preservedParts_) {
        if (suppressedPreservedParts.contains(part.name)) continue;
        if (generatedVbaProject_ && part.name == "xl/vbaProject.bin")
            z.addUnique(part.name, generatedVbaData, false);
        else
            z.addUnique(part.name, part.data, part.compress);
    }
    if (!suppressedPreservedParts.empty() && z.contains("[Content_Types].xml")) {
        auto contentTypesXml = z.get("[Content_Types].xml");
        for (const auto& overrideNode : internal::tags(contentTypesXml, "Override")) {
            auto partName = internal::attribute(overrideNode, "PartName");
            if (!partName.empty() && partName.front() == '/') partName.erase(partName.begin());
            if (!suppressedPreservedParts.contains(partName)) continue;
            const auto position = contentTypesXml.find(overrideNode);
            if (position != std::string::npos) contentTypesXml.erase(position, overrideNode.size());
        }
        z.replace("[Content_Types].xml", std::move(contentTypesXml));
    }
    z.save(p);
}
}
