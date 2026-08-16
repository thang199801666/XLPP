#include <unordered_set>
#include <cstring>
#include <fstream>
#include "WorkbookPackageReader.h"
#include "WorkbookChartsheetIO.h"
#include "WorkbookChartsheetPackage.h"
#include "WorkbookChartSerializer.h"
#include "WorkbookNamespaces.h"
#include "WorkbookStylesIO.h"
#include "WorkbookDrawingIO.h"
#include "WorkbookChartReader.h"
#include "WorkbookPivotWrite.h"
#include "WorkbookPartXml.h"
#include "WorkbookDrawingEdit.h"
#include "WorkbookChartEditing.h"
#include "WorkbookSheetWriter.h"
#include "WorkbookPivotRead.h"
#include "WorkbookSlicerIO.h"
#include "../Legacy/XlsBinaryReader.h"
#include "../Legacy/XlsBinaryWriter.h"
#include "../Legacy/XlsbBinaryReader.h"
#include "../Legacy/XlsbBinaryWriter.h"
#include <XLPP/Workbook/Workbook.h>
#include "../XML/XmlUtilities.h"
#include "../XML/NumericParsing.h"
#include "../Packaging/ZipArchive.h"
#include "../Packaging/RelationshipGraph.h"
#include "../Threading/ThreadPool.h"
#include "../Vba/VbaProjectBinary.h"
#include "../Encryption/OfficeCrypto.h"
#include "../Internal/WorksheetName.h"
#include <algorithm>
#include <array>
#include <cctype>
#include <cmath>
#include <cstdlib>
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
#include <regex>
#include <string_view>

using xlpp::internal::xmlEscape;
using xlpp::internal::writeXmlEscaped;
using xlpp::internal::chartSeriesCacheXml;
using xlpp::internal::chartView3DXml;
using xlpp::internal::drawingTags;
using xlpp::internal::drawingTagText;
using xlpp::internal::generatedChartTypeUsesXYAxes;
using xlpp::internal::generatedChartTypeHasAxes;
using xlpp::internal::nsMain;
using xlpp::internal::nsRelsPkg;
using xlpp::internal::nsRelsDoc;
using xlpp::internal::nsCtPkg;
using xlpp::internal::nsCoreProps;
using xlpp::internal::nsExtProps;
using xlpp::internal::nsVTypes;
using xlpp::internal::StyleCatalog;
using xlpp::internal::DxfCatalog;
using xlpp::internal::stylesXml;
using xlpp::internal::parseStyleCatalog;
using xlpp::internal::parseDifferentialStyles;
using xlpp::internal::commentsXml;
using xlpp::internal::commentsVml;
using xlpp::internal::drawingXml;
using xlpp::internal::drawingRelationshipsXml;
using xlpp::internal::tableXml;
using xlpp::internal::resolvePackagePart;
using xlpp::internal::partExtension;
using xlpp::internal::loadImages;
using xlpp::internal::parseDrawingMarker;
using xlpp::internal::parseChartType;
using xlpp::internal::parseChartGrouping;
using xlpp::internal::parseChartAxes;
using xlpp::internal::axisTitleById;
using xlpp::internal::parseChartPlots;
using xlpp::internal::parseChartDataLabels;
using xlpp::internal::chartTitleText;
using xlpp::internal::axisTitleText;
using xlpp::internal::chartBoolValue;
using xlpp::internal::seriesDirectSpPr;
using xlpp::internal::axisDirectSpPr;
using xlpp::internal::parseChartSeriesCache;
using xlpp::internal::parseChartThemePalette;
using xlpp::internal::parseChartStyleResources;
using xlpp::internal::parseChartRichText;
using xlpp::internal::parseChartTextStyle;
using xlpp::internal::parseChartDataPoints;
using xlpp::internal::parseChartTrendlines;
using xlpp::internal::parseChartErrorBars;
using xlpp::internal::parseChartLineFormat;
using xlpp::internal::parseChartFillFormat;
using xlpp::internal::parseChartMarkerFormat;
using xlpp::internal::parseChartDataTable;
using xlpp::internal::parseChartUpDownBars;
using xlpp::internal::parseChartManualLayout;
using xlpp::internal::parseChartWallFormat;
using xlpp::internal::parseChartView3D;
using xlpp::internal::parseChartAnchorInfo;
using xlpp::internal::PivotValueKind;
using xlpp::internal::PivotSharedItem;
using xlpp::internal::resolvedPivotFieldIndex;
using xlpp::internal::pivotValueKind;
using xlpp::internal::publicPivotValueKind;
using xlpp::internal::pivotSharedItems;
using xlpp::internal::pivotTableXml;
using xlpp::internal::pivotFieldIsPureData;
using xlpp::internal::writePivotValue;
using xlpp::internal::pivotCacheXml;
using xlpp::internal::pivotCacheRecordsXml;
using xlpp::internal::quotePivotSheetName;
using xlpp::internal::pivotCellText;
using xlpp::internal::pivotCachesEquivalent;
using xlpp::internal::effectivePivotTable;
using xlpp::internal::relationshipKind;
using xlpp::internal::relationshipsForSource;
using xlpp::internal::sameRelationship;
using xlpp::internal::allocateRelationshipId;
using xlpp::internal::replaceRelationshipReference;
using xlpp::internal::mergeRelationshipsXml;
using xlpp::internal::extractTagBlocks;
using xlpp::internal::eraseTagBlocks;
using xlpp::internal::joinBlocks;
using xlpp::internal::insertBefore;
using xlpp::internal::findPreservedPart;
using xlpp::internal::maximumDrawingObjectId;
using xlpp::internal::appendedImageAnchorXml;
using xlpp::internal::appendedChartAnchorXml;
using xlpp::internal::replaceSimpleDrawingText;
using xlpp::internal::replaceAttributeInNode;
using xlpp::internal::patchFirstDrawingNodeAttributes;
using xlpp::internal::patchDrawingMarker;
using xlpp::internal::drawingObjectIdFromStableId;
using xlpp::internal::anchorMatchesStableId;
using xlpp::internal::anchorReferencesRelationship;
using xlpp::internal::drawingReferencesRelationship;
using xlpp::internal::patchImportedImageAnchor;
using xlpp::internal::chartAnchorMatchesStableId;
using xlpp::internal::chartAnchorReferencesRelationship;
using xlpp::internal::drawingReferencesChartRelationship;
using xlpp::internal::removeImportedChartAnchor;
using xlpp::internal::patchImportedChartAnchor;
using xlpp::internal::replaceSimpleElementText;
using xlpp::internal::eraseChartCacheBlocks;
using xlpp::internal::patchSeriesReferenceContainer;
using xlpp::internal::chartXml;
using xlpp::internal::combinedChartXml;
using xlpp::internal::generatedChartSeriesXml;
using xlpp::internal::chartSpaceDirectSpPr;
using xlpp::internal::plotAreaDirectSpPr;
using xlpp::internal::patchNestedLineFormat;
using xlpp::internal::generatedChartWallXml;
using xlpp::internal::generatedPlotAuxiliaryXml;
using xlpp::internal::generatedDataTableXml;
using xlpp::internal::patchMarkerFormatInOwner;
using xlpp::internal::nextAvailablePartId;
using xlpp::internal::suppressExclusivePartClosure;
using xlpp::internal::applyChartChangesToPreservedDrawing;
using xlpp::internal::applyImageChangesToPreservedDrawing;
using xlpp::internal::rebuildWorksheetTail;
using xlpp::internal::workbookViewsXml;
using xlpp::internal::preserveWorkbookNodes;
using xlpp::internal::mergeWorkbookPivotCaches;
using xlpp::internal::nextAvailablePivotCacheId;
using xlpp::internal::nextAvailableMediaId;
using xlpp::internal::sheetXml;
using xlpp::internal::parseRichTextRuns;
using xlpp::internal::slicerCacheXml;
using xlpp::internal::slicerXml;
using xlpp::internal::slicerCacheRelationshipId;
using xlpp::internal::slicerRelationshipId;
using xlpp::internal::insertSlicerListExt;
using xlpp::internal::insertWorkbookSlicerCachesExt;
using xlpp::internal::LoadedSharedString;
using xlpp::internal::loadPivotTables;
using xlpp::internal::filterOperatorName;
using xlpp::internal::parseFilterOperator;
using xlpp::internal::conditionalOperatorName;
using xlpp::internal::parseConditionalOperator;
using xlpp::internal::writeCfvo;
using xlpp::internal::parseCfvo;
using xlpp::internal::dataValidationTypeName;
using xlpp::internal::parseDataValidationType;
using xlpp::internal::dataValidationOperatorName;
using xlpp::internal::parseDataValidationOperator;
using xlpp::internal::dataValidationErrorStyleName;
using xlpp::internal::parseDataValidationErrorStyle;
using xlpp::internal::SstIndex;


#include "WorkbookDrawingAccess.h"

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
// Namespace URI pairs for transitional (default) and strict OOXML (ISO 29500
// strict). Strict uses the purl.oclc.org base URIs; transitional uses the
// schemas.openxmlformats.org URIs.
std::string quoteSheetName(const std::string& name) {
    std::string escaped;
    escaped.reserve(name.size() + 2);
    escaped.push_back('\'');
    for (const char ch : name) {
        escaped.push_back(ch);
        if (ch == '\'') escaped.push_back('\'');
    }
    escaped.push_back('\'');
    return escaped;
}

std::string absoluteReferenceToken(std::string token) {
    token.erase(std::remove(token.begin(), token.end(), '$'), token.end());
    if (token.empty()) return token;
    std::size_t letters = 0;
    while (letters < token.size() && std::isalpha(static_cast<unsigned char>(token[letters]))) ++letters;
    if (letters == token.size()) return "$" + token;
    if (letters == 0) return "$" + token;
    return "$" + token.substr(0, letters) + "$" + token.substr(letters);
}

std::string absoluteReference(std::string reference) {
    if (!reference.empty() && reference.front() == '=') reference.erase(reference.begin());
    const auto bang = reference.rfind('!');
    if (bang != std::string::npos) reference = reference.substr(bang + 1);
    std::ostringstream result;
    std::size_t start = 0;
    bool firstArea = true;
    while (start <= reference.size()) {
        const auto comma = reference.find(',', start);
        const auto area = reference.substr(start, comma == std::string::npos ? std::string::npos : comma - start);
        const auto colon = area.find(':');
        if (!firstArea) result << ',';
        firstArea = false;
        if (colon == std::string::npos) result << absoluteReferenceToken(area);
        else result << absoluteReferenceToken(area.substr(0, colon)) << ':' << absoluteReferenceToken(area.substr(colon + 1));
        if (comma == std::string::npos) break;
        start = comma + 1;
    }
    return result.str();
}

std::string qualifiedPrintReference(const std::string& sheetName, const std::string& reference) {
    std::ostringstream result;
    const auto absolute = absoluteReference(reference);
    std::size_t start = 0;
    bool first = true;
    while (start <= absolute.size()) {
        const auto comma = absolute.find(',', start);
        if (!first) result << ',';
        first = false;
        result << quoteSheetName(sheetName) << '!' << absolute.substr(start, comma == std::string::npos ? std::string::npos : comma - start);
        if (comma == std::string::npos) break;
        start = comma + 1;
    }
    return result.str();
}

std::string localPrintReference(std::string value) {
    std::ostringstream result;
    std::size_t start = 0;
    bool first = true;
    while (start <= value.size()) {
        const auto comma = value.find(',', start);
        auto area = value.substr(start, comma == std::string::npos ? std::string::npos : comma - start);
        const auto bang = area.rfind('!');
        if (bang != std::string::npos) area = area.substr(bang + 1);
        area.erase(std::remove(area.begin(), area.end(), '$'), area.end());
        if (!first) result << ',';
        first = false;
        result << area;
        if (comma == std::string::npos) break;
        start = comma + 1;
    }
    return result.str();
}

std::string contentTypes(std::size_t sheetCount,
                         const std::vector<std::size_t>& chartsheetIds,
                         const std::vector<std::size_t>& printerSettingsIds,
                         std::size_t tableCount,
                         std::size_t commentCount,
                         const std::vector<std::size_t>& drawingIds,
                         const std::vector<xlpp::PreservedPart>& preserved,
                         bool strict,
                         bool hasSst,
                         const std::vector<std::size_t>& chartIds,
                         const std::vector<std::size_t>& pivotIds,
                         const std::vector<std::size_t>& pivotCacheIds,
                         bool hasCustomProperties,
                         bool macroEnabled,
                         bool templateWorkbook,
                         const std::vector<std::size_t>& slicerCacheIds = {},
                         const std::vector<std::size_t>& slicerIds = {}) {
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
    const char* workbookContentType = nullptr;
    if (templateWorkbook) {
        workbookContentType = macroEnabled
            ? "application/vnd.ms-excel.template.macroEnabled.main+xml"
            : "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
    } else {
        workbookContentType = macroEnabled
            ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    }
    addOverride("xl/workbook.xml", workbookContentType);
    addOverride("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");

    for (std::size_t index = 1; index <= sheetCount; ++index)
        addOverride("xl/worksheets/sheet" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    for (const auto index : chartsheetIds)
        addOverride("xl/chartsheets/sheet" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml");
    for (const auto index : printerSettingsIds)
        addOverride("xl/printerSettings/printerSettings" + std::to_string(index) + ".bin",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings");
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
                    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    for (const auto index : pivotIds) {
        addOverride("xl/pivotTables/pivotTable" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");
    }
    for (const auto index : pivotCacheIds) {
        addOverride("xl/pivotCache/pivotCacheDefinition" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
        addOverride("xl/pivotCache/pivotCacheRecords" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml");
    }
    for (const auto index : slicerCacheIds)
        addOverride("xl/slicerCaches/slicerCache" + std::to_string(index) + ".xml",
                    "application/vnd.ms-excel.slicerCache+xml");
    for (const auto index : slicerIds)
        addOverride("xl/slicers/slicer" + std::to_string(index) + ".xml",
                    "application/vnd.ms-excel.slicer+xml");
    for (const auto& part : preserved) {
        if (!part.overrideType.empty()) addOverride(part.name, part.overrideType);
        else if (!part.defaultType.empty()) addDefault(part.extension, part.defaultType);
    }
    xml << "</Types>";
    return xml.str();
}

// Package parts that `save` regenerates from the model; everything else is
// preserved verbatim across a load/save round trip.
bool isRegeneratedPart(const std::string& name) {
    return name == "[Content_Types].xml" || name == "_rels/.rels"
        || name == "docProps/core.xml" || name == "docProps/app.xml" || name == "docProps/custom.xml"
        || name == "xl/workbook.xml" || name == "xl/_rels/workbook.xml.rels"
        || name == "xl/styles.xml" || name == "xl/sharedStrings.xml"
        || name.rfind("xl/worksheets/", 0) == 0;
}

// Effective content types for preserved parts: <Override> rules win; parts
// without one fall back to the matching <Default> rule for their extension.
std::string extensionOf(const std::string& name) {
    const auto dot = name.rfind('.');
    if (dot == std::string::npos || dot + 1 == name.size()) return {};
    return name.substr(dot + 1);
}



std::string rootrels(bool strict, bool hasCustomProperties){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\""<<nsRelsPkg(strict)<<"\"><Relationship Id=\"rId1\" Type=\""<<nsRelsDoc(strict)<<"/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\""<<nsRelsPkg(strict)<<"/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\""<<nsRelsDoc(strict)<<"/extended-properties\" Target=\"docProps/app.xml\"/>";if(hasCustomProperties)x<<"<Relationship Id=\"rId4\" Type=\""<<nsRelsDoc(strict)<<"/custom-properties\" Target=\"docProps/custom.xml\"/>";x<<"</Relationships>";return x.str();}
std::string corePropertiesXml(const xlpp::DocumentProperties& p, bool strict){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\""<<nsCoreProps(strict)<<"\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"; if(!p.title().empty())x<<"<dc:title>"<<xmlEscape(p.title())<<"</dc:title>";if(!p.subject().empty())x<<"<dc:subject>"<<xmlEscape(p.subject())<<"</dc:subject>";if(!p.creator().empty())x<<"<dc:creator>"<<xmlEscape(p.creator())<<"</dc:creator>";if(!p.description().empty())x<<"<dc:description>"<<xmlEscape(p.description())<<"</dc:description>";if(!p.keywords().empty())x<<"<cp:keywords>"<<xmlEscape(p.keywords())<<"</cp:keywords>";if(!p.category().empty())x<<"<cp:category>"<<xmlEscape(p.category())<<"</cp:category>";if(!p.lastModifiedBy().empty())x<<"<cp:lastModifiedBy>"<<xmlEscape(p.lastModifiedBy())<<"</cp:lastModifiedBy>";x<<"</cp:coreProperties>";return x.str();}
std::string appPropertiesXml(bool strict){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\""<<nsExtProps(strict)<<"\" xmlns:vt=\""<<nsVTypes(strict)<<"\"><Application>XL++</Application><AppVersion>1.0</AppVersion></Properties>";return x.str();}
std::string customPropertiesXml(const xlpp::CustomProperties& props){std::ostringstream x;x<<R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">)";for(std::size_t i=0;i<props.items().size();++i){const auto&p=props.items()[i];x<<"<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\""<<(i+2)<<"\" name=\""<<xmlEscape(p.name())<<"\"><vt:"<<xmlEscape(p.type())<<">"<<xmlEscape(p.value())<<"</vt:"<<xmlEscape(p.type())<<"></property>";}x<<"</Properties>";return x.str();}

std::vector<std::string> resolvedWorksheetVbaCodeNames(const std::deque<xlpp::Worksheet>& sheets) {
    auto lower = [](std::string value) {
        std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
        return value;
    };
    std::set<std::string> used;
    std::vector<std::string> result(sheets.size());
    for (std::size_t i = 0; i < sheets.size(); ++i) {
        if (sheets[i].vbaCodeName().empty()) continue;
        xlpp::internal::validateVbaModuleName(sheets[i].vbaCodeName());
        const auto key = lower(sheets[i].vbaCodeName());
        if (!used.insert(key).second)
            throw std::invalid_argument("Duplicate worksheet VBA code name: " + sheets[i].vbaCodeName());
        result[i] = sheets[i].vbaCodeName();
    }
    std::size_t candidate = 1;
    for (std::size_t i = 0; i < result.size(); ++i) {
        if (!result[i].empty()) continue;
        for (;;) {
            auto name = "Sheet" + std::to_string(candidate++);
            if (used.insert(lower(name)).second) {
                result[i] = std::move(name);
                break;
            }
        }
    }
    return result;
}

std::vector<std::string> ensureWorksheetVbaCodeNames(std::deque<xlpp::Worksheet>& sheets) {
    const auto resolved = resolvedWorksheetVbaCodeNames(static_cast<const std::deque<xlpp::Worksheet>&>(sheets));
    for (std::size_t i = 0; i < sheets.size(); ++i)
        if (sheets[i].vbaCodeName().empty()) sheets[i].setVbaCodeName(resolved[i]);
    return resolved;
}

// Serialize every worksheet to XML, using a ThreadPool when workers > 1.
// Output is indexed by worksheet order and is identical to the sequential result.
// A cache snapshot is retained for deterministic output, but mutable-reference APIs require re-serialization for correctness.
std::vector<std::string> serializeSheets(const std::deque<xlpp::Worksheet>& sheets,
                                         const StyleCatalog& styles, const DxfCatalog& dxfs,
                                         bool date1904, bool strict, const std::vector<std::string>* vbaCodeNames, std::size_t workers,
                                         bool parallelRows,
                                         const SstIndex* sstIndex,
                                         std::vector<std::string>* cache,
                                         bool& cacheStrict, bool& cacheDate1904) {
    std::vector<std::string> result(sheets.size());
    std::vector<char> needsSerialize(sheets.size(), true);
    if (cache) {
        // Worksheet exposes mutable references (Cell&, Style&, AutoFilter&, ...).
        // A caller may keep one of those references and mutate it after a save,
        // bypassing Worksheet::dirty_. Re-serializing is therefore required for
        // correctness until mutations are tracked by owning proxy objects.
        if (cacheStrict != strict || cacheDate1904 != date1904) cache->clear();
    }

    // Serialize dirty sheets in parallel
    std::vector<std::size_t> dirtyIndexes;
    for (std::size_t i = 0; i < sheets.size(); ++i)
        if (needsSerialize[i]) dirtyIndexes.push_back(i);

    auto serializeOne = [&](std::size_t i) {
        const auto codeName = vbaCodeNames ? vbaCodeNames->at(i) : std::string{};
        result[i] = sheetXml(sheets[i], styles, dxfs, date1904, strict, sstIndex, 0, codeName);
    };

    if (workers > 1 && dirtyIndexes.size() > 1) {
        xlpp::internal::ThreadPool pool(std::min(workers, dirtyIndexes.size()));
        pool.parallelFor(0, dirtyIndexes.size(), [&](std::size_t j) {
            serializeOne(dirtyIndexes[j]);
        });
    } else if (parallelRows && workers > 1 && dirtyIndexes.size() == 1) {
        for (auto i : dirtyIndexes) {
            const auto codeName = vbaCodeNames ? vbaCodeNames->at(i) : std::string{};
            result[i] = sheetXml(sheets[i], styles, dxfs, date1904, strict, sstIndex, workers, codeName);
        }
    } else {
        for (auto i : dirtyIndexes) serializeOne(i);
    }

    // Update cache
    if (cache) {
        if (cache->size() < sheets.size()) cache->resize(sheets.size());
        for (std::size_t i = 0; i < sheets.size(); ++i)
            (*cache)[i] = result[i];
        cacheStrict = strict;
        cacheDate1904 = date1904;
    }
    return result;
}

void loadCharts(xlpp::Worksheet& ws, const std::string& sheetXml, const xlpp::internal::ZipArchive& z,
                const std::string& sheetPart) {
    const auto drawings = xlpp::internal::tags(sheetXml, "drawing");
    if (drawings.empty()) return;
    const auto sheetSlash = sheetPart.find_last_of('/');
    const auto sheetFile = sheetPart.substr(sheetSlash + 1);
    const auto sheetRelsPart = sheetPart.substr(0, sheetSlash + 1) + "_rels/" + sheetFile + ".rels";
    if (!z.contains(sheetRelsPart)) return;
    std::unordered_map<std::string, std::string> sheetRelationships;
    for (const auto& rel : xlpp::internal::tags(z.get(sheetRelsPart), "Relationship"))
        if (xlpp::internal::attribute(rel, "Type").find("/drawing") != std::string::npos)
            sheetRelationships[xlpp::internal::attribute(rel, "Id")] = xlpp::internal::attribute(rel, "Target");

    for (const auto& drawingNode : drawings) {
        const auto relationshipId = xlpp::internal::attribute(drawingNode, "r:id");
        const auto relationship = sheetRelationships.find(relationshipId);
        if (relationship == sheetRelationships.end()) continue;
        const auto drawingPart = resolvePackagePart(sheetPart, relationship->second);
        if (!z.contains(drawingPart)) continue;
        const auto drawingXmlText = z.get(drawingPart);
        const auto drawingSlash = drawingPart.find_last_of('/');
        const auto drawingFile = drawingPart.substr(drawingSlash + 1);
        const auto drawingRelsPart = drawingPart.substr(0, drawingSlash + 1) + "_rels/" + drawingFile + ".rels";
        if (!z.contains(drawingRelsPart)) continue;
        std::unordered_map<std::string, std::string> drawingRelationships;
        for (const auto& rel : xlpp::internal::tags(z.get(drawingRelsPart), "Relationship"))
            if (xlpp::internal::attribute(rel, "Type").find("/chart") != std::string::npos)
                drawingRelationships[xlpp::internal::attribute(rel, "Id")] = xlpp::internal::attribute(rel, "Target");

        const auto loadAnchors = [&](const char* prefixedName, const char* localName, xlpp::DrawingAnchorType anchorType) {
            for (const auto& anchorNode : drawingTags(drawingXmlText, prefixedName, localName)) {
                const auto frames = drawingTags(anchorNode, "xdr:graphicFrame", "graphicFrame");
                if (frames.empty()) continue;
                const auto& frame = frames.front();
                const auto chartRefs = drawingTags(frame, "c:chart", "chart");
                if (chartRefs.empty()) continue;
                const auto chartRelationshipId = xlpp::internal::attribute(chartRefs.front(), "r:id");
                const auto chartRelationship = drawingRelationships.find(chartRelationshipId);
                if (chartRelationship == drawingRelationships.end()) continue;
                const auto chartPart = resolvePackagePart(drawingPart, chartRelationship->second);
                if (!z.contains(chartPart)) continue;
                const auto chartText = z.get(chartPart);

                auto plots = parseChartPlots(chartText);
                const auto primaryType = plots.empty() ? parseChartType(chartText) : plots.front().type;
                xlpp::Chart chart(primaryType);
                chart.setGrouping(plots.empty() ? parseChartGrouping(chartText, chart.type()) : plots.front().grouping);
                chart.setThemePalette(parseChartThemePalette(z));
                chart.setStyleResources(parseChartStyleResources(z, chartPart));
                const auto pivotSources = drawingTags(chartText, "c:pivotSource", "pivotSource");
                if (!pivotSources.empty()) {
                    xlpp::ChartPivotSource pivotSource;
                    pivotSource.present = true;
                    pivotSource.pivotTableName = drawingTagText(pivotSources.front(), "c:name", "name");
                    const auto formatIds = drawingTags(pivotSources.front(), "c:fmtId", "fmtId");
                    if (!formatIds.empty()) {
                        const auto value = xlpp::internal::attribute(formatIds.front(), "val");
                        if (!value.empty()) {
                            try { pivotSource.formatId = std::stoi(value); } catch (...) {}
                        }
                    }
                    chart.setPivotSource(std::move(pivotSource));
                }
                chart.setTitle(chartTitleText(chartText));
                const auto chartTitleNodes = drawingTags(chartText, "c:title", "title");
                if (!chartTitleNodes.empty()) chart.setTitleRichText(parseChartRichText(chartTitleNodes.front()));
                chart.setView3D(parseChartView3D(chartText));
                chart.setFloorFormat(parseChartWallFormat(chartText, "c:floor", "floor"));
                chart.setSideWallFormat(parseChartWallFormat(chartText, "c:sideWall", "sideWall"));
                chart.setBackWallFormat(parseChartWallFormat(chartText, "c:backWall", "backWall"));
                auto axes = parseChartAxes(chartText, plots);
                std::uint64_t primaryXAxisId = 0;
                std::uint64_t primaryYAxisId = 0;
                if (!plots.empty() && plots.front().axisIds.size() >= 2) {
                    primaryXAxisId = plots.front().axisIds[0];
                    primaryYAxisId = plots.front().axisIds[1];
                }
                for (auto& plot : plots) {
                    plot.usesSecondaryAxes = std::any_of(plot.axisIds.begin(), plot.axisIds.end(), [&](std::uint64_t axisId) {
                        const auto it = std::find_if(axes.begin(), axes.end(), [&](const auto& axis) { return axis.id == axisId; });
                        return it != axes.end() && it->secondary;
                    });
                }
                chart.setPrimaryAxisIds(primaryXAxisId, primaryYAxisId);
                chart.setXAxisTitle(primaryXAxisId != 0 ? axisTitleById(axes, primaryXAxisId) : std::string{});
                chart.setYAxisTitle(primaryYAxisId != 0 ? axisTitleById(axes, primaryYAxisId) : std::string{});
                if (primaryXAxisId == 0 || primaryYAxisId == 0) {
                    const bool xyValueAxes = chart.type() == xlpp::Chart::Type::Scatter || chart.type() == xlpp::Chart::Type::Bubble;
                    if (primaryXAxisId == 0) chart.setXAxisTitle(axisTitleText(chartText, xyValueAxes ? "c:valAx" : "c:catAx",
                                                                               xyValueAxes ? "valAx" : "catAx", 0));
                    if (primaryYAxisId == 0) chart.setYAxisTitle(axisTitleText(chartText, "c:valAx", "valAx", xyValueAxes ? 1 : 0));
                }
                chart.setAxes(axes);
                chart.setPlots(plots);
                const auto scatterStyles = drawingTags(chartText, "c:scatterStyle", "scatterStyle");
                if (!scatterStyles.empty()) {
                    const auto scatterStyle = xlpp::internal::attribute(scatterStyles.front(), "val");
                    if (!scatterStyle.empty()) chart.setScatterStyle(scatterStyle);
                }
                const auto plotAreas = drawingTags(chartText, "c:plotArea", "plotArea");
                if (!plotAreas.empty()) {
                    chart.setPlotAreaLayout(parseChartManualLayout(plotAreas.front()));
                    chart.setDataTable(parseChartDataTable(plotAreas.front()));
                    const auto plotSpPr = plotAreaDirectSpPr(plotAreas.front());
                    if (!plotSpPr.empty()) { chart.setPlotAreaLineFormat(parseChartLineFormat(plotSpPr)); chart.setPlotAreaFillFormat(parseChartFillFormat(plotSpPr)); }
                }
                const auto chartAreaSpPr = chartSpaceDirectSpPr(chartText);
                if (!chartAreaSpPr.empty()) { chart.setChartAreaLineFormat(parseChartLineFormat(chartAreaSpPr)); chart.setChartAreaFillFormat(parseChartFillFormat(chartAreaSpPr)); }
                const auto styleNodes = drawingTags(chartText, "c:style", "style");
                if (!styleNodes.empty()) chart.setStyle(xlpp::internal::attribute(styleNodes.front(), "val"));
                const auto legendNodes = drawingTags(chartText, "c:legend", "legend");
                chart.setShowLegend(!legendNodes.empty());
                if (!legendNodes.empty()) {
                    const auto& legendXml = legendNodes.front();
                    const auto positions = drawingTags(legendXml, "c:legendPos", "legendPos");
                    if (!positions.empty()) chart.setLegendPosition(xlpp::internal::attribute(positions.front(), "val"));
                    xlpp::ChartLegendFormat legendFormat; legendFormat.present = true;
                    legendFormat.overlay = chartBoolValue(legendXml, "c:overlay", "overlay");
                    legendFormat.layout = parseChartManualLayout(legendXml);
                    const auto spPr = drawingTags(legendXml, "c:spPr", "spPr");
                    if (!spPr.empty()) { legendFormat.line = parseChartLineFormat(spPr.front()); legendFormat.fill = parseChartFillFormat(spPr.front()); }
                    chart.setLegendFormat(std::move(legendFormat));
                }
                for (const auto& seriesNode : drawingTags(chartText, "c:ser", "ser")) {
                    xlpp::ChartSeries series;
                    const auto txNodes = drawingTags(seriesNode, "c:tx", "tx");
                    if (!txNodes.empty()) {
                        auto title = drawingTagText(txNodes.front(), "c:v", "v");
                        const auto titleReference = drawingTagText(txNodes.front(), "c:f", "f");
                        auto titleCache = parseChartSeriesCache(txNodes.front());
                        if (title.empty() && titleCache.present && !titleCache.points.empty()) title = titleCache.points.front().value;
                        if (title.empty()) title = titleReference;
                        series.setTitle(title);
                        series.setTitleReference(titleReference);
                        series.setTitleCache(std::move(titleCache));
                    }
                    const auto valNodes = drawingTags(seriesNode, "c:val", "val");
                    if (!valNodes.empty()) { series.setValuesReference(drawingTagText(valNodes.front(), "c:f", "f")); series.setValuesCache(parseChartSeriesCache(valNodes.front())); }
                    else {
                        const auto yValNodes = drawingTags(seriesNode, "c:yVal", "yVal");
                        if (!yValNodes.empty()) { series.setValuesReference(drawingTagText(yValNodes.front(), "c:f", "f")); series.setValuesCache(parseChartSeriesCache(yValNodes.front())); }
                    }
                    const auto catNodes = drawingTags(seriesNode, "c:cat", "cat");
                    if (!catNodes.empty()) { series.setCategoriesReference(drawingTagText(catNodes.front(), "c:f", "f")); series.setCategoriesCache(parseChartSeriesCache(catNodes.front())); }
                    else {
                        const auto xValNodes = drawingTags(seriesNode, "c:xVal", "xVal");
                        if (!xValNodes.empty()) { series.setCategoriesReference(drawingTagText(xValNodes.front(), "c:f", "f")); series.setCategoriesCache(parseChartSeriesCache(xValNodes.front())); }
                    }
                    const auto bubbleNodes = drawingTags(seriesNode, "c:bubbleSize", "bubbleSize");
                    if (!bubbleNodes.empty()) {
                        series.setBubbleSizeReference(drawingTagText(bubbleNodes.front(), "c:f", "f"));
                        series.setBubbleSizeCache(parseChartSeriesCache(bubbleNodes.front()));
                    }
                    series.setDataLabels(parseChartDataLabels(seriesNode));
                    series.setTrendlines(parseChartTrendlines(seriesNode));
                    series.setErrorBars(parseChartErrorBars(seriesNode));
                    const auto seriesSpPr = seriesDirectSpPr(seriesNode);
                    if (!seriesSpPr.empty()) {
                        series.setLineFormat(parseChartLineFormat(seriesSpPr));
                        series.setFillFormat(parseChartFillFormat(seriesSpPr));
                    }
                    series.setMarkerFormat(parseChartMarkerFormat(seriesNode));
                    series.setDataPoints(parseChartDataPoints(seriesNode));
                    chart.addSeries(std::move(series));
                }

                const auto anchorInfo = parseChartAnchorInfo(anchorNode, anchorType, frame);
                chart.setAnchorInfo(anchorInfo);
                if (anchorInfo.widthEmu > 0) chart.setWidth(std::max(1, static_cast<int>(std::llround(static_cast<double>(anchorInfo.widthEmu) / 9525.0))));
                if (anchorInfo.heightEmu > 0) chart.setHeight(std::max(1, static_cast<int>(std::llround(static_cast<double>(anchorInfo.heightEmu) / 9525.0))));
                const auto nonVisual = drawingTags(frame, "xdr:cNvPr", "cNvPr");
                std::string objectId;
                std::string objectName = "Chart";
                if (!nonVisual.empty()) {
                    objectId = xlpp::internal::attribute(nonVisual.front(), "id");
                    const auto parsedName = xlpp::internal::attribute(nonVisual.front(), "name");
                    if (!parsedName.empty()) objectName = parsedName;
                }
                chart.setStableId(drawingPart + "#" + (objectId.empty() ? chartRelationshipId : objectId));
                chart.setSourceDrawingPart(drawingPart);
                chart.setSourceChartPart(chartPart);
                chart.setSourceRelationshipId(chartRelationshipId);
                chart.setDrawingObjectName(objectName);
                chart.setImported(true);
                ws.addLoadedChart(std::move(chart));
            }
        };

        loadAnchors("xdr:oneCellAnchor", "oneCellAnchor", xlpp::DrawingAnchorType::OneCell);
        loadAnchors("xdr:twoCellAnchor", "twoCellAnchor", xlpp::DrawingAnchorType::TwoCell);
        loadAnchors("xdr:absoluteAnchor", "absoluteAnchor", xlpp::DrawingAnchorType::Absolute);
    }
}


std::string readFilePrefix(const std::filesystem::path& path, std::size_t bytes) {
    std::string prefix(bytes, '\0');
    std::ifstream in(path, std::ios::binary);
    if (in) in.read(prefix.data(), static_cast<std::streamsize>(bytes));
    return prefix;
}
void parseSheet(xlpp::Worksheet& ws, const std::string& xml, const xlpp::internal::ZipArchive& z, const std::string& target, const StyleCatalog& styleCatalog, const std::vector<xlpp::Style>& dxfStyles, const std::vector<LoadedSharedString>& shared, bool date1904, std::size_t maxCells, std::size_t& totalCells) {
    using namespace xlpp;
    // Sparklines live in the x14 worksheet extension list. Parse the modeled
    // sparklineGroup/sparkline structure and keep the raw block for
    // byte-preserving round-trips when no group is edited.
    for (const auto& sparklineGroupsNode : internal::tags(xml, "x14:sparklineGroups")) {
        std::vector<xlpp::SparklineGroup> groups;
        std::string rawGroups = sparklineGroupsNode;
        for (const auto& groupNode : internal::tags(sparklineGroupsNode, "x14:sparklineGroup")) {
            xlpp::SparklineGroup group;
            group.rawXml = groupNode;
            const auto type = internal::attribute(groupNode, "type");
            if (!type.empty()) group.type = type;
            const auto lineWeight = internal::attribute(groupNode, "lineWeight");
            if (lineWeight == "1.5") group.lineStyle = "smooth";
            group.displayHidden = internal::attribute(groupNode, "displayHidden") == "1";
            group.displayXAxis = internal::attribute(groupNode, "displayXAxis") == "1";
            group.displayMarkers = internal::attribute(groupNode, "markers") == "1";
            group.high = internal::attribute(groupNode, "high") == "1";
            group.low = internal::attribute(groupNode, "low") == "1";
            group.first = internal::attribute(groupNode, "first") == "1";
            group.last = internal::attribute(groupNode, "last") == "1";
            group.negative = internal::attribute(groupNode, "negative") == "1";
            group.colorSeries = internal::attribute(groupNode, "colorSeries") == "1";
            group.colorAxis = internal::attribute(groupNode, "colorAxis") == "1";
            group.colorMarkers = internal::attribute(groupNode, "colorMarkers") == "1";
            group.colorFirst = internal::attribute(groupNode, "colorFirst") == "1";
            group.colorLast = internal::attribute(groupNode, "colorLast") == "1";
            group.colorHigh = internal::attribute(groupNode, "colorHigh") == "1";
            group.colorLow = internal::attribute(groupNode, "colorLow") == "1";
            group.rightToLeft = internal::attribute(groupNode, "rightToLeft") == "1";
            const auto colorSeriesNodes = internal::tags(groupNode, "x14:colorSeries");
            if (!colorSeriesNodes.empty()) {
                const auto rgbNodes = internal::tags(colorSeriesNodes.front(), "x14:rgb");
                if (!rgbNodes.empty()) group.markersColor = internal::attribute(rgbNodes.front(), "value");
            }
            const auto colorNegativeNodes = internal::tags(groupNode, "x14:colorNegative");
            if (!colorNegativeNodes.empty()) {
                const auto rgbNodes = internal::tags(colorNegativeNodes.front(), "x14:rgb");
                if (!rgbNodes.empty()) group.negativeColor = internal::attribute(rgbNodes.front(), "value");
            }
            const auto colorAxisNodes = internal::tags(groupNode, "x14:colorAxis");
            if (!colorAxisNodes.empty()) {
                const auto rgbNodes = internal::tags(colorAxisNodes.front(), "x14:rgb");
                if (!rgbNodes.empty()) group.axisColor = internal::attribute(rgbNodes.front(), "value");
            }
            const auto dateAxisNodes = internal::tags(groupNode, "xm:f");
            if (!dateAxisNodes.empty()) group.dateAxis = internal::tagText(groupNode, "xm:f");
            const auto sparklineNodes = internal::tags(groupNode, "x14:sparkline");
            for (const auto& sparklineNode : sparklineNodes) {
                xlpp::Sparkline sparkline;
                sparkline.reference = internal::tagText(sparklineNode, "xm:f");
                sparkline.location = internal::tagText(sparklineNode, "xm:sqref");
                group.sparklines.push_back(std::move(sparkline));
            }
            groups.push_back(std::move(group));
        }
        if (!groups.empty()) ws.setSparklineGroups(std::move(groups));
        if (groups.empty() && !rawGroups.empty()) ws.setSparklineGroupsRawXml(std::move(rawGroups));
    }
    const auto sheetProperties = internal::tags(xml, "sheetPr");
    if (!sheetProperties.empty()) {
        const auto codeName = internal::attribute(sheetProperties.front(), "codeName");
        if (!codeName.empty()) ws.setVbaCodeName(codeName);
    }
    const auto sheetFormats = internal::tags(xml, "sheetFormatPr");
    if (!sheetFormats.empty()) ws.setLoadedSheetFormatPrXml(sheetFormats.front());
for (const auto& margin : internal::tags(xml, "pageMargins")) {
    const auto setDouble=[&](const char* attributeName, auto setter){const auto attributeValue=internal::attribute(margin,attributeName);if(!attributeValue.empty()) setter(internal::parseDoubleExact(attributeValue, attributeName));};
    setDouble("left",[&](double v){ws.pageMargins().setLeft(v);}); setDouble("right",[&](double v){ws.pageMargins().setRight(v);});
    setDouble("top",[&](double v){ws.pageMargins().setTop(v);}); setDouble("bottom",[&](double v){ws.pageMargins().setBottom(v);});
    setDouble("header",[&](double v){ws.pageMargins().setHeader(v);}); setDouble("footer",[&](double v){ws.pageMargins().setFooter(v);});
}
for (const auto& setup : internal::tags(xml, "pageSetup")) {
    const auto orientation=internal::attribute(setup,"orientation");
    if(orientation=="portrait") ws.pageSetup().setOrientation(PageOrientation::Portrait); else if(orientation=="landscape") ws.pageSetup().setOrientation(PageOrientation::Landscape);
    const auto paper=internal::attribute(setup,"paperSize"); if(!paper.empty()) ws.pageSetup().setPaperSize(static_cast<PaperSize>(internal::parseIntegerExact<unsigned>(paper, "paperSize")));
    const auto scale=internal::attribute(setup,"scale"); if(!scale.empty()) ws.pageSetup().setScale(internal::parseIntegerExact<unsigned>(scale, "page scale"));
    const auto fw=internal::attribute(setup,"fitToWidth"), fh=internal::attribute(setup,"fitToHeight"); if(!fw.empty()){ws.pageSetup().setFitToPage(true);ws.pageSetup().setFitToWidth(internal::parseIntegerExact<unsigned>(fw, "fitToWidth"));} if(!fh.empty())ws.pageSetup().setFitToHeight(internal::parseIntegerExact<unsigned>(fh, "fitToHeight"));
    ws.pageSetup().setBlackAndWhite(internal::attribute(setup,"blackAndWhite")=="1"); ws.pageSetup().setDraft(internal::attribute(setup,"draft")=="1");
    const auto first=internal::attribute(setup,"firstPageNumber"); if(!first.empty())ws.pageSetup().setFirstPageNumber(internal::parseIntegerExact<unsigned>(first, "firstPageNumber")); ws.pageSetup().setUseFirstPageNumber(internal::attribute(setup,"useFirstPageNumber")=="1");
    const auto paperHeight=internal::attribute(setup,"paperHeight"); if(!paperHeight.empty()) ws.pageSetup().setPaperHeight(paperHeight);
    const auto paperWidth=internal::attribute(setup,"paperWidth"); if(!paperWidth.empty()) ws.pageSetup().setPaperWidth(paperWidth);
    const auto pageOrder=internal::attribute(setup,"pageOrder"); if(pageOrder=="downThenOver") ws.pageSetup().setPageOrder(PageOrder::DownThenOver); else if(pageOrder=="overThenDown") ws.pageSetup().setPageOrder(PageOrder::OverThenDown);
    const auto printerDefaults=internal::attribute(setup,"usePrinterDefaults"); if(!printerDefaults.empty()) ws.pageSetup().setUsePrinterDefaults(printerDefaults=="1"||printerDefaults=="true");
    const auto cellComments=internal::attribute(setup,"cellComments"); if(cellComments=="asDisplayed") ws.pageSetup().setCellComments(PageCellComments::AsDisplayed); else if(cellComments=="atEnd") ws.pageSetup().setCellComments(PageCellComments::AtEnd);
    const auto errors=internal::attribute(setup,"errors"); if(errors=="displayed") ws.pageSetup().setErrors(PageErrorDisplay::Displayed); else if(errors=="blank") ws.pageSetup().setErrors(PageErrorDisplay::Blank); else if(errors=="dash") ws.pageSetup().setErrors(PageErrorDisplay::Dash); else if(errors=="NA") ws.pageSetup().setErrors(PageErrorDisplay::NA);
    const auto hdpi=internal::attribute(setup,"horizontalDpi"); if(!hdpi.empty()) ws.pageSetup().setHorizontalDpi(internal::parseIntegerExact<unsigned>(hdpi,"horizontalDpi"));
    const auto vdpi=internal::attribute(setup,"verticalDpi"); if(!vdpi.empty()) ws.pageSetup().setVerticalDpi(internal::parseIntegerExact<unsigned>(vdpi,"verticalDpi"));
    const auto copies=internal::attribute(setup,"copies"); if(!copies.empty()) ws.pageSetup().setCopies(internal::parseIntegerExact<unsigned>(copies,"copies"));
    const auto setupRid=internal::attribute(setup,"r:id"); if(!setupRid.empty()) ws.pageSetup().setRelationshipId(setupRid);
}
for (const auto& printOptions : internal::tags(xml, "printOptions")) { ws.printOptions().setHorizontalCentered(internal::attribute(printOptions,"horizontalCentered")=="1"); ws.printOptions().setVerticalCentered(internal::attribute(printOptions,"verticalCentered")=="1"); ws.printOptions().setHeadings(internal::attribute(printOptions,"headings")=="1"); ws.printOptions().setGridLines(internal::attribute(printOptions,"gridLines")=="1"); }
for (const auto& hf : internal::tags(xml, "headerFooter")) { ws.headerFooter().setDifferentOddEven(internal::attribute(hf,"differentOddEven")=="1"); ws.headerFooter().setDifferentFirst(internal::attribute(hf,"differentFirst")=="1"); ws.headerFooter().setOddHeader(internal::tagText(hf,"oddHeader")); ws.headerFooter().setOddFooter(internal::tagText(hf,"oddFooter")); ws.headerFooter().setEvenHeader(internal::tagText(hf,"evenHeader")); ws.headerFooter().setEvenFooter(internal::tagText(hf,"evenFooter")); ws.headerFooter().setFirstHeader(internal::tagText(hf,"firstHeader")); ws.headerFooter().setFirstFooter(internal::tagText(hf,"firstFooter")); }
for (const auto& rowBreaksTag : internal::tags(xml, "rowBreaks")) {
    std::vector<std::size_t> rows;
    for (const auto& brk : internal::tags(rowBreaksTag, "brk")) {
        const auto idText = internal::attribute(brk, "id");
        std::size_t id = 0;
        if (internal::tryParseIntegerExact(idText, id)) rows.push_back(id);
    }
    if (!rows.empty()) ws.setRowBreaks(std::move(rows));
}
for (const auto& colBreaksTag : internal::tags(xml, "colBreaks")) {
    std::vector<std::size_t> cols;
    for (const auto& brk : internal::tags(colBreaksTag, "brk")) {
        const auto idText = internal::attribute(brk, "id");
        std::size_t id = 0;
        if (internal::tryParseIntegerExact(idText, id)) cols.push_back(id);
    }
    if (!cols.empty()) ws.setColumnBreaks(std::move(cols));
}
for (const auto& protectionNode : internal::tags(xml, "sheetProtection")) { ws.protection().setEnabled(true); ws.protection().setPasswordHash(internal::attribute(protectionNode,"password")); ws.protection().setSelectLockedCells(internal::attribute(protectionNode,"selectLockedCells")!="1"); ws.protection().setSelectUnlockedCells(internal::attribute(protectionNode,"selectUnlockedCells")!="1"); ws.protection().setFormatCells(internal::attribute(protectionNode,"formatCells")!="1"); ws.protection().setFormatColumns(internal::attribute(protectionNode,"formatColumns")!="1"); ws.protection().setFormatRows(internal::attribute(protectionNode,"formatRows")!="1"); ws.protection().setInsertRows(internal::attribute(protectionNode,"insertRows")!="1"); ws.protection().setInsertColumns(internal::attribute(protectionNode,"insertColumns")!="1"); ws.protection().setDeleteRows(internal::attribute(protectionNode,"deleteRows")!="1"); ws.protection().setDeleteColumns(internal::attribute(protectionNode,"deleteColumns")!="1"); ws.protection().setSort(internal::attribute(protectionNode,"sort")!="1"); ws.protection().setAutoFilter(internal::attribute(protectionNode,"autoFilter")!="1"); }
for (const auto& sv : internal::tags(xml, "sheetView")) {
    auto& view = ws.sheetView();
    const auto workbookViewId = internal::attribute(sv, "workbookViewId");
    if (!workbookViewId.empty()) view.setWorkbookViewId(internal::parseIntegerExact<int>(workbookViewId, "workbookViewId"));
    const auto zoom = internal::attribute(sv, "zoomScale");
    if (!zoom.empty()) view.setZoomScale(internal::parseIntegerExact<int>(zoom, "zoomScale"));
    const auto normalZoom = internal::attribute(sv, "zoomScaleNormal");
    if (!normalZoom.empty()) view.setZoomScaleNormal(internal::parseIntegerExact<int>(normalZoom, "zoomScaleNormal"));
    view.setShowGridLines(internal::attribute(sv, "showGridLines") != "0");
    view.setShowRowColHeaders(internal::attribute(sv, "showRowColHeaders") != "0");
    view.setTabSelected(internal::attribute(sv, "tabSelected") == "1");
    view.setRightToLeft(internal::attribute(sv, "rightToLeft") == "1");
    view.setShowOutlineSymbols(internal::attribute(sv, "showOutlineSymbols") != "0");
}
for (const auto& tc : internal::tags(xml, "tabColor")) {
    const auto rgb = internal::attribute(tc, "rgb");
    if (!rgb.empty()) ws.sheetView().setTabColor(std::string(rgb));
}
for (auto& pane : internal::tags(xml, "pane")) {
    const auto state = internal::attribute(pane, "state");
    const auto topLeft = internal::attribute(pane, "topLeftCell");
    if (state == "frozen" && !topLeft.empty()) ws.freezePanes(topLeft);
    auto& view = ws.sheetView();
    const auto activePane = internal::attribute(pane, "activePane");
    if (!activePane.empty()) view.setPane(activePane);
    if (!topLeft.empty()) view.setTopLeftCell(topLeft);
    const auto xSplit = internal::attribute(pane, "xSplit");
    if (!xSplit.empty()) view.setXSplit(static_cast<int>(internal::parseDoubleExact(xSplit, "xSplit")));
    const auto ySplit = internal::attribute(pane, "ySplit");
    if (!ySplit.empty()) view.setYSplit(static_cast<int>(internal::parseDoubleExact(ySplit, "ySplit")));
}
for (auto& col : internal::tags(xml, "col")) {
    const auto minText = internal::attribute(col, "min");
    if (minText.empty()) continue;
    const auto minColumn = internal::parseIntegerExact<std::size_t>(minText, "column min");
    const auto maxText = internal::attribute(col, "max");
    const auto maxColumn = maxText.empty() ? minColumn : internal::parseIntegerExact<std::size_t>(maxText, "column max");
    if (minColumn == 0 || minColumn > maxColumn) throw std::runtime_error("Malformed sheet: invalid column range");
    if (maxColumn - minColumn + 1 > 1048576u) throw std::runtime_error("Malformed sheet: column range too large");
    for (std::size_t column = minColumn; column <= maxColumn; ++column) {
        auto& dimension = ws.columnDimension(column);
        const auto width = internal::attribute(col, "width");
        if (!width.empty()) dimension.width = internal::parseDoubleExact(width, "column width");
        dimension.hidden = internal::attribute(col, "hidden") == "1";
        dimension.bestFit = internal::attribute(col, "bestFit") == "1";
        const auto outline = internal::attribute(col, "outlineLevel");
        if (!outline.empty()) dimension.outlineLevel = internal::parseIntegerExact<int>(outline, "outlineLevel");
        dimension.collapsed = internal::attribute(col, "collapsed") == "1";
    }
}
internal::tagsForEach(xml, "row", [&](std::string_view row) {
    const auto indexText = internal::attribute(row, "r");
    if (indexText.empty()) return;
    auto& dimension = ws.rowDimension(internal::parseIntegerExact<std::size_t>(indexText, "row index"));
    const auto height = internal::attribute(row, "ht");
    if (!height.empty()) dimension.height = internal::parseDoubleExact(height, "row height");
    dimension.hidden = internal::attribute(row, "hidden") == "1";
    const auto outline = internal::attribute(row, "outlineLevel");
    if (!outline.empty()) dimension.outlineLevel = internal::parseIntegerExact<int>(outline, "outlineLevel");
    dimension.collapsed = internal::attribute(row, "collapsed") == "1";
});

for (auto& autoFilterTag : internal::tags(xml, "autoFilter")) {
    auto& autoFilter = ws.autoFilter();
    autoFilter.setReference(internal::attribute(autoFilterTag, "ref"));
    for (auto& columnTag : internal::tags(autoFilterTag, "filterColumn")) {
        const auto columnIdText = internal::attribute(columnTag, "colId");
        if (columnIdText.empty()) continue;
        auto& column = autoFilter.column(internal::parseIntegerExact<std::size_t>(columnIdText, "filter colId"));
        for (auto& filtersTag : internal::tags(columnTag, "filters")) {
            column.setIncludeBlank(internal::attribute(filtersTag, "blank") == "1");
            for (auto& filterTag : internal::tags(filtersTag, "filter"))
                column.addValue(internal::attribute(filterTag, "val"));
        }
        for (auto& customFiltersTag : internal::tags(columnTag, "customFilters")) {
            column.setAndMode(internal::attribute(customFiltersTag, "and") == "1");
            for (auto& customFilterTag : internal::tags(customFiltersTag, "customFilter"))
                column.addCustomFilter(parseFilterOperator(internal::attribute(customFilterTag, "operator")),
                                       internal::attribute(customFilterTag, "val"));
        }
        for (auto& top10Tag : internal::tags(columnTag, "top10")) {
            const bool top = internal::attribute(top10Tag, "top") == "1";
            const bool percent = internal::attribute(top10Tag, "percent") == "1";
            int value = 10;
            const auto valText = internal::attribute(top10Tag, "val");
            long long valueLL = 0;
            if (internal::tryParseIntegerExact(valText, valueLL)) value = static_cast<int>(valueLL);
            column.setTop10(top, value, percent);
        }
        for (auto& dynamicTag : internal::tags(columnTag, "dynamicFilter")) {
            const auto type = internal::attribute(dynamicTag, "type");
            std::optional<double> val;
            double parsedVal = 0.0;
            const auto valText = internal::attribute(dynamicTag, "val");
            if (!valText.empty() && internal::tryParseDoubleExact(valText, parsedVal)) val = parsedVal;
            if (!type.empty()) column.setDynamicFilter(type, val);
        }
        for (auto& extensionTag : internal::tags(columnTag, "extLst")) {
            column.setFilterExtension(extensionTag);
        }
    }
    for (auto& sortStateTag : internal::tags(autoFilterTag, "sortState")) {
        auto& sortState = autoFilter.sortState();
        sortState.setReference(internal::attribute(sortStateTag, "ref"));
        sortState.setCaseSensitive(internal::attribute(sortStateTag, "caseSensitive") == "1");
        for (auto& conditionTag : internal::tags(sortStateTag, "sortCondition"))
            sortState.addCondition(internal::attribute(conditionTag, "ref"),
                                   internal::attribute(conditionTag, "descending") == "1");
    }
}

for (auto& formattingTag : internal::tags(xml, "conditionalFormatting")) {
    const auto reference = internal::attribute(formattingTag, "sqref");
    if (reference.empty()) continue;
    auto& formatting = ws.conditionalFormatting().add(reference);
    for (auto& ruleTag : internal::tags(formattingTag, "cfRule")) {
        const auto type = internal::attribute(ruleTag, "type");
        const auto formulaNodes = internal::tags(ruleTag, "formula");
        std::vector<std::string> formulas;
        for (const auto& formulaNode : formulaNodes) formulas.push_back(internal::tagText(formulaNode, "formula"));
        ConditionalRule rule;
        if (type == "cellIs") {
            rule = ConditionalRule::cellIs(parseConditionalOperator(internal::attribute(ruleTag, "operator")), formulas.empty() ? std::string{} : formulas.front());
            rule.setFormulas(std::move(formulas));
        } else if (type == "dataBar") {
            rule = ConditionalRule::dataBar();
            for (const auto& dbTag : internal::tags(ruleTag, "dataBar")) {
                const auto dir = internal::attribute(dbTag, "direction");
                if (!dir.empty()) rule.getDataBar().direction = dir;
                rule.getDataBar().showValue = internal::attribute(dbTag, "showValue") != "0";
                rule.getDataBar().gradient = internal::attribute(dbTag, "gradient") != "0";
                std::vector<Cfvo> cfvos;
                for (const auto& cfvoTag : internal::tags(dbTag, "cfvo"))
                    cfvos.push_back(parseCfvo(cfvoTag));
                // Excel 2010+ writes min/mid/max (3 cfvo). XL++'s own writer
                // appends a trailing autoMin cfvo for the axis position, so
                // treat a third "autoMin" cfvo as the axis position.
                if (cfvos.size() >= 3 && cfvos[2].type != "autoMin") {
                    rule.getDataBar().min = cfvos[0];
                    rule.getDataBar().mid = cfvos[1];
                    rule.getDataBar().max = cfvos[2];
                } else {
                    if (cfvos.size() >= 1) rule.getDataBar().min = cfvos[0];
                    if (cfvos.size() >= 2) rule.getDataBar().max = cfvos[1];
                    if (cfvos.size() >= 3) {
                        const auto valText = internal::attribute(internal::tags(dbTag, "cfvo")[2], "val");
                        double position = 0.0;
                        if (!valText.empty() && internal::tryParseDoubleExact(valText, position))
                            rule.getDataBar().axisPosition = position;
                        else
                            rule.getDataBar().axisPosition = 0.0;
                    }
                }
                const auto colorTags = internal::tags(dbTag, "color");
                if (!colorTags.empty()) rule.getDataBar().color = internal::attribute(colorTags.front(), "rgb");
                const auto axisColor = internal::tags(dbTag, "axisColor");
                if (!axisColor.empty()) rule.getDataBar().axisColor = internal::attribute(axisColor.front(), "rgb");
                const auto negFill = internal::tags(dbTag, "negativeFillColor");
                if (!negFill.empty()) rule.getDataBar().negativeBarColor = internal::attribute(negFill.front(), "rgb");
                const auto negBorder = internal::tags(dbTag, "negativeBorderColor");
                if (!negBorder.empty()) rule.getDataBar().negativeBarBorderColor = internal::attribute(negBorder.front(), "rgb");
                const auto borderColor = internal::tags(dbTag, "borderColor");
                if (!borderColor.empty()) rule.getDataBar().borderColor = internal::attribute(borderColor.front(), "rgb");
            }
        } else if (type == "colorScale") {
            rule = ConditionalRule::colorScale();
            for (const auto& csTag : internal::tags(ruleTag, "colorScale")) {
                const auto cfvoTags = internal::tags(csTag, "cfvo");
                const auto colorTags = internal::tags(csTag, "color");
                std::size_t idx = 0;
                for (const auto& cfvoTag : cfvoTags) {
                    auto stop = parseCfvo(cfvoTag);
                    if (idx < colorTags.size()) stop.color = internal::attribute(colorTags[idx], "rgb");
                    rule.getColorScale().addStop(std::move(stop));
                    ++idx;
                }
            }
        } else if (type == "iconSet") {
            rule = ConditionalRule::iconSet();
            for (const auto& isTag : internal::tags(ruleTag, "iconSet")) {
                rule.getIconSet().icons = internal::attribute(isTag, "iconSet");
                rule.getIconSet().reverse = internal::attribute(isTag, "reverse") == "1";
                rule.getIconSet().showValue = internal::attribute(isTag, "showValue") != "0";
                for (const auto& cfvoTag : internal::tags(isTag, "cfvo"))
                    rule.getIconSet().addThreshold(parseCfvo(cfvoTag));
            }
        } else if (type == "containsText" || type == "notContainsText" || type == "beginsWith" || type == "endsWith") {
            rule = type == "containsText" ? ConditionalRule::containsText({})
                : type == "notContainsText" ? ConditionalRule::notContainsText({})
                : type == "beginsWith" ? ConditionalRule::beginsWith({}) : ConditionalRule::endsWith({});
            // Recover the searched text from the formula, e.g.
            // NOT(ISERROR(SEARCH("needle",A1))).
            std::string text;
            if (!formulas.empty()) {
                const auto& f = formulas.front();
                const auto q1 = f.find('"');
                if (q1 != std::string::npos) {
                    const auto q2 = f.find('"', q1 + 1);
                    if (q2 != std::string::npos) text = f.substr(q1 + 1, q2 - q1 - 1);
                }
            }
            rule.setText(std::move(text));
        } else if (type == "aboveAverage") {
            const bool above = internal::attribute(ruleTag, "aboveAverage") != "0";
            rule = above ? ConditionalRule::aboveAverage() : ConditionalRule::belowAverage();
            rule.setEqualAverage(internal::attribute(ruleTag, "equalAverage") == "1");
            rule.setStdDev(internal::attribute(ruleTag, "stdDev") == "1");
            const auto sd = internal::attribute(ruleTag, "stdDevVal");
            if (!sd.empty()) {
                int count = 2;
                if (internal::tryParseIntegerExact(sd, count)) rule.setStdDevCount(count);
            }
        } else if (type == "top10") {
            const bool bottom = internal::attribute(ruleTag, "bottom") == "1";
            const bool percent = internal::attribute(ruleTag, "percent") == "1";
            int rank = 10;
            const auto rankText = internal::attribute(ruleTag, "rank");
            if (!rankText.empty() && internal::tryParseIntegerExact(rankText, rank)) {}
            rule = ConditionalRule::top10(!bottom, rank, percent);
        } else if (type == "duplicateValues") {
            rule = ConditionalRule::duplicateValues();
        } else if (type == "uniqueValues") {
            rule = ConditionalRule::uniqueValues();
        } else {
            rule = ConditionalRule::formula(formulas.empty() ? std::string{} : formulas.front());
            rule.setFormulas(std::move(formulas));
        }
        const auto priority = internal::attribute(ruleTag, "priority");
        if (!priority.empty()) rule.setPriority(internal::parseIntegerExact<std::size_t>(priority, "conditional formatting priority"));
        rule.setStopIfTrue(internal::attribute(ruleTag, "stopIfTrue") == "1");
        const auto dxfId = internal::attribute(ruleTag, "dxfId");
        if (!dxfId.empty()) {
            const auto id = internal::parseIntegerExact<std::size_t>(dxfId, "dxfId");
            if (id < dxfStyles.size()) rule.setDifferentialStyle(dxfStyles[id]);
        }
        formatting.addRule(std::move(rule));
    }
}

for (auto& validationTag : internal::tags(xml, "dataValidation")) {
    const auto reference = internal::attribute(validationTag, "sqref");
    if (reference.empty()) continue;
    DataValidation validation(parseDataValidationType(internal::attribute(validationTag, "type")));
    validation.setReference(reference);
    validation.setOperator(parseDataValidationOperator(internal::attribute(validationTag, "operator")));
    validation.setErrorStyle(parseDataValidationErrorStyle(internal::attribute(validationTag, "errorStyle")));
    validation.setAllowBlank(internal::attribute(validationTag, "allowBlank") == "1");
    validation.setShowDropDown(internal::attribute(validationTag, "showDropDown") == "1");
    validation.setShowInputMessage(internal::attribute(validationTag, "showInputMessage") == "1");
    validation.setShowErrorMessage(internal::attribute(validationTag, "showErrorMessage") == "1");
    validation.setPromptTitle(internal::attribute(validationTag, "promptTitle"));
    validation.setPrompt(internal::attribute(validationTag, "prompt"));
    validation.setErrorTitle(internal::attribute(validationTag, "errorTitle"));
    validation.setError(internal::attribute(validationTag, "error"));
    validation.setFormula1(internal::tagText(validationTag, "formula1"));
    validation.setFormula2(internal::tagText(validationTag, "formula2"));
    ws.dataValidations().add(std::move(validation));
}

{
    const auto sheetRelationships = internal::loadPackageRelationships(z, target);

    // Comments are owner relationships, so their target must be an Internal
    // comments part that exists. Duplicate relationship IDs and invalid target
    // syntax have already been rejected by loadPackageRelationships().
    for (const auto& [id, relationship] : sheetRelationships) {
        if (!internal::relationshipTypeEndsWith(relationship.type, "/comments")) continue;
        const auto commentsTarget = internal::requireInternalPackageTarget(
            z, target, relationship, "worksheet comments");
        const auto& commentsText = z.get(commentsTarget);
        std::vector<std::string> authors;
        for (const auto& authorNode : internal::tags(commentsText, "author"))
            authors.push_back(internal::tagText(authorNode, "author"));
        for (const auto& commentNode : internal::tags(commentsText, "comment")) {
            const auto ref = internal::attribute(commentNode, "ref");
            if (ref.empty()) throw std::runtime_error("worksheet comment is missing its cell reference");
            std::string author;
            const auto authorIdText = internal::attribute(commentNode, "authorId");
            if (!authorIdText.empty()) {
                const auto authorId = internal::parseIntegerExact<std::size_t>(authorIdText, "comment authorId");
                if (authorId >= authors.size())
                    throw std::runtime_error("worksheet comment references an out-of-range authorId");
                author = authors[authorId];
            }
            std::string text;
            for (const auto& textNode : internal::tags(commentNode, "t"))
                text += internal::tagText(textNode, "t");
            Comment comment(std::move(text), author);
            comment.setVisible(internal::attribute(commentNode, "visible") == "1");
            ws.cell(ref).setComment(std::move(comment));
        }
    }

    // Legacy VML comment shapes carry the box size (and visibility hint).
    // Recover width/height so Comment.width()/height() round-trip. The VML
    // part is optional: some packages reference it without shipping the part.
    for (const auto& [id, relationship] : sheetRelationships) {
        if (!internal::relationshipTypeEndsWith(relationship.type, "/vmlDrawing")) continue;
        std::string vmlTarget;
        try {
            vmlTarget = internal::requireInternalPackageTarget(z, target, relationship, "worksheet vmlDrawing");
        } catch (const std::exception&) {
            continue;
        }
        if (!z.contains(vmlTarget)) continue;
        const auto& vmlText = z.get(vmlTarget);
        for (const auto& shapeNode : internal::tags(vmlText, "v:shape")) {
            const auto style = internal::attribute(shapeNode, "style");
            double widthPx = 0.0, heightPx = 0.0;
            auto findMeasure = [](const std::string& text, const char* key) -> double {
                const auto keyPos = text.find(key);
                if (keyPos == std::string::npos) return 0.0;
                const auto start = keyPos + std::strlen(key);
                const auto end = text.find_first_of("pxpt;", start);
                if (end == std::string::npos) return 0.0;
                double value = 0.0;
                if (!internal::tryParseDoubleExact(text.substr(start, end - start), value)) return 0.0;
                return value;
            };
            widthPx = findMeasure(style, "width:");
            heightPx = findMeasure(style, "height:");
            if (widthPx <= 0.0 || heightPx <= 0.0) continue;
            const auto rows = internal::tags(shapeNode, "x:Row");
            const auto cols = internal::tags(shapeNode, "x:Column");
            if (rows.empty() || cols.empty()) continue;
            const auto row = internal::parseIntegerExact<std::size_t>(internal::tagText(rows.front(), "x:Row"), "vml comment Row") + 1;
            const auto col = internal::parseIntegerExact<std::size_t>(internal::tagText(cols.front(), "x:Column"), "vml comment Column") + 1;
            auto& commentCell = ws.cell(row, col);
            if (commentCell.hasComment()) {
                commentCell.comment().setWidth(widthPx * 0.75);
                commentCell.comment().setHeight(heightPx * 0.75);
            }
        }
        (void)id;
        break;
    }

    for (const auto& part : internal::tags(xml, "tablePart")) {
        const auto relationshipId = internal::attribute(part, "r:id");
        const auto& relationship = internal::requirePackageRelationship(
            sheetRelationships, relationshipId, "/table", "worksheet tablePart");
        const auto tableTarget = internal::requireInternalPackageTarget(
            z, target, relationship, "worksheet tablePart");
        const auto& tableText = z.get(tableTarget);
        const auto tableNodes = internal::tags(tableText, "table");
        if (tableNodes.empty())
            throw std::runtime_error("worksheet table relationship target has no table element");
        const auto& tableNode = tableNodes.front();
        const auto tableName = internal::attribute(tableNode, "name");
        const auto tableRef = internal::attribute(tableNode, "ref");
        if (tableName.empty() || tableRef.empty())
            throw std::runtime_error("worksheet table is missing required name/ref attributes");
        auto& table = ws.addTable(tableName, tableRef);
        const auto displayName = internal::attribute(tableNode, "displayName");
        if(!displayName.empty()) table.setDisplayName(displayName);
        table.setShowHeaderRow(internal::attribute(tableNode,"headerRowCount") != "0");
        table.setShowTotalsRow(internal::attribute(tableNode,"totalsRowShown") == "1");
        for (const auto& columnNode : internal::tags(tableText,"tableColumn")) {
            auto& column = table.addColumn(internal::attribute(columnNode,"name"));
            const auto totalsFunction = internal::attribute(columnNode,"totalsRowFunction");
            if (!totalsFunction.empty()) column.setTotalsRowFunction(totalsFunction);
            const auto totalsLabel = internal::attribute(columnNode,"totalsRowLabel");
            if (!totalsLabel.empty()) column.setTotalsRowLabel(totalsLabel);
            const auto formulaNodes = internal::tags(columnNode,"calculatedColumnFormula");
            if (!formulaNodes.empty()) column.setTotalsRowFormula(internal::tagText(formulaNodes.front(),"calculatedColumnFormula"));
        }
        const auto styleNodes = internal::tags(tableText,"tableStyleInfo");
        if(!styleNodes.empty()) {
            const auto& style=styleNodes.front();
            table.styleInfo().setName(internal::attribute(style,"name"));
            table.styleInfo().setShowFirstColumn(internal::attribute(style,"showFirstColumn")=="1");
            table.styleInfo().setShowLastColumn(internal::attribute(style,"showLastColumn")=="1");
            table.styleInfo().setShowRowStripes(internal::attribute(style,"showRowStripes")!="0");
            table.styleInfo().setShowColumnStripes(internal::attribute(style,"showColumnStripes")=="1");
        }
    }

    // External hyperlinks use an r:id; location-only hyperlinks are internal
    // workbook locations and intentionally have no relationship record.
    for (const auto& linkNode : internal::tags(xml, "hyperlink")) {
        const auto ref = internal::attribute(linkNode, "ref");
        if (ref.empty()) throw std::runtime_error("worksheet hyperlink is missing its cell reference");
        Hyperlink link;
        const auto relationshipId = internal::attribute(linkNode, "r:id");
        if (!relationshipId.empty()) {
            const auto& relationship = internal::requirePackageRelationship(
                sheetRelationships, relationshipId, "/hyperlink", "worksheet hyperlink");
            if (relationship.targetMode == "External") {
                link.setTarget(relationship.target);
                link.setExternal(true);
            } else {
                link.setTarget(internal::requireInternalPackageTarget(
                    z, target, relationship, "worksheet hyperlink"));
                link.setExternal(false);
            }
        } else {
            link.setTarget(internal::attribute(linkNode,"location"));
            link.setExternal(false);
        }
        link.setDisplay(internal::attribute(linkNode,"display"));
        link.setTooltip(internal::attribute(linkNode,"tooltip"));
        ws.cell(ref).setHyperlink(std::move(link));
    }
}

loadImages(ws, xml, z, target);
loadCharts(ws, xml, z, target);
loadPivotTables(ws, xml, z, target);

internal::tagsForEach(xml, "mergeCell", [&](std::string_view merge) {
    const auto ref = internal::attribute(merge, "ref");
    if (!ref.empty()) ws.mergeCells(ref);
});
// P1R: account for cells materialized by comments/hyperlinks before sheetData.
if (!ws.cells().empty()) {
    const auto preexisting = ws.cells().size();
    if (maxCells != 0 && (totalCells > maxCells || preexisting > maxCells - totalCells))
        throw std::runtime_error("Workbook exceeds configured maxCells limit");
    totalCells += preexisting;
}
internal::tagsForEach(xml, "c", [&](std::string_view c) {
    const auto a = internal::attribute(c, "r");
    if (a.empty()) throw std::runtime_error("worksheet cell is missing its r reference");
    const auto t = internal::attribute(c, "t");
    const auto beforeCellCount = ws.cells().size();
    auto& cell = ws.cellNoTrack(a);
    if (ws.cells().size() != beforeCellCount) {
        if (maxCells != 0 && totalCells >= maxCells)
            throw std::runtime_error("Workbook exceeds configured maxCells limit");
        ++totalCells;
    }
    const auto styleText = internal::attribute(c, "s");
    if (!styleText.empty()) {
        const auto styleId = internal::parseIntegerExact<std::size_t>(styleText, "cell style index");
        if (styleId >= styleCatalog.items.size())
            throw std::runtime_error("cell style index is outside cellXfs table");
        cell.setRawStyleIndex(styleId);
        cell.style() = styleCatalog.items[styleId];
    }
    // Most cells carry no formula; skip the element scan entirely when the
    // <f> marker is absent so hot-path cells stay allocation-free.
    std::vector<std::string> formulaTags;
    if (c.find("<f") != std::string_view::npos) formulaTags = internal::tags(c, "f");
    if (!formulaTags.empty()) {
        const auto& formulaTag = formulaTags.front();
        auto formulaText = internal::tagText(c, "f");
        cell.setFormula(formulaText);
        const auto formulaType = internal::attribute(formulaTag, "t");
        const auto reference = internal::attribute(formulaTag, "ref");
        const auto sharedIndex = internal::attribute(formulaTag, "si");
        const bool alwaysCalculateArray = internal::attribute(formulaTag, "aca") == "1";
        const bool calculateOnLoad = internal::attribute(formulaTag, "ca") == "1";

        FormulaType parsedType = FormulaType::Normal;
        // Detect Excel 365 dynamic array formulas: _xlfn. prefix + aca="1".
        if (formulaText.rfind("_xlfn.", 0) == 0 && alwaysCalculateArray)
            parsedType = FormulaType::DynamicArray;
        else if (formulaType == "shared") parsedType = FormulaType::Shared;
        else if (formulaType == "array") parsedType = FormulaType::Array;
        else if (formulaType == "dataTable") parsedType = FormulaType::DataTable;

        // Keep ordinary formulas allocation-free. Only materialize FormulaMetadata
        // when the OOXML formula actually carries non-default metadata.
        if (parsedType != FormulaType::Normal || !reference.empty() || !sharedIndex.empty() ||
            alwaysCalculateArray || calculateOnLoad) {
            auto& metadata = cell.formulaMetadata();
            metadata.setType(parsedType);
            if (!reference.empty()) metadata.setReference(reference);
            if (!sharedIndex.empty()) metadata.setSharedIndex(internal::parseIntegerExact<unsigned>(sharedIndex, "shared formula index"));
            metadata.setAlwaysCalculateArray(alwaysCalculateArray);
            metadata.setCalculateOnLoad(calculateOnLoad);
        }
    }
    if (t == "inlineStr") {
        if (auto richText = parseRichTextRuns(c)) cell.setRichText(std::move(*richText));
        else cell.setValue(internal::tagText(c, "t"));
    }
    else {
        const auto v = internal::tagText(c, "v");
        if (t == "s" && !v.empty()) {
            const auto i = internal::parseIntegerExact<std::size_t>(v, "shared string index");
            if (i >= shared.size()) throw std::runtime_error("shared string index is outside sharedStrings table");
            if (shared[i].richText) cell.setRichText(*shared[i].richText);
            else cell.setValue(shared[i].plainText);
        }
        else if (t == "b") cell.setValue(v == "1");
        else if (t == "e") cell.setError(cellErrorFromString(v));
        else if (!v.empty()) {
            const auto number = internal::parseDoubleExact(v, "cell numeric value");
            if (xlpp::isDateFormatCode(cell.style().numberFormat(), cell.style().numFmtId()))
                cell.setValue(xlpp::fromExcelSerial(number, date1904));
            else
                cell.setValue(number);
        }
    }
    });

}
} // namespace

namespace xlpp {

}namespace xlpp {
NamedStyle& Workbook::addNamedStyle(NamedStyle style){if(style.name().empty())throw std::invalid_argument("Named style name cannot be empty");if(namedStyle(style.name()))throw std::invalid_argument("Named style already exists: "+style.name());namedStyles_.push_back(std::move(style));return namedStyles_.back();}
NamedStyle* Workbook::namedStyle(const std::string& name) noexcept{for(auto& style:namedStyles_)if(style.name()==name)return &style;return nullptr;}
const NamedStyle* Workbook::namedStyle(const std::string& name) const noexcept{for(const auto& style:namedStyles_)if(style.name()==name)return &style;return nullptr;}
void Workbook::applyNamedStyle(Cell& cell,const std::string& name) const{const auto* style=namedStyle(name);if(!style)throw std::out_of_range("Unknown named style: "+name);cell.style()=style->style();cell.setNamedStyle(name);}
DefinedName& Workbook::addDefinedName(DefinedName name) {
    if (name.localSheetId().has_value() && *name.localSheetId() >= sheets_.size())
        throw std::out_of_range("Defined name localSheetId is outside the workbook worksheet range");
    const auto duplicate = std::find_if(definedNames_.begin(), definedNames_.end(), [&](const DefinedName& existing) {
        return existing.localSheetId() == name.localSheetId() &&
               internal::worksheetNamesEquivalent(existing.name(), name.name());
    });
    if (duplicate != definedNames_.end())
        throw std::invalid_argument("Defined name already exists in the same scope: " + name.name());
    definedNames_.push_back(std::move(name));
    return definedNames_.back();
}
DefinedName* Workbook::definedName(const std::string& name) noexcept {
    for (auto& item : definedNames_)
        if (internal::worksheetNamesEquivalent(item.name(), name)) return &item;
    return nullptr;
}
const DefinedName* Workbook::definedName(const std::string& name) const noexcept {
    for (const auto& item : definedNames_)
        if (internal::worksheetNamesEquivalent(item.name(), name)) return &item;
    return nullptr;
}
Worksheet& Workbook::addWorksheet(std::string name) {
    internal::validateWorksheetName(name);
    const bool duplicateWorksheet = std::any_of(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
    const bool duplicateChartsheet = std::any_of(chartsheets_.begin(), chartsheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
    if (duplicateWorksheet || duplicateChartsheet) throw std::invalid_argument("Duplicate workbook sheet name");
    sheets_.emplace_back(std::move(name));
    sheetOrder_.push_back({WorkbookSheetKind::Worksheet, sheets_.size() - 1});
    if (generatedVbaProject_) ensureWorksheetVbaCodeNames(sheets_);
    return sheets_.back();
}
Slicer& Workbook::addSlicer(const std::string& worksheetName, Slicer slicer) {
    if (slicer.name.empty()) throw std::invalid_argument("Slicer name cannot be empty");
    if (worksheet(worksheetName) == nullptr) throw std::invalid_argument("Slicer worksheet does not exist: " + worksheetName);
    slicer.worksheetName = worksheetName;
    slicers_.push_back(std::move(slicer));
    return slicers_.back();
}
Worksheet* Workbook::worksheet(const std::string& n)noexcept{auto i=std::find_if(sheets_.begin(),sheets_.end(),[&](auto&s){return internal::worksheetNamesEquivalent(s.name(),n);});return i==sheets_.end()?nullptr:&*i;}const Worksheet* Workbook::worksheet(const std::string& n)const noexcept{auto i=std::find_if(sheets_.begin(),sheets_.end(),[&](auto&s){return internal::worksheetNamesEquivalent(s.name(),n);});return i==sheets_.end()?nullptr:&*i;}
Worksheet& Workbook::operator[](std::size_t index){return sheets_.at(index);}
const Worksheet& Workbook::operator[](std::size_t index) const{return sheets_.at(index);}
std::size_t Workbook::index(const Worksheet& sheet) const{const auto it=std::find_if(sheets_.begin(),sheets_.end(),[&](auto&s){return &s==&sheet;});if(it==sheets_.end())throw std::out_of_range("Worksheet not in this workbook");return static_cast<std::size_t>(std::distance(sheets_.begin(),it));}
std::vector<std::string> Workbook::sheetNames() const{std::vector<std::string> names;names.reserve(sheets_.size());for(const auto& sheet:sheets_)names.push_back(sheet.name());return names;}
Worksheet& Workbook::copyWorksheet(const Worksheet& source, std::string newName) {
    internal::validateWorksheetName(newName);
    const bool duplicateWorksheet = std::any_of(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), newName); });
    const bool duplicateChartsheet = std::any_of(chartsheets_.begin(), chartsheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), newName); });
    if (duplicateWorksheet || duplicateChartsheet) throw std::invalid_argument("Duplicate workbook sheet name");
    Worksheet copy = source; copy.rename(std::move(newName)); copy.setVbaCodeName({});
    sheets_.push_back(std::move(copy));
    sheetOrder_.push_back({WorkbookSheetKind::Worksheet, sheets_.size() - 1});
    if (generatedVbaProject_) ensureWorksheetVbaCodeNames(sheets_);
    return sheets_.back();
}

Chartsheet& Workbook::addChartsheet(std::string name, Chart chart) {
    internal::validateWorksheetName(name);
    const bool duplicateWorksheet = std::any_of(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
    const bool duplicateChartsheet = std::any_of(chartsheets_.begin(), chartsheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
    if (duplicateWorksheet || duplicateChartsheet) throw std::invalid_argument("Duplicate workbook sheet name");
    chartsheets_.emplace_back(std::move(name), std::move(chart));
    sheetOrder_.push_back({WorkbookSheetKind::Chartsheet, chartsheets_.size() - 1});
    return chartsheets_.back();
}
Chartsheet* Workbook::chartsheet(const std::string& name) noexcept {
    const auto it = std::find_if(chartsheets_.begin(), chartsheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
    return it == chartsheets_.end() ? nullptr : &*it;
}
const Chartsheet* Workbook::chartsheet(const std::string& name) const noexcept {
    const auto it = std::find_if(chartsheets_.begin(), chartsheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
    return it == chartsheets_.end() ? nullptr : &*it;
}
bool Workbook::renameChartsheet(const std::string& oldName, std::string newName) {
    internal::validateWorksheetName(newName);
    auto* target = chartsheet(oldName);
    if (!target) return false;
    if (internal::worksheetNamesEquivalent(target->name(), newName) && target->name() == newName) return true;
    const bool duplicateWorksheet = std::any_of(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), newName); });
    const bool duplicateChartsheet = std::any_of(chartsheets_.begin(), chartsheets_.end(), [&](const auto& sheet) { return &sheet != target && internal::worksheetNamesEquivalent(sheet.name(), newName); });
    if (duplicateWorksheet || duplicateChartsheet) throw std::invalid_argument("Duplicate workbook sheet name");
    target->rename(std::move(newName));
    return true;
}
bool Workbook::removeChartsheet(const std::string& name) {
    const auto it = std::find_if(chartsheets_.begin(), chartsheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
    if (it == chartsheets_.end()) return false;
    if (sheetOrder_.size() == 1) throw std::logic_error("Cannot remove the last sheet from a workbook");
    const auto removedIndex = static_cast<std::size_t>(std::distance(chartsheets_.begin(), it));
    const auto orderIt = std::find_if(sheetOrder_.begin(), sheetOrder_.end(), [&](const auto& entry) { return entry.kind == WorkbookSheetKind::Chartsheet && entry.kindIndex == removedIndex; });
    const auto workbookIndex = orderIt == sheetOrder_.end() ? sheetOrder_.size() : static_cast<std::size_t>(std::distance(sheetOrder_.begin(), orderIt));
    if (it->imported() && !it->sourcePart().empty()) {
        std::set<std::string> removedParts;
        suppressExclusivePartClosure(it->sourcePart(), preservedRelationships_, removedParts);
        preservedParts_.erase(std::remove_if(preservedParts_.begin(), preservedParts_.end(), [&](const PreservedPart& part) {
            return removedParts.contains(part.name);
        }), preservedParts_.end());
        preservedRelationships_.erase(std::remove_if(preservedRelationships_.begin(), preservedRelationships_.end(), [&](const PreservedRelationship& relationship) {
            if (removedParts.contains(relationship.sourcePart)) return true;
            if (relationship.targetMode == "External") return false;
            const auto target = resolvePackagePart(relationship.sourcePart, relationship.target);
            return removedParts.contains(target);
        }), preservedRelationships_.end());
    }
    // Remove locally-scoped defined names owned by this workbook tab and shift later scopes.
    definedNames_.erase(std::remove_if(definedNames_.begin(), definedNames_.end(), [&](DefinedName& defined) {
        if (defined.localSheetId() && *defined.localSheetId() == workbookIndex) return true;
        if (defined.localSheetId() && *defined.localSheetId() > workbookIndex) defined.setLocalSheetId(*defined.localSheetId() - 1);
        return false;
    }), definedNames_.end());
    const auto oldActive = activeWorkbookSheetIndex_;
    if (orderIt != sheetOrder_.end()) sheetOrder_.erase(orderIt);
    for (auto& entry : sheetOrder_) if (entry.kind == WorkbookSheetKind::Chartsheet && entry.kindIndex > removedIndex) --entry.kindIndex;
    chartsheets_.erase(it);
    if (!sheetOrder_.empty()) {
        if (oldActive > workbookIndex) activeWorkbookSheetIndex_ = oldActive - 1;
        else if (oldActive == workbookIndex) activeWorkbookSheetIndex_ = std::min(workbookIndex, sheetOrder_.size() - 1);
        if (sheetOrder_[activeWorkbookSheetIndex_].visibility != WorkbookSheetVisibility::Visible)
            activeWorkbookSheetIndex_ = firstVisibleWorkbookSheetIndex();
        firstVisibleWorkbookSheetIndex_ = firstVisibleWorkbookSheetIndex();
    }
    return true;
}
std::vector<std::string> Workbook::workbookSheetNames() const {
    std::vector<std::string> result; result.reserve(sheetOrder_.size());
    for (const auto& entry : sheetOrder_) result.push_back(entry.kind == WorkbookSheetKind::Worksheet ? sheets_.at(entry.kindIndex).name() : chartsheets_.at(entry.kindIndex).name());
    return result;
}
std::vector<WorkbookSheetDescriptor> Workbook::workbookSheets() const {
    std::vector<WorkbookSheetDescriptor> result; result.reserve(sheetOrder_.size());
    for (std::size_t i = 0; i < sheetOrder_.size(); ++i) {
        const auto& entry = sheetOrder_[i];
        result.push_back({entry.kind, entry.kindIndex,
            entry.kind == WorkbookSheetKind::Worksheet ? sheets_.at(entry.kindIndex).name() : chartsheets_.at(entry.kindIndex).name(),
            entry.visibility, i == activeWorkbookSheetIndex_});
    }
    return result;
}
WorkbookSheetVisibility Workbook::workbookSheetVisibility(std::size_t index) const {
    if (index >= sheetOrder_.size()) throw std::out_of_range("Workbook sheet index is out of range");
    return sheetOrder_[index].visibility;
}
std::size_t Workbook::firstVisibleWorkbookSheetIndex() const noexcept {
    for (std::size_t i = 0; i < sheetOrder_.size(); ++i)
        if (sheetOrder_[i].visibility == WorkbookSheetVisibility::Visible) return i;
    return sheetOrder_.size();
}
void Workbook::setWorkbookSheetVisibility(std::size_t index, WorkbookSheetVisibility visibility) {
    if (index >= sheetOrder_.size()) throw std::out_of_range("Workbook sheet index is out of range");
    if (sheetOrder_[index].visibility == visibility) return;
    if (visibility != WorkbookSheetVisibility::Visible) {
        std::size_t visibleCount = 0;
        for (const auto& entry : sheetOrder_) if (entry.visibility == WorkbookSheetVisibility::Visible) ++visibleCount;
        if (sheetOrder_[index].visibility == WorkbookSheetVisibility::Visible && visibleCount <= 1)
            throw std::logic_error("Workbook must contain at least one visible sheet");
    }
    sheetOrder_[index].visibility = visibility;
    if (index == activeWorkbookSheetIndex_ && visibility != WorkbookSheetVisibility::Visible) {
        const auto replacement = firstVisibleWorkbookSheetIndex();
        if (replacement >= sheetOrder_.size()) throw std::logic_error("Workbook must contain a visible active sheet");
        activeWorkbookSheetIndex_ = replacement;
    }
    firstVisibleWorkbookSheetIndex_ = firstVisibleWorkbookSheetIndex();
}
void Workbook::setActiveWorkbookSheetIndex(std::size_t index) {
    if (index >= sheetOrder_.size()) throw std::out_of_range("Workbook active sheet index is out of range");
    if (sheetOrder_[index].visibility != WorkbookSheetVisibility::Visible)
        throw std::logic_error("A hidden or veryHidden sheet cannot be the active workbook sheet");
    activeWorkbookSheetIndex_ = index;
    firstVisibleWorkbookSheetIndex_ = firstVisibleWorkbookSheetIndex();
}
bool Workbook::setActiveWorkbookSheet(const std::string& name) {
    for (std::size_t i = 0; i < sheetOrder_.size(); ++i) {
        const auto& entry = sheetOrder_[i];
        const auto& candidate = entry.kind == WorkbookSheetKind::Worksheet ? sheets_.at(entry.kindIndex).name() : chartsheets_.at(entry.kindIndex).name();
        if (!internal::worksheetNamesEquivalent(candidate, name)) continue;
        setActiveWorkbookSheetIndex(i);
        return true;
    }
    return false;
}
void Workbook::moveWorkbookSheet(std::size_t fromIndex, std::size_t toIndex) {
    if (fromIndex >= sheetOrder_.size() || toIndex >= sheetOrder_.size()) throw std::out_of_range("Workbook sheet order index is out of range");
    if (fromIndex == toIndex) return;
    const auto oldActive = activeWorkbookSheetIndex_;
    auto entry = sheetOrder_[fromIndex];
    sheetOrder_.erase(sheetOrder_.begin() + static_cast<std::ptrdiff_t>(fromIndex));
    sheetOrder_.insert(sheetOrder_.begin() + static_cast<std::ptrdiff_t>(toIndex), entry);
    if (oldActive == fromIndex) activeWorkbookSheetIndex_ = toIndex;
    else if (fromIndex < oldActive && oldActive <= toIndex) activeWorkbookSheetIndex_ = oldActive - 1;
    else if (toIndex <= oldActive && oldActive < fromIndex) activeWorkbookSheetIndex_ = oldActive + 1;
    firstVisibleWorkbookSheetIndex_ = firstVisibleWorkbookSheetIndex();
}

void Workbook::save(const std::filesystem::path& p) const { save(p, SaveOptions{}); }
void Workbook::save(const std::filesystem::path& p, const SaveOptions& options) const {
    if (sheetOrder_.empty()) throw std::runtime_error("Workbook needs at least one worksheet or chartsheet");
    if (options.validateModelBeforeSave) {
        const auto validation = validateModelIntegrity();
        const bool rejectWarnings = options.rejectModelWarningsBeforeSave && validation.warningCount() != 0;
        if (!validation.ok() || rejectWarnings) {
            std::ostringstream message;
            message << "Workbook model validation failed with " << validation.errorCount() << " error(s)";
            if (options.rejectModelWarningsBeforeSave) message << " and " << validation.warningCount() << " warning(s)";
            std::size_t emitted = 0;
            for (const auto& issue : validation.issues) {
                if (issue.severity == ModelValidationSeverity::Warning && !options.rejectModelWarningsBeforeSave) continue;
                message << (emitted == 0 ? ": " : "; ") << issue.code;
                if (!issue.worksheetName.empty()) message << " [" << issue.worksheetName << "]";
                if (!issue.message.empty()) message << " " << issue.message;
                if (++emitted == 4) break;
            }
            throw std::runtime_error(message.str());
        }
    }
    const auto saveExtension = p.extension().string();
    if (saveExtension.size() == 4 && (saveExtension == ".xls" || saveExtension == ".XLS")) {
        std::vector<unsigned char> binary;
        internal::writeLegacyXls(*this, binary);
        internal::writeBinaryFile(p, binary);
        return;
    }
    if (saveExtension.size() == 5 && (saveExtension == ".xlsb" || saveExtension == ".XLSB")) {
        internal::writeXlsbPackage(*this, p, zlibLevel(options.compressionLevel));
        return;
    }
    const bool strict = options.strictNamespace;
    StyleCatalog styleCatalog;
    DxfCatalog dxfCatalog;
    std::size_t tableCount = 0;
    std::size_t commentCount = 0;
    internal::SstIndex sstIndex;
    // Node-based unordered_map storage keeps key addresses stable across
    // rehash, so the order vector can point at keys instead of owning a second
    // copy of every unique shared string.
    std::vector<const std::string*> sstStrings;
    std::size_t sstOccurrences = 0;
    std::vector<char> sheetHasComments(sheets_.size(), false);
    std::vector<char> sheetHasExternalLinks(sheets_.size(), false);

    constexpr auto noSourceSheet = std::numeric_limits<std::size_t>::max();
    std::vector<std::size_t> sourceSheetIndices(sheets_.size(), noSourceSheet);
    std::vector<bool> preserveDrawing(sheets_.size(), false);
    std::vector<bool> preservePivot(sheets_.size(), false);
    std::vector<std::size_t> generatedPivotStart(sheets_.size(), 0);
    std::vector<std::size_t> generatedPivotCount(sheets_.size(), 0);
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
        const auto sourceName = std::find(sourceSheetNames_.begin(), sourceSheetNames_.end(), sheets_[i].name());
        if (sourceName == sourceSheetNames_.end()) continue;
        const auto sourceIndex = static_cast<std::size_t>(std::distance(sourceSheetNames_.begin(), sourceName));
        if (sourceIndex >= sourceSheetXml_.size() || sourceIndex >= sourceSheetParts_.size()) continue;
        sourceSheetIndices[i] = sourceIndex;
        preserveDrawing[i] = !sheets_[i].drawingsDirty()
            && !internal::tags(sourceSheetXml_[sourceIndex], "drawing").empty();
        const auto sourceRelationships = relationshipsForSource(preservedRelationships_, sourceSheetParts_[sourceIndex]);
        const bool sourceHasPivot = !internal::tags(sourceSheetXml_[sourceIndex], "pivotTableParts").empty()
            || std::any_of(sourceRelationships.begin(), sourceRelationships.end(), [](const auto& relationship) {
                return relationshipKind(relationship) == "pivotTable";
            });
        preservePivot[i] = sourceHasPivot && !sheets_[i].pivotsDirty();
    }
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
        if (sheets_[i].pivotsDirty()) {
            generatedPivotStart[i] = 0;
            generatedPivotCount[i] = sheets_[i].pivotTables().size();
        } else if (sheets_[i].pivotAppendDirty()) {
            generatedPivotStart[i] = sheets_[i].loadedPivotCount();
            generatedPivotCount[i] = sheets_[i].appendedPivotCount();
        }
    }

    const auto firstDrawingId = nextAvailablePartId(preservedParts_, "xl/drawings/drawing", ".xml");
    const auto firstChartId = nextAvailablePartId(preservedParts_, "xl/charts/chart", ".xml");
    const auto firstChartsheetId = nextAvailablePartId(preservedParts_, "xl/chartsheets/sheet", ".xml");
    const auto firstPrinterSettingsId = nextAvailablePartId(preservedParts_, "xl/printerSettings/printerSettings", ".bin");
    const auto firstPivotId = nextAvailablePartId(preservedParts_, "xl/pivotTables/pivotTable", ".xml");
    const auto firstPivotCachePartId = nextAvailablePartId(preservedParts_, "xl/pivotCache/pivotCacheDefinition", ".xml");
    std::vector<std::size_t> generatedDrawingIds;
    std::vector<std::size_t> generatedChartIds;
    std::vector<std::size_t> generatedChartsheetIds;
    std::vector<std::size_t> generatedPrinterSettingsIds;
    std::vector<std::size_t> chartsheetPartIds(chartsheets_.size(), 0);
    std::vector<std::size_t> chartsheetDrawingIds(chartsheets_.size(), 0);
    std::vector<std::size_t> chartsheetChartIds(chartsheets_.size(), 0);
    std::vector<std::size_t> chartsheetPrinterSettingsIds(chartsheets_.size(), 0);
    std::vector<std::size_t> generatedPivotIds;
    std::vector<std::size_t> generatedPivotCachePartIds;
    std::unordered_map<std::size_t, std::size_t> generatedPivotCacheIds;
    std::unordered_map<std::size_t, std::size_t> generatedPivotCachePartIdByPivot;
    std::unordered_map<std::size_t, xlpp::PivotTable> generatedPivotModels;
    std::unordered_map<std::size_t, xlpp::PivotTable> generatedPivotCacheModels;
    std::unordered_map<std::size_t, std::size_t> generatedCacheLogicalIds;
    std::unordered_map<std::string, std::size_t> sharedGeneratedCacheParts;
    std::size_t nextDrawingId = firstDrawingId;
    std::size_t nextChartId = firstChartId;
    std::size_t nextChartsheetId = firstChartsheetId;
    std::size_t nextPrinterSettingsId = firstPrinterSettingsId;
    std::size_t nextPivotId = firstPivotId;
    std::size_t nextPivotCachePartId = firstPivotCachePartId;
    std::size_t nextPivotCacheId = nextAvailablePivotCacheId(sourceWorkbookXml_);

    std::size_t totalModelCells = 0;
    for (const auto& s : sheets_) totalModelCells += s.cells().size();
    sstIndex.reserve(totalModelCells);

    for (std::size_t sheetIndex = 0; sheetIndex < sheets_.size(); ++sheetIndex) {
        const auto& sheet = sheets_[sheetIndex];
        tableCount += sheet.tables().size();
        if (!preserveDrawing[sheetIndex] && (!sheet.images().empty() || !sheet.shapes().empty() || sheet.chartCount() > 0)) {
            generatedDrawingIds.push_back(nextDrawingId++);
            for (std::size_t chartIndex = 0; chartIndex < sheet.chartCount(); ++chartIndex)
                generatedChartIds.push_back(nextChartId++);
        } else if (preserveDrawing[sheetIndex] && sheet.appendedChartCount() > 0) {
            for (std::size_t chartIndex = 0; chartIndex < sheet.appendedChartCount(); ++chartIndex)
                generatedChartIds.push_back(nextChartId++);
        }
        for (std::size_t pivotIndex = 0; pivotIndex < generatedPivotCount[sheetIndex]; ++pivotIndex) {
            const auto pivotPartId = nextPivotId++;
            generatedPivotIds.push_back(pivotPartId);
            const auto modelIndex = generatedPivotStart[sheetIndex] + pivotIndex;
            auto effective = effectivePivotTable(sheet.pivotTables()[modelIndex], sheets_, sheet, 1);
            const auto sharedKey = effective.cache().sharedCacheKey();

            std::size_t cachePartId = 0;
            std::size_t cacheId = 0;
            if (!sharedKey.empty()) {
                const auto shared = sharedGeneratedCacheParts.find(sharedKey);
                if (shared != sharedGeneratedCacheParts.end()) {
                    cachePartId = shared->second;
                    const auto existing = generatedPivotCacheModels.find(cachePartId);
                    if (existing == generatedPivotCacheModels.end()
                        || !pivotCachesEquivalent(existing->second.cache(), effective.cache()))
                        throw std::invalid_argument("PivotTables sharing cache key '" + sharedKey + "' have incompatible cache definitions");
                    cacheId = generatedCacheLogicalIds.at(cachePartId);
                }
            }
            if (cachePartId == 0) {
                cachePartId = nextPivotCachePartId++;
                cacheId = nextPivotCacheId++;
                generatedPivotCachePartIds.push_back(cachePartId);
                generatedCacheLogicalIds.emplace(cachePartId, cacheId);
                if (!sharedKey.empty()) sharedGeneratedCacheParts.emplace(sharedKey, cachePartId);
            }
            effective.cache().setCacheId(static_cast<int>(cacheId));
            generatedPivotCacheIds.emplace(pivotPartId, cacheId);
            generatedPivotCachePartIdByPivot.emplace(pivotPartId, cachePartId);
            generatedPivotModels.emplace(pivotPartId, effective);
            generatedPivotCacheModels.try_emplace(cachePartId, effective);
        }
        for (const auto& entry : sheet.cells()) {
            const auto& cell = entry.second;
            if (cell.hasNonDefaultStyle()) styleCatalog.id(cell.style());
            if (cell.hasComment()) sheetHasComments[sheetIndex] = true;
            if (cell.hasHyperlink() && cell.hyperlinkValue()->external())
                sheetHasExternalLinks[sheetIndex] = true;
            if (const auto* sv = std::get_if<std::string>(&cell.value())) {
                ++sstOccurrences;
                const auto [it, inserted] = sstIndex.try_emplace(*sv, sstStrings.size());
                if (inserted) sstStrings.push_back(&it->first);
            }
        }
        if (sheetHasComments[sheetIndex]) ++commentCount;
        for (const auto& formatting : sheet.conditionalFormatting().entries())
            for (const auto& rule : formatting.rules()) if (rule.hasDifferentialStyle()) dxfCatalog.id(rule.differentialStyle());
    }
    for (std::size_t chartsheetIndex = 0; chartsheetIndex < chartsheets_.size(); ++chartsheetIndex) {
        const auto& chartSheet = chartsheets_[chartsheetIndex];
        const bool regenerateChartClosure = !chartSheet.imported() || chartSheet.chartDirty();
        if (regenerateChartClosure) {
            if (!chartSheet.hasChart())
                throw std::invalid_argument("Generated chartsheet must contain a chart: " + chartSheet.name());
            chartsheetPartIds[chartsheetIndex] = nextChartsheetId++;
            chartsheetDrawingIds[chartsheetIndex] = nextDrawingId++;
            chartsheetChartIds[chartsheetIndex] = nextChartId++;
            generatedChartsheetIds.push_back(chartsheetPartIds[chartsheetIndex]);
            generatedDrawingIds.push_back(chartsheetDrawingIds[chartsheetIndex]);
            generatedChartIds.push_back(chartsheetChartIds[chartsheetIndex]);
        }
        if (chartSheet.hasPrinterSettings() &&
            (!chartSheet.imported() || chartSheet.printerSettingsDirty_ || chartSheet.printerSettingsSourcePart_.empty())) {
            chartsheetPrinterSettingsIds[chartsheetIndex] = nextPrinterSettingsId++;
            generatedPrinterSettingsIds.push_back(chartsheetPrinterSettingsIds[chartsheetIndex]);
        }
    }
    for (const auto& named : namedStyles_) styleCatalog.id(named.style());
    const bool macroEnabled = hasVbaProject();
    std::vector<std::string> vbaCodeNames;
    if (macroEnabled) vbaCodeNames = resolvedWorksheetVbaCodeNames(sheets_);
    auto sheetXmls = serializeSheets(sheets_, styleCatalog, dxfCatalog, date1904_, strict,
                                           macroEnabled ? &vbaCodeNames : nullptr,
                                           options.parallelSheets ? options.parallelWorkers : 0,
                                           options.parallelRows, &sstIndex, &cachedSheetXml_,
                                           cachedSheetXmlStrict_, cachedSheetXmlDate1904_);
    for (auto& sheet : sheets_) sheet.clearDirty();
    // Slicer parts are named slicerCacheK.xml / slicerK.xml with 1-based ids.
    std::vector<std::size_t> slicerCacheIds, slicerIds;
    for (std::size_t si = 0; si < slicers_.size(); ++si) {
        slicerCacheIds.push_back(si + 1);
        slicerIds.push_back(si + 1);
    }
    internal::ZipArchive z;
    std::set<std::string> suppressedPreservedParts;
    z.setCompressionLevel(zlibLevel(options.compressionLevel));
    z.setCompressionStrategy(zlibStrategy(options.compressionStrategy));
    z.setParallelWorkers(options.parallelWorkers);
    z.add("[Content_Types].xml", contentTypes(sheets_.size(), generatedChartsheetIds, generatedPrinterSettingsIds, tableCount, commentCount,
        generatedDrawingIds, preservedParts_, strict, !sstStrings.empty(), generatedChartIds,
        generatedPivotIds, generatedPivotCachePartIds, !customProps_.empty(), macroEnabled, template_,
        slicerCacheIds, slicerIds));
    const auto rootOriginalRelationships = relationshipsForSource(preservedRelationships_, {});
    const auto mergedRootRelationships = mergeRelationshipsXml(rootrels(strict, !customProps_.empty()), rootOriginalRelationships,
        [](const PreservedRelationship& relationship) {
            const auto kind = relationshipKind(relationship);
            return kind != "officeDocument" && kind != "core-properties"
                && kind != "extended-properties" && kind != "custom-properties";
        }, strict, nullptr);
    z.add("_rels/.rels", mergedRootRelationships);
    z.add("docProps/core.xml", corePropertiesXml(properties_, strict));
    z.add("docProps/app.xml", appPropertiesXml(strict));
    if (!customProps_.empty()) z.add("docProps/custom.xml", customPropertiesXml(customProps_));
    for (const auto cachePartId : generatedPivotCachePartIds) {
        const auto& effectivePivot = generatedPivotCacheModels.at(cachePartId);
        z.add("xl/pivotCache/pivotCacheDefinition" + std::to_string(cachePartId) + ".xml", pivotCacheXml(effectivePivot, strict));
        z.add("xl/pivotCache/pivotCacheRecords" + std::to_string(cachePartId) + ".xml", pivotCacheRecordsXml(effectivePivot, strict));
        z.add("xl/pivotCache/_rels/pivotCacheDefinition" + std::to_string(cachePartId) + ".xml.rels",
              "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + nsRelsPkg(strict) + "\"><Relationship Id=\"rId1\" Type=\"" + nsRelsDoc(strict) + "/pivotCacheRecords\" Target=\"pivotCacheRecords" + std::to_string(cachePartId) + ".xml\"/></Relationships>");
    }
    z.add("xl/styles.xml", stylesXml(styleCatalog, namedStyles_, dxfCatalog, strict));
    if (!sstStrings.empty()) {
        std::ostringstream sstXml;
        sstXml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"" << nsMain(strict)
               << "\" count=\"" << sstOccurrences << "\" uniqueCount=\"" << sstStrings.size() << "\">";
        for (const auto* text : sstStrings) {
            sstXml << "<si><t xml:space=\"preserve\">"; writeXmlEscaped(sstXml, *text); sstXml << "</t></si>";
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
    const auto firstVisibleSheet = firstVisibleWorkbookSheetIndex();
    if (firstVisibleSheet >= sheetOrder_.size()) throw std::logic_error("Workbook must contain at least one visible sheet");
    const auto effectiveActiveSheet = activeWorkbookSheetIndex_ < sheetOrder_.size()
        && sheetOrder_[activeWorkbookSheetIndex_].visibility == WorkbookSheetVisibility::Visible
        ? activeWorkbookSheetIndex_ : firstVisibleSheet;
    wb << workbookViewsXml(sourceWorkbookXml_, effectiveActiveSheet, firstVisibleSheet);
    wb << "<sheets>";
    rels << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << nsRelsPkg(strict) << "\">";
    std::size_t globalTableId = 1;
    std::size_t globalCommentId = 1;
    std::size_t globalDrawingId = firstDrawingId;
    std::size_t globalMediaId = nextAvailableMediaId(preservedParts_);
    std::size_t globalChartId = firstChartId;
    std::size_t globalPivotId = firstPivotId;
    for (std::size_t orderIndex = 0; orderIndex < sheetOrder_.size(); ++orderIndex) {
        const auto& entry = sheetOrder_[orderIndex];
        const auto relationshipId = orderIndex + 1;
        if (entry.kind == WorkbookSheetKind::Worksheet) {
            const auto& sheet = sheets_.at(entry.kindIndex);
            wb << "<sheet name=\"" << xmlEscape(sheet.name()) << "\" sheetId=\"" << relationshipId
               << "\" r:id=\"rId" << relationshipId << "\"";
            if (entry.visibility == WorkbookSheetVisibility::Hidden) wb << " state=\"hidden\"";
            else if (entry.visibility == WorkbookSheetVisibility::VeryHidden) wb << " state=\"veryHidden\"";
            wb << "/>";
            rels << "<Relationship Id=\"rId" << relationshipId << "\" Type=\"" << nsRelsDoc(strict)
                 << "/worksheet\" Target=\"worksheets/sheet" << entry.kindIndex + 1 << ".xml\"/>";
        } else {
            const auto& chartSheet = chartsheets_.at(entry.kindIndex);
            wb << "<sheet name=\"" << xmlEscape(chartSheet.name()) << "\" sheetId=\"" << relationshipId
               << "\" r:id=\"rId" << relationshipId << "\"";
            if (entry.visibility == WorkbookSheetVisibility::Hidden) wb << " state=\"hidden\"";
            else if (entry.visibility == WorkbookSheetVisibility::VeryHidden) wb << " state=\"veryHidden\"";
            wb << "/>";
            std::string target;
            if (chartSheet.imported() && !chartSheet.chartDirty() && !chartSheet.sourcePart().empty()) {
                target = chartSheet.sourcePart();
                if (target.rfind("xl/", 0) == 0) target.erase(0, 3);
                else if (!target.empty() && target.front() != '/') target.insert(target.begin(), '/');
            } else {
                const auto partId = chartsheetPartIds.at(entry.kindIndex);
                if (partId == 0) throw std::logic_error("Generated chartsheet part ID was not assigned");
                target = "chartsheets/sheet" + std::to_string(partId) + ".xml";
            }
            rels << "<Relationship Id=\"rId" << relationshipId << "\" Type=\"" << nsRelsDoc(strict)
                 << "/chartsheet\" Target=\"" << xmlEscape(target) << "\"/>";
        }
    }
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
        const auto& sheet = sheets_[i];
        const bool hasLinks = sheetHasExternalLinks[i] != 0;
        const bool hasComments = sheetHasComments[i] != 0;
        const bool hasGeneratedImages = !preserveDrawing[i] && !sheet.images().empty();
        const bool hasSheetShapes = !preserveDrawing[i] && !sheet.shapes().empty();
        const bool hasSheetCharts = !preserveDrawing[i] && sheet.chartCount() > 0;
        const bool hasSheetPivots = generatedPivotCount[i] > 0;
        const auto sourceSheetIndex = sourceSheetIndices[i];
        const std::string originalSheetPart = sourceSheetIndex != noSourceSheet
            ? sourceSheetParts_[sourceSheetIndex] : std::string{};
        const std::string originalSheetXml = sourceSheetIndex != noSourceSheet
            ? sourceSheetXml_[sourceSheetIndex] : std::string{};
        const auto originalSheetRelationships = relationshipsForSource(preservedRelationships_, originalSheetPart);
        // A mutable access to imported pivotTables() means the caller chose
        // regeneration through XL++'s object model. Retire the old pivot-table
        // root parts so they do not become orphaned when the worksheet owner
        // relationship is replaced. The original cache is kept if the workbook
        // still owns it; this is intentionally conservative for shared caches.
        if (sourceSheetIndex != noSourceSheet && !preservePivot[i] && generatedPivotCount[i] > 0) {
            for (const auto& relationship : originalSheetRelationships) {
                if (relationshipKind(relationship) != "pivotTable") continue;
                const auto originalPivotPart = internal::RelationshipGraph::resolveTarget(originalSheetPart, relationship.target);
                if (originalPivotPart.empty()) continue;
                suppressedPreservedParts.insert(originalPivotPart);
                suppressedPreservedParts.insert(internal::RelationshipGraph::relationshipsPartForSource(originalPivotPart));
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
        if (!sheet.tables().empty() || hasLinks || hasComments || hasGeneratedImages || hasSheetShapes || hasSheetCharts || hasSheetPivots || hasPreservedSheetRelationships) {
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
            if (hasGeneratedImages || hasSheetShapes || hasSheetCharts) {
                sheetRels << "<Relationship Id=\"rIdDrawing\" Type=\"" << nsRelsDoc(strict) << "/drawing\" Target=\"../drawings/drawing" << globalDrawingId << ".xml\"/>";
                const auto firstChartId = globalChartId;
                for (std::size_t ci = 0; ci < sheet.chartCount(); ++ci, ++globalChartId)
                    z.add("xl/charts/chart" + std::to_string(globalChartId) + ".xml", chartXml(sheet.chart(ci), strict));
                z.add("xl/drawings/drawing" + std::to_string(globalDrawingId) + ".xml", drawingXml(sheet, strict));
                z.add("xl/drawings/_rels/drawing" + std::to_string(globalDrawingId) + ".xml.rels", drawingRelationshipsXml(sheet, globalMediaId, firstChartId, strict));
                for (const auto& image : sheet.images()) {
                    const std::string bytes(reinterpret_cast<const char*>(image.bytes().data()), image.bytes().size());
                    z.add("xl/media/image" + std::to_string(globalMediaId++) + "." + image.extension(), bytes, false);
                }
                ++globalDrawingId;
            }
            if (preserveDrawing[i] && (sheet.appendedImageCount() > 0 ||
                !internal::WorkbookDrawingAccess::imageEdits(sheet).empty())) {
                if (!applyImageChangesToPreservedDrawing(z, sheet, originalSheetPart, originalSheetXml, preservedRelationships_, preservedParts_,
                                                         globalMediaId, suppressedPreservedParts))
                    throw std::runtime_error("Cannot apply selective image mutation to preserved drawing for worksheet: " + sheet.name());
            }
            if (preserveDrawing[i] && (sheet.appendedChartCount() > 0 ||
                !internal::WorkbookDrawingAccess::chartEdits(sheet).empty())) {
                if (!applyChartChangesToPreservedDrawing(z, sheet, originalSheetPart, originalSheetXml,
                                                         preservedRelationships_, preservedParts_, globalChartId,
                                                         suppressedPreservedParts))
                    throw std::runtime_error("Cannot apply selective chart mutation to preserved drawing for worksheet: " + sheet.name());
            }
            if (hasSheetPivots) {
                static_cast<void>(generatedPivotCacheModels);
                for (std::size_t generatedIndex = 0; generatedIndex < generatedPivotCount[i]; ++generatedIndex, ++globalPivotId) {
                    const auto pi = generatedPivotStart[i] + generatedIndex;
                    static_cast<void>(pi);
                    sheetRels << "<Relationship Id=\"rIdPivot" << (generatedIndex + 1) << "\" Type=\"" << nsRelsDoc(strict) << "/pivotTable\" Target=\"../pivotTables/pivotTable" << globalPivotId << ".xml\"/>";

                    const auto cacheId = generatedPivotCacheIds.at(globalPivotId);
                    const auto cachePartId = generatedPivotCachePartIdByPivot.at(globalPivotId);
                    const auto& effectivePivot = generatedPivotModels.at(globalPivotId);
                    z.add("xl/pivotTables/pivotTable" + std::to_string(globalPivotId) + ".xml", pivotTableXml(effectivePivot, cacheId, strict));
                    z.add("xl/pivotTables/_rels/pivotTable" + std::to_string(globalPivotId) + ".xml.rels",
                          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + nsRelsPkg(strict) + "\"><Relationship Id=\"rId1\" Type=\"" + nsRelsDoc(strict) + "/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition" + std::to_string(cachePartId) + ".xml\"/></Relationships>");
                }
            }
            sheetRels << "</Relationships>";
            auto generatedSheetXml = sheetXmls[i];
            auto mergedSheetRelationships = mergeRelationshipsXml(sheetRels.str(), originalSheetRelationships,
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
            for (std::size_t si = 0; si < slicers_.size(); ++si) {
                if (slicers_[si].worksheetName != sheet.name()) continue;
                const std::string slicerRel = "<Relationship Id=\"" + slicerRelationshipId(si) + "\" Type=\""
                    + nsRelsDoc(strict) + "/slicer\" Target=\"../slicers/slicer" + std::to_string(si + 1) + ".xml\"/>";
                const auto endTag = std::string("</Relationships>");
                const auto relEnd = mergedSheetRelationships.find(endTag);
                if (relEnd != std::string::npos) mergedSheetRelationships.insert(relEnd, slicerRel);
                z.add("xl/slicerCaches/slicerCache" + std::to_string(si + 1) + ".xml", slicerCacheXml(slicers_[si]));
                z.add("xl/slicers/slicer" + std::to_string(si + 1) + ".xml", slicerXml(slicers_[si]));
                generatedSheetXml = insertSlicerListExt(std::move(generatedSheetXml),
                                                        {slicerRelationshipId(si)}, strict);
            }
            z.add("xl/worksheets/sheet"+std::to_string(i+1)+".xml", generatedSheetXml);
            z.add("xl/worksheets/_rels/sheet"+std::to_string(i+1)+".xml.rels", mergedSheetRelationships);
        } else {
            auto generatedSheetXml = rebuildWorksheetTail(sheetXmls[i], originalSheetXml,
                                                          preserveDrawing[i], preservePivot[i], preserveTables, preserveComments);
            for (std::size_t si = 0; si < slicers_.size(); ++si) {
                if (slicers_[si].worksheetName != sheet.name()) continue;
                z.add("xl/slicerCaches/slicerCache" + std::to_string(si + 1) + ".xml", slicerCacheXml(slicers_[si]));
                z.add("xl/slicers/slicer" + std::to_string(si + 1) + ".xml", slicerXml(slicers_[si]));
                generatedSheetXml = insertSlicerListExt(std::move(generatedSheetXml),
                                                        {slicerRelationshipId(si)}, strict);
            }
            z.add("xl/worksheets/sheet"+std::to_string(i+1)+".xml", generatedSheetXml);
        }
    }
    internal::writeChartsheetPackageParts(
        z, chartsheets_, chartsheetPartIds, chartsheetDrawingIds, chartsheetChartIds,
        chartsheetPrinterSettingsIds, preservedRelationships_, suppressedPreservedParts, strict,
        [](const Chart& chart, bool strictMode) { return chartXml(chart, strictMode); });
    wb << "</sheets>";
    // Workbook-level slicer cache relationships.
    if (!slicers_.empty()) {
        for (std::size_t si = 0; si < slicers_.size(); ++si) {
            rels << "<Relationship Id=\"" << slicerCacheRelationshipId(si) << "\" Type=\"" << nsRelsDoc(strict)
                 << "/slicerCache\" Target=\"slicerCaches/slicerCache" << (si + 1) << ".xml\"/>";
        }
    }
    std::size_t nextWorkbookRelId = sheetOrder_.size() + 1;
    std::ostringstream generatedPivotCaches;
    if (!generatedPivotIds.empty()) {
        generatedPivotCaches << "<pivotCaches>";
        for (const auto cachePartId : generatedPivotCachePartIds) {
            const auto cacheId = generatedCacheLogicalIds.at(cachePartId);
            generatedPivotCaches << "<pivotCache cacheId=\"" << cacheId << "\" r:id=\"rId" << nextWorkbookRelId << "\"/>";
            rels << "<Relationship Id=\"rId" << nextWorkbookRelId++ << "\" Type=\"" << nsRelsDoc(strict) << "/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition" << cachePartId << ".xml\"/>";
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
            const auto orderIt = std::find_if(sheetOrder_.begin(), sheetOrder_.end(), [&](const auto& entry) {
                return entry.kind == WorkbookSheetKind::Worksheet && entry.kindIndex == sheetIndex;
            });
            if (orderIt == sheetOrder_.end()) throw std::logic_error("Worksheet is missing from workbook sheet order");
            const auto workbookSheetIndex = static_cast<std::size_t>(std::distance(sheetOrder_.begin(), orderIt));
            if (!sheet.printArea().empty()) {
                wb << "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"" << workbookSheetIndex << "\">"
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
            return kind != "worksheet" && kind != "chartsheet" && kind != "styles" && kind != "sharedStrings" && kind != "vbaProject";
        }, strict, &workbookXml);
    const auto rewrittenGeneratedPivotCaches = joinBlocks(extractTagBlocks(workbookXml, "pivotCaches"));
    eraseTagBlocks(workbookXml, "pivotCaches");
    workbookXml = preserveWorkbookNodes(std::move(workbookXml), sourceWorkbookXml_, false);
    insertBefore(workbookXml, "</workbook>", mergeWorkbookPivotCaches(sourceWorkbookXml_, rewrittenGeneratedPivotCaches));
    if (!slicers_.empty()) {
        std::vector<std::string> slicerCacheRels;
        for (std::size_t si = 0; si < slicers_.size(); ++si)
            slicerCacheRels.push_back(slicerCacheRelationshipId(si));
        workbookXml = insertWorkbookSlicerCachesExt(std::move(workbookXml), slicerCacheRels, strict);
    }
    z.add("xl/workbook.xml", workbookXml);
    z.add("xl/_rels/workbook.xml.rels", mergedWorkbookRelationships);

    // A source-generated VBA project contains one document module per worksheet.
    // Rebuild it at save time so worksheets added or removed after
    // setVbaModuleText() cannot leave vbaProject.bin out of sync with sheetPr
    // codeName attributes. Externally attached projects remain byte-for-byte
    // preserved because their document-module mapping is owned by the caller.
    std::string generatedVbaData;
    if (generatedVbaProject_) {
        auto modules = vbaModules();
        VbaProjectInfo info;
        try { info = vbaProjectInfo(); } catch (...) {}
        const auto bytes = internal::buildVbaProjectBinary(modules, vbaCodeNames, info);
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
    if (options.validatePackageBeforeWrite) {
        const auto packageValidation = internal::RelationshipGraph::fromArchive(z).validate();
        if (!packageValidation.ok()) {
            std::ostringstream message;
            message << "Workbook package validation failed before write";
            if (!packageValidation.relationshipSyntaxErrors.empty())
                message << ": relationship_syntax=" << packageValidation.relationshipSyntaxErrors.size();
            if (!packageValidation.duplicateRelationshipIds.empty())
                message << "; duplicate_relationship_ids=" << packageValidation.duplicateRelationshipIds.size();
            if (!packageValidation.danglingRelationships.empty())
                message << "; dangling_relationships=" << packageValidation.danglingRelationships.size();
            if (!packageValidation.orphanedParts.empty())
                message << "; orphaned_parts=" << packageValidation.orphanedParts.size();
            if (!packageValidation.contentTypeErrors.empty())
                message << "; content_type_errors=" << packageValidation.contentTypeErrors.size();
            if (!packageValidation.ownerReferenceErrors.empty())
                message << "; owner_reference_errors=" << packageValidation.ownerReferenceErrors.size();
            throw std::runtime_error(message.str());
        }
    }

    if (!options.encryption.enabled) {
        z.save(p);
    } else {
        // P1I: serialize the plaintext OOXML ZIP directly to memory. The
        // password-to-open path no longer materializes a temporary inner
        // .xlsx file containing the unencrypted workbook.
        auto packageBytes = z.saveToBytes();
        std::vector<unsigned char> encrypted;
        try {
            if (options.encryption.mode == PackageEncryptionMode::Agile) {
                internal::AgileEncryptionParameters parameters;
                parameters.spinCount = options.encryption.spinCount;
                parameters.keyBits = options.encryption.keyBits;
                parameters.hashAlgorithm = options.encryption.hashAlgorithm;
                encrypted = internal::encryptAgileOfficePackage(packageBytes, options.encryption.password, parameters);
            } else {
                internal::StandardEncryptionParameters parameters;
                parameters.keyBits = options.encryption.keyBits;
                encrypted = internal::encryptStandardOfficePackage(packageBytes, options.encryption.password, parameters);
            }
        } catch (...) {
            std::fill(packageBytes.begin(), packageBytes.end(), 0);
            throw;
        }
        std::fill(packageBytes.begin(), packageBytes.end(), 0);
        internal::writeBinaryFile(p, encrypted);
    }
}
void Workbook::load(const std::filesystem::path& p) { load(p, LoadOptions{}); }
void Workbook::load(const std::filesystem::path& p, const LoadOptions& options) {
    Workbook candidate;
    candidate.loadInPlace(p, options);
    *this = std::move(candidate);
}
void Workbook::loadInPlace(const std::filesystem::path& p, const LoadOptions& options) { clear(); diagnostics_ = LoadDiagnostics{};
    if (options.maxFileBytes != 0) {
        std::error_code ec;
        const auto outerSize = std::filesystem::file_size(p, ec);
        if (!ec && outerSize > options.maxFileBytes)
            throw std::runtime_error("Workbook exceeds configured maxFileBytes limit");
    }
    internal::ZipOpenLimits limits;
    limits.maxEntries = options.maxEntries;
    limits.maxEntryBytes = options.maxEntryBytes;
    limits.maxTotalBytes = options.maxTotalBytes;
    limits.maxFileBytes = options.maxFileBytes;
    limits.cancel = options.cancel;
    limits.progress = options.progress;
    internal::ZipArchive z;
    if (internal::isEncryptedOfficeCompoundFile(p)) {
        const auto encrypted = internal::readBinaryFile(p);
        internal::OfficeDecryptionLimits cryptoLimits;
        cryptoLimits.maxSpinCount = options.maxEncryptionSpinCount;
        cryptoLimits.maxPlainPackageBytes = options.maxDecryptedPackageBytes != 0
            ? options.maxDecryptedPackageBytes : options.maxFileBytes;
        cryptoLimits.allowStandardEncryption = options.allowStandardEncryption;
        cryptoLimits.requireAgileDataIntegrity = options.requireAgileDataIntegrity;
        cryptoLimits.maxEncryptionInfoBytes = options.maxEncryptionInfoBytes;
        auto package = internal::decryptOfficePackage(encrypted, options.passwordToOpen, cryptoLimits);
        if (options.maxFileBytes != 0 && package.size() > options.maxFileBytes)
            throw std::runtime_error("Decrypted workbook exceeds configured maxFileBytes limit");
        // P1I: parse the decrypted OOXML ZIP directly from memory. This avoids
        // writing plaintext workbook contents to a temporary .xlsx file.
        try {
            z = internal::ZipArchive::open(package, limits);
        } catch (...) {
            std::fill(package.begin(), package.end(), 0);
            throw;
        }
        std::fill(package.begin(), package.end(), 0);
    } else if (internal::isOle2CompoundFile(readFilePrefix(p, 512))) {
        // Classic binary .xls (BIFF8) workbook: parse the OLE2 compound file.
        if (options.maxFileBytes != 0) {
            std::error_code ec;
            const auto size = std::filesystem::file_size(p, ec);
            if (!ec && size > options.maxFileBytes)
                throw std::runtime_error("Workbook exceeds configured maxFileBytes limit");
        }
        const auto bytes = internal::readBinaryFile(p);
        internal::readLegacyXls(bytes, *this);
        return;
    } else {
        z = internal::ZipArchive::open(p, limits);
    }
    if (internal::isXlsbPackage(z)) {
        // Binary (BIFF12) workbook: read the basic sheet/cell model directly.
        internal::readXlsb(z, *this);
        return;
    }
    const auto relationshipGraph = internal::RelationshipGraph::fromArchive(z);
    preservedRelationships_ = relationshipGraph.relationships();
    const auto packageContentTypes = internal::loadContentTypeCatalog(z);
    if (const auto workbookType = packageContentTypes.overrides.find("/xl/workbook.xml");
        workbookType != packageContentTypes.overrides.end()) {
        const auto& type = workbookType->second;
        template_ = type == "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
            || type == "application/vnd.ms-excel.template.macroEnabled.main+xml";
    }
    const auto relationshipValidation = relationshipGraph.validate();
    for (const auto& issue : relationshipValidation.relationshipSyntaxErrors)
        diagnostics_.warnings.push_back("Relationship syntax issue: " + issue);
    for (const auto& issue : relationshipValidation.duplicateRelationshipIds)
        diagnostics_.warnings.push_back("Duplicate relationship: " + issue);
    for (const auto& issue : relationshipValidation.danglingRelationships)
        diagnostics_.warnings.push_back("Dangling relationship: " + issue);
    for (const auto& issue : relationshipValidation.orphanedParts)
        diagnostics_.warnings.push_back("Orphaned package part: " + issue);
    for (const auto& issue : relationshipValidation.contentTypeErrors)
        diagnostics_.warnings.push_back("Content-type issue: " + issue);
    for (const auto& issue : relationshipValidation.ownerReferenceErrors)
        diagnostics_.warnings.push_back("Broken owner reference: " + issue);
    if(z.contains("docProps/core.xml")){auto cp=z.get("docProps/core.xml");properties_.setTitle(internal::tagText(cp,"dc:title"));properties_.setSubject(internal::tagText(cp,"dc:subject"));properties_.setCreator(internal::tagText(cp,"dc:creator"));properties_.setDescription(internal::tagText(cp,"dc:description"));properties_.setKeywords(internal::tagText(cp,"cp:keywords"));properties_.setCategory(internal::tagText(cp,"cp:category"));properties_.setLastModifiedBy(internal::tagText(cp,"cp:lastModifiedBy"));}if(z.contains("docProps/custom.xml")){auto cust=z.get("docProps/custom.xml");for(const auto& p:internal::tags(cust,"property")){const auto n=internal::attribute(p,"name");if(n.empty())continue;const auto vtText=internal::tagText(p,"vt:lpwstr");if(!vtText.empty()){customProps_.add(CustomProperty(std::string(n),vtText));continue;}if(const auto i4Text=internal::tagText(p,"vt:i4");!i4Text.empty()){customProps_.add(CustomProperty(std::string(n),internal::parseIntegerExact<int>(i4Text, "custom property i4")));continue;}if(const auto r8Text=internal::tagText(p,"vt:r8");!r8Text.empty()){customProps_.add(CustomProperty(std::string(n),internal::parseDoubleExact(r8Text, "custom property r8")));continue;}if(const auto boolText=internal::tagText(p,"vt:bool");!boolText.empty()){customProps_.add(CustomProperty(std::string(n),boolText=="true"));continue;}customProps_.add(CustomProperty(std::string(n),vtText));}}StyleCatalog styleCatalog;std::vector<Style> dxfStyles;if(z.contains("xl/styles.xml")){const auto stylesText=z.get("xl/styles.xml");styleCatalog=parseStyleCatalog(stylesText);dxfStyles=parseDifferentialStyles(stylesText);for(const auto& node:internal::tags(stylesText,"cellStyle")){const auto name=internal::attribute(node,"name");if(name.empty()||name=="Normal")continue;const auto xf=internal::attribute(node,"xfId");if(xf.empty())continue;const auto id=internal::parseIntegerExact<std::size_t>(xf, "cellStyle xfId");if(id<styleCatalog.items.size())namedStyles_.emplace_back(name,styleCatalog.items[id]);}}std::vector<LoadedSharedString> shared;
if (z.contains("xl/sharedStrings.xml")) {
    const auto sstXml = z.get("xl/sharedStrings.xml");
    internal::tagsForEach(sstXml, "si", [&](std::string_view si) {
        if (options.maxSharedStrings != 0 && shared.size() >= options.maxSharedStrings)
            throw std::runtime_error("Workbook exceeds configured maxSharedStrings limit");
        LoadedSharedString item;
        item.richText = parseRichTextRuns(si);
        item.plainText = item.richText ? item.richText->plainText() : internal::tagText(si, "t");
        shared.push_back(std::move(item));
    });
}
auto wb=z.get("xl/workbook.xml");
sourceWorkbookXml_ = wb;
strictNamespaces_ = wb.find("http://purl.oclc.org/ooxml/spreadsheetml/main") != std::string::npos;
for(const auto& protectionNode:internal::tags(wb,"workbookProtection")){
    protection_.setLockStructure(internal::attribute(protectionNode,"lockStructure")=="1");
    protection_.setLockWindows(internal::attribute(protectionNode,"lockWindows")=="1");
    protection_.setLockRevision(internal::attribute(protectionNode,"lockRevision")=="1");
    protection_.setWorkbookPasswordHash(internal::attribute(protectionNode,"workbookPassword"));
}
for(const auto& prNode:internal::tags(wb,"workbookPr"))
    date1904_ = internal::attribute(prNode,"date1904")=="1";
for(const auto& calcNode:internal::tags(wb,"calcPr")){
    const auto cm = internal::attribute(calcNode,"calcMode"); if(!cm.empty()) calcProps_.setCalcMode(std::string(cm));
    const auto ci = internal::attribute(calcNode,"calcId"); if(!ci.empty()) calcProps_.setCalcId(internal::parseIntegerExact<int>(ci, "calcId"));
    calcProps_.setCalcOnSave(internal::attribute(calcNode,"calcOnSave") == "1");
    calcProps_.setFullCalcOnLoad(internal::attribute(calcNode,"fullCalcOnLoad") == "1");
    calcProps_.setFullPrecision(internal::attribute(calcNode,"fullPrecision") != "0");
    calcProps_.setIterate(internal::attribute(calcNode,"iterate") == "1");
    const auto ic = internal::attribute(calcNode,"iterateCount"); if(!ic.empty()) calcProps_.setIterateCount(internal::parseIntegerExact<int>(ic, "iterateCount"));
    const auto id = internal::attribute(calcNode,"iterateDelta"); if(!id.empty()) calcProps_.setIterateDelta(internal::parseDoubleExact(id, "iterateDelta"));
}
struct PendingPrintName { std::string name; std::size_t sheetIndex; std::string value; };
std::vector<PendingPrintName> pendingPrintNames;
const auto definedNameNodes = internal::tags(wb, "definedName");
if (options.maxDefinedNames != 0 && definedNameNodes.size() > options.maxDefinedNames)
    throw std::runtime_error("Workbook exceeds configured maxDefinedNames limit");
for(const auto& node : definedNameNodes){
    const auto name = internal::attribute(node,"name");
    const auto value = internal::tagText(node,"definedName");
    const auto local=internal::attribute(node,"localSheetId");
    if ((name == "_xlnm.Print_Area" || name == "_xlnm.Print_Titles") && !local.empty()) {
        pendingPrintNames.push_back({name, internal::parseIntegerExact<std::size_t>(local, "definedName localSheetId"), value});
        continue;
    }
    DefinedName item(name,value);
    if(!local.empty()) item.setLocalSheetId(internal::parseIntegerExact<std::size_t>(local, "definedName localSheetId"));
    item.setHidden(internal::attribute(node,"hidden")=="1");
    item.setComment(internal::attribute(node,"comment"));
    definedNames_.push_back(std::move(item));
}
const auto targets = internal::loadPackageRelationships(z, "xl/workbook.xml");
const auto sheetNodes = internal::tags(wb, "sheet");
std::size_t requestedActiveTab = 0;
bool hasRequestedActiveTab = false;
for (const auto& viewNode : internal::tags(wb, "workbookView")) {
    const auto active = internal::attribute(viewNode, "activeTab");
    if (!active.empty()) { requestedActiveTab = internal::parseIntegerExact<std::size_t>(active, "workbook activeTab"); hasRequestedActiveTab = true; }
    break;
}
const auto parseVisibility = [](std::string_view state) {
    if (state.empty() || state == "visible") return WorkbookSheetVisibility::Visible;
    if (state == "hidden") return WorkbookSheetVisibility::Hidden;
    if (state == "veryHidden") return WorkbookSheetVisibility::VeryHidden;
    throw std::runtime_error("Unsupported workbook sheet state: " + std::string(state));
};
std::size_t totalLoadedCells = 0;
std::size_t loadedWorksheetCount = 0;
std::size_t loadedChartsheetCount = 0;
for(const auto& s : sheetNodes) {
    const auto name = internal::attribute(s,"name");
    try {
        const auto rid = internal::attribute(s,"r:id");
        if (rid.empty()) throw std::runtime_error("workbook sheet has an empty relationship Id");
        const auto& relationship = internal::requirePackageRelationship(
            targets, rid, {}, "workbook sheet");
        if (internal::relationshipTypeEndsWith(relationship.type, "/worksheet")) {
            if (options.maxWorksheets != 0 && ++loadedWorksheetCount > options.maxWorksheets)
                throw std::runtime_error("Workbook exceeds configured maxWorksheets limit");
            const auto target = internal::requireInternalPackageTarget(
                z, "xl/workbook.xml", relationship, "workbook worksheet");
            auto& ws=addWorksheet(name);
            sheetOrder_.back().visibility = parseVisibility(internal::attribute(s, "state"));
            sourceSheetNames_.push_back(name);
            sourceSheetParts_.push_back(target);
            sourceSheetXml_.push_back(z.get(target));
            parseSheet(ws, sourceSheetXml_.back(), z, target, styleCatalog, dxfStyles, shared, date1904_,
                       options.maxCells, totalLoadedCells);
        } else if (internal::relationshipTypeEndsWith(relationship.type, "/chartsheet")) {
            if (options.maxChartsheets != 0 && ++loadedChartsheetCount > options.maxChartsheets)
                throw std::runtime_error("Workbook exceeds configured maxChartsheets limit");
            const auto target = internal::requireInternalPackageTarget(
                z, "xl/workbook.xml", relationship, "workbook chartsheet");
            internal::validateWorksheetName(name);
            const bool duplicateWorksheet = std::any_of(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
            const bool duplicateChartsheet = std::any_of(chartsheets_.begin(), chartsheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
            if (duplicateWorksheet || duplicateChartsheet) throw std::runtime_error("Duplicate workbook sheet name: " + name);
            chartsheets_.emplace_back(name);
            auto& chartSheet = chartsheets_.back();
            const auto chartSheetXml = z.get(target);
            internal::parseChartsheetModel(chartSheet, chartSheetXml);
            std::string drawingRelationshipId = "rId1";
            const auto chartSheetRelationships = internal::loadPackageRelationships(z, target);
            for (const auto& [relationshipId, record] : chartSheetRelationships) {
                if (internal::relationshipTypeEndsWith(record.type, "/drawing")) { drawingRelationshipId = relationshipId; break; }
            }

            if (chartSheet.hasPageSetup() && chartSheet.pageSetup().relationshipId()) {
                const auto& printerRelationship = internal::requirePackageRelationship(
                    chartSheetRelationships, *chartSheet.pageSetup().relationshipId(),
                    "/printerSettings", "chartsheet pageSetup");
                const auto printerTarget = internal::requireInternalPackageTarget(
                    z, target, printerRelationship, "chartsheet printerSettings");
                chartSheet.setImportedPrinterSettings(
                    printerTarget, *chartSheet.pageSetup().relationshipId(), z.get(printerTarget));
            }

            // Header/footer picture ownership: the chartsheet top-level
            // legacyDrawingHF relationship points at a VML drawing part that
            // hosts header/footer images. Preserve the identity and bytes so a
            // metadata edit never drops the picture part.
            const auto headerFooterXml = internal::tags(chartSheetXml, "legacyDrawingHF");
            if (!headerFooterXml.empty()) {
                const auto hfRelationshipId = internal::attribute(headerFooterXml.front(), "r:id");
                if (!hfRelationshipId.empty()) {
                    const auto hfIt = chartSheetRelationships.find(hfRelationshipId);
                    if (hfIt != chartSheetRelationships.end()
                        && internal::relationshipTypeEndsWith(hfIt->second.type, "/vmlDrawing")) {
                        const auto hfTarget = internal::requireInternalPackageTarget(
                            z, target, hfIt->second, "chartsheet header/footer VML");
                        chartSheet.setHeaderFooterDrawing(
                            hfTarget, hfRelationshipId, z.get(hfTarget));
                    }
                }
            }

            Worksheet chartLoader("XLPP_Chartsheet_Loader");
            loadCharts(chartLoader, chartSheetXml, z, target);
            const auto& constLoader = static_cast<const Worksheet&>(chartLoader);
            if (constLoader.chartCount() > 0) chartSheet.chart_ = constLoader.chart(0);
            chartSheet.setImportedSource(target, chartSheetXml, drawingRelationshipId);
            sheetOrder_.push_back({WorkbookSheetKind::Chartsheet, chartsheets_.size() - 1, parseVisibility(internal::attribute(s, "state"))});
        } else {
            throw std::runtime_error("Unsupported workbook sheet relationship Type: " + relationship.type);
        }
    } catch (const std::exception& e) {
        diagnostics_.errors.push_back("Sheet '" + name + "': " + e.what());
        if (!options.lenient) throw;
        continue;
    }
}
    if (!sheetOrder_.empty()) {
        const auto visibleIndex = firstVisibleWorkbookSheetIndex();
        if (visibleIndex >= sheetOrder_.size()) {
            if (!options.lenient) throw std::runtime_error("Workbook contains no visible sheet");
            sheetOrder_.front().visibility = WorkbookSheetVisibility::Visible;
            diagnostics_.warnings.push_back("Workbook contained no visible sheet; first loaded sheet was made visible");
        }
        firstVisibleWorkbookSheetIndex_ = firstVisibleWorkbookSheetIndex();
        activeWorkbookSheetIndex_ = firstVisibleWorkbookSheetIndex_;
        if (hasRequestedActiveTab && requestedActiveTab < sheetNodes.size()) {
            const auto requestedName = internal::attribute(sheetNodes[requestedActiveTab], "name");
            for (std::size_t i = 0; i < sheetOrder_.size(); ++i) {
                const auto& entry = sheetOrder_[i];
                const auto& candidateName = entry.kind == WorkbookSheetKind::Worksheet ? sheets_.at(entry.kindIndex).name() : chartsheets_.at(entry.kindIndex).name();
                if (internal::worksheetNamesEquivalent(candidateName, requestedName) && entry.visibility == WorkbookSheetVisibility::Visible) {
                    activeWorkbookSheetIndex_ = i;
                    break;
                }
            }
        }
    }
    // Apply built-in print names after all worksheets have been created.
    for (const auto& pending : pendingPrintNames) {
        if (pending.sheetIndex >= sheetOrder_.size()) continue;
        const auto& owner = sheetOrder_[pending.sheetIndex];
        if (owner.kind != WorkbookSheetKind::Worksheet || owner.kindIndex >= sheets_.size()) continue;
        auto& sheet = sheets_[owner.kindIndex];
        if (pending.name == "_xlnm.Print_Area") {
            sheet.setPrintArea(localPrintReference(pending.value));
            continue;
        }
        std::string rows;
        std::string cols;
        const auto localValue = localPrintReference(pending.value);
        std::size_t areaStart = 0;
        while (areaStart <= localValue.size()) {
            const auto comma = localValue.find(',', areaStart);
            const auto area = localValue.substr(areaStart, comma == std::string::npos ? std::string::npos : comma - areaStart);
            const auto first = area.find_first_not_of(' ');
            if (first != std::string::npos && std::isdigit(static_cast<unsigned char>(area[first]))) {
                if (!rows.empty()) rows += ',';
                rows += area;
            } else if (!area.empty()) {
                if (!cols.empty()) cols += ',';
                cols += area;
            }
            if (comma == std::string::npos) break;
            areaStart = comma + 1;
        }
        sheet.setPrintTitlesRows(rows);
        sheet.setPrintTitlesCols(cols);
    }
    preservedParts_.clear();
    const auto& overrides = packageContentTypes.overrides;
    const auto& defaults = packageContentTypes.defaults;
    for (const auto& partName : z.entryNames()) {
        if (isRegeneratedPart(partName)) continue;
        PreservedPart part;
        part.name = partName;
        part.data = z.get(partName);
        const auto overrideIt = overrides.find("/" + partName);
        if (overrideIt != overrides.end()) {
            part.overrideType = overrideIt->second;
        } else {
            const auto ext = extensionOf(partName);
            const auto defaultIt = defaults.find(ext);
            if (defaultIt != defaults.end()) { part.extension = ext; part.defaultType = defaultIt->second; }
        }
        part.compress = !(part.extension == "png" || part.extension == "jpg" || part.extension == "jpeg"
            || part.extension == "gif" || part.extension == "bmp" || part.extension == "bin");
        preservedParts_.push_back(std::move(part));
    }
    // Source-generated XL++ projects carry a stable project ID/profile. Mark
    // them so a later sheet insertion/removal followed by save() rebuilds the
    // document-module list. Arbitrary externally authored projects remain
    // preserved verbatim.
    if (const auto it = std::find_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
            return part.name == "xl/vbaProject.bin";
        }); it != preservedParts_.end()) {
        const std::vector<unsigned char> bytes(it->data.begin(), it->data.end());
        generatedVbaProject_ = internal::isXlppGeneratedVbaProjectBinary(bytes);
    }
    // After a successful load, mark all sheets clean so the next save can
    // benefit from differential caching.
    for (auto& s : sheets_) s.clearDirty();
}
}
