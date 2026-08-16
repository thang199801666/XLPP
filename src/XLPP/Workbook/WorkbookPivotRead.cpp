#include "WorkbookPivotRead.h"
#include "WorkbookPivotWrite.h"
#include "WorkbookChartReader.h"
#include "WorkbookPartXml.h"
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Pivot/PivotTable.h>
#include "../XML/XmlUtilities.h"
#include "../XML/NumericParsing.h"
#include "../Packaging/ZipArchive.h"
#include "../Packaging/RelationshipGraph.h"
#include <algorithm>
#include <cctype>
#include <cstdlib>
#include <limits>
#include <sstream>
#include <stdexcept>
#include <string>
#include <unordered_map>
#include <vector>
#include <cstdint>

namespace xlpp {
namespace internal {
struct LoadedPivotValueNode {
    std::string type;
    std::string value;
};

std::vector<LoadedPivotValueNode> pivotValueNodes(std::string_view container) {
    std::vector<LoadedPivotValueNode> result;
    std::size_t pos = 0;
    while ((pos = container.find('<', pos)) != std::string_view::npos) {
        if (pos + 1 >= container.size()) break;
        const char first = container[pos + 1];
        if (first == '/' || first == '?' || first == '!') { ++pos; continue; }
        const auto end = container.find('>', pos + 1);
        if (end == std::string_view::npos) break;
        auto head = container.substr(pos, end - pos + 1);
        std::size_t nameStart = 1;
        auto nameEnd = nameStart;
        while (nameEnd < head.size() && head[nameEnd] != ' ' && head[nameEnd] != '\t' && head[nameEnd] != '\r'
               && head[nameEnd] != '\n' && head[nameEnd] != '/' && head[nameEnd] != '>') ++nameEnd;
        auto name = std::string(head.substr(nameStart, nameEnd - nameStart));
        if (const auto colon = name.find(':'); colon != std::string::npos) name = name.substr(colon + 1);
        if (name == "x" || name == "n" || name == "s" || name == "b" || name == "e" || name == "d" || name == "m")
            result.push_back({name, xlpp::internal::attribute(head, "v")});
        pos = end + 1;
    }
    return result;
}

bool pivotBoolAttribute(const std::string& node, const char* name, bool defaultValue) {
    const auto value = xlpp::internal::attribute(node, name);
    if (value.empty()) return defaultValue;
    return value != "0" && value != "false";
}

int pivotIntAttribute(const std::string& node, const char* name, int defaultValue) {
    const auto value = xlpp::internal::attribute(node, name);
    if (value.empty()) return defaultValue;
    try { return std::stoi(value); } catch (...) { return defaultValue; }
}

std::string firstBalancedPivotElement(const std::string& xml, std::string_view localName) {
    const auto openToken = "<" + std::string(localName);
    const auto closeToken = "</" + std::string(localName) + ">";
    const auto begin = xml.find(openToken);
    if (begin == std::string::npos) return {};
    const auto headEnd = xml.find('>', begin + openToken.size());
    if (headEnd == std::string::npos) return {};
    if (headEnd > begin && xml[headEnd - 1] == '/') return xml.substr(begin, headEnd - begin + 1);

    std::size_t depth = 1;
    std::size_t pos = headEnd + 1;
    while (depth != 0 && pos < xml.size()) {
        const auto nextOpen = xml.find(openToken, pos);
        const auto nextClose = xml.find(closeToken, pos);
        if (nextClose == std::string::npos) return {};
        if (nextOpen != std::string::npos && nextOpen < nextClose) {
            const auto afterName = nextOpen + openToken.size();
            if (afterName < xml.size()) {
                const char boundary = xml[afterName];
                if (boundary == ' ' || boundary == '\t' || boundary == '\r' || boundary == '\n' || boundary == '>' || boundary == '/') {
                    const auto nestedHeadEnd = xml.find('>', afterName);
                    if (nestedHeadEnd == std::string::npos) return {};
                    if (nestedHeadEnd == nextOpen || xml[nestedHeadEnd - 1] != '/') ++depth;
                    pos = nestedHeadEnd + 1;
                    continue;
                }
            }
            pos = nextOpen + 1;
            continue;
        }
        --depth;
        pos = nextClose + closeToken.size();
    }
    return depth == 0 ? xml.substr(begin, pos - begin) : std::string{};
}

std::string pivotRelationshipTarget(const xlpp::internal::ZipArchive& z,
                                    const std::string& sourcePart,
                                    std::string_view relationshipKindFragment) {
    const auto relPart = xlpp::internal::RelationshipGraph::relationshipsPartForSource(sourcePart);
    if (!z.contains(relPart)) return {};
    for (const auto& rel : xlpp::internal::tags(z.get(relPart), "Relationship")) {
        if (xlpp::internal::attribute(rel, "Type").find(relationshipKindFragment) == std::string::npos) continue;
        return resolvePackagePart(sourcePart, xlpp::internal::attribute(rel, "Target"));
    }
    return {};
}

xlpp::PivotCache loadPivotCache(const xlpp::internal::ZipArchive& z,
                                const std::string& cachePart,
                                int cacheId) {
    xlpp::PivotCache cache;
    cache.setCacheId(std::max(1, cacheId));
    if (!cachePart.empty()) cache.setSharedCacheKey(cachePart);
    if (cachePart.empty() || !z.contains(cachePart)) return cache;
    const auto cacheXmlText = z.get(cachePart);
    const auto roots = xlpp::internal::tags(cacheXmlText, "pivotCacheDefinition");
    if (roots.empty()) return cache;
    const auto& root = roots.front();
    cache.setRefreshOnLoad(pivotBoolAttribute(root, "refreshOnLoad", true));
    cache.setSaveData(pivotBoolAttribute(root, "saveData", true));
    cache.setEnableRefresh(pivotBoolAttribute(root, "enableRefresh", true));
    cache.setMissingItemsLimit(pivotIntAttribute(root, "missingItemsLimit", -1));

    const auto sources = xlpp::internal::tags(root, "worksheetSource");
    if (!sources.empty()) {
        const auto sheet = xlpp::internal::attribute(sources.front(), "sheet");
        const auto ref = xlpp::internal::attribute(sources.front(), "ref");
        const auto name = xlpp::internal::attribute(sources.front(), "name");
        if (!name.empty()) cache.setSourceName(name);
        if (!ref.empty()) cache.setSourceData(sheet.empty() ? ref : quotePivotSheetName(sheet) + "!" + ref);
    }

    // OLAP pivot caches bind to a cube/data connection instead of a worksheet
    // range. Keep the source identity and olapInfo metadata plus any
    // calculatedMember nodes lossless so a selective patch never truncates MDX.
    const auto cacheSources = xlpp::internal::tags(root, "cacheSource");
    if (!cacheSources.empty()) {
        const auto sourceType = xlpp::internal::attribute(cacheSources.front(), "type");
        if (sourceType == "olap") {
            xlpp::PivotOlapSourceInfo olap;
            olap.sourceType = sourceType;
            olap.connectionId = pivotIntAttribute(cacheSources.front(), "connectionId", -1);
            const auto olapInfos = xlpp::internal::tags(cacheSources.front(), "olapInfo");
            if (!olapInfos.empty()) {
                olap.preserveFormatting = pivotBoolAttribute(olapInfos.front(), "preserveFormatting", true);
                olap.localCube = xlpp::internal::attribute(olapInfos.front(), "localCube");
                olap.localConnection = xlpp::internal::attribute(olapInfos.front(), "localConnection");
                olap.rawOlapInfoXml = olapInfos.front();
            }
            for (const auto& memberNode : xlpp::internal::tags(root, "calculatedMember")) {
                xlpp::PivotCalculatedMember member;
                member.name = xlpp::internal::attribute(memberNode, "name");
                member.mdx = xlpp::internal::attribute(memberNode, "mdx");
                member.memberName = xlpp::internal::attribute(memberNode, "memberName");
                member.hierarchy = pivotIntAttribute(memberNode, "hierarchy", -1);
                member.solveOrder = xlpp::internal::attribute(memberNode, "solveOrder");
                member.set = xlpp::internal::attribute(memberNode, "set");
                member.rawXml = memberNode;
                cache.calculatedMembers().push_back(std::move(member));
            }
            cache.setOlap(std::move(olap));
        }
    }

    std::vector<std::vector<std::string>> sharedItems;
    std::vector<std::vector<xlpp::PivotCacheValueKind>> sharedItemKinds;
    for (const auto& fieldNode : xlpp::internal::tags(root, "cacheField")) {
        cache.addField(xlpp::internal::attribute(fieldNode, "name"));
        const auto fieldIndex = cache.fields().size() - 1;
        const auto formula = xlpp::internal::attribute(fieldNode, "formula");
        if (!formula.empty()) cache.setFieldFormula(fieldIndex, formula);
        cache.setFieldCaption(fieldIndex, xlpp::internal::attribute(fieldNode, "caption"));
        cache.setFieldNumberFormatId(fieldIndex, std::max(0, pivotIntAttribute(fieldNode, "numFmtId", 0)));
        cache.setFieldDatabaseField(fieldIndex, pivotBoolAttribute(fieldNode, "databaseField", formula.empty()));
        std::vector<std::string> items;
        std::vector<xlpp::PivotCacheValueKind> itemKinds;
        const auto shared = xlpp::internal::tags(fieldNode, "sharedItems");
        if (!shared.empty()) {
            for (const auto& value : pivotValueNodes(shared.front())) {
                items.push_back(value.type == "b" ? (value.value == "1" ? "true" : "false") : value.value);
                itemKinds.push_back(publicPivotValueKind(value.type));
            }
        }
        const auto groups = xlpp::internal::tags(fieldNode, "fieldGroup");
        if (!groups.empty()) {
            xlpp::PivotFieldGroup group;
            const auto& groupNode = groups.front();
            group.parentField = pivotIntAttribute(groupNode, "par", -1);
            group.baseField = pivotIntAttribute(groupNode, "base", -1);
            const auto ranges = xlpp::internal::tags(groupNode, "rangePr");
            if (!ranges.empty()) {
                const auto& range = ranges.front();
                group.groupBy = xlpp::internal::attribute(range, "groupBy");
                group.autoStart = pivotBoolAttribute(range, "autoStart", true);
                group.autoEnd = pivotBoolAttribute(range, "autoEnd", true);
                const auto readDouble = [&](const char* name) -> std::optional<double> {
                    const auto value = xlpp::internal::attribute(range, name);
                    if (value.empty()) return std::nullopt;
                    try { return std::stod(value); } catch (...) { return std::nullopt; }
                };
                group.startNumber = readDouble("startNum");
                group.endNumber = readDouble("endNum");
                group.interval = readDouble("groupInterval");
                group.startDate = xlpp::internal::attribute(range, "startDate");
                group.endDate = xlpp::internal::attribute(range, "endDate");
            }
            const auto groupItems = xlpp::internal::tags(groupNode, "groupItems");
            if (!groupItems.empty()) {
                for (const auto& value : pivotValueNodes(groupItems.front())) {
                    const auto kind = value.type == "n" ? xlpp::PivotCacheValueKind::Number
                        : value.type == "d" ? xlpp::PivotCacheValueKind::DateTime
                        : value.type == "b" ? xlpp::PivotCacheValueKind::Boolean
                        : value.type == "e" ? xlpp::PivotCacheValueKind::Error
                        : value.type == "m" ? xlpp::PivotCacheValueKind::Missing
                        : xlpp::PivotCacheValueKind::String;
                    group.addTypedGroupItem(kind, value.value);
                }
            }
            cache.setFieldGroup(fieldIndex, std::move(group));
        }
        sharedItems.push_back(std::move(items));
        sharedItemKinds.push_back(std::move(itemKinds));
    }

    const auto recordsPart = pivotRelationshipTarget(z, cachePart, "/pivotCacheRecords");
    if (!recordsPart.empty() && z.contains(recordsPart)) {
        const auto recordsXml = z.get(recordsPart);
        for (const auto& recordNode : xlpp::internal::tags(recordsXml, "r")) {
            const auto values = pivotValueNodes(recordNode);
            if (values.empty() && !cache.fields().empty()) continue;
            std::vector<std::string> record;
            std::vector<xlpp::PivotCacheValueKind> kinds;
            record.reserve(cache.fields().size());
            kinds.reserve(cache.fields().size());
            for (std::size_t fieldIndex = 0; fieldIndex < cache.fields().size(); ++fieldIndex) {
                if (fieldIndex >= values.size()) {
                    record.emplace_back();
                    kinds.push_back(xlpp::PivotCacheValueKind::Missing);
                    continue;
                }
                const auto& value = values[fieldIndex];
                if (value.type == "x") {
                    try {
                        const auto itemIndex = xlpp::internal::parseIntegerExact<std::size_t>(value.value, "pivot shared-item index");
                        if (fieldIndex >= sharedItems.size() || itemIndex >= sharedItems[fieldIndex].size())
                            throw std::runtime_error("Pivot shared-item index is out of range");
                        record.push_back(sharedItems[fieldIndex][itemIndex]);
                        kinds.push_back(fieldIndex < sharedItemKinds.size() && itemIndex < sharedItemKinds[fieldIndex].size()
                            ? sharedItemKinds[fieldIndex][itemIndex] : xlpp::PivotCacheValueKind::String);
                    } catch (...) {
                        record.emplace_back();
                        kinds.push_back(xlpp::PivotCacheValueKind::Missing);
                    }
                } else if (value.type == "b") {
                    record.push_back(value.value == "1" ? "true" : "false");
                    kinds.push_back(xlpp::PivotCacheValueKind::Boolean);
                } else {
                    record.push_back(value.value);
                    kinds.push_back(publicPivotValueKind(value.type));
                }
            }
            cache.addTypedRecord(std::move(record), std::move(kinds));
        }
    }
    return cache;
}

void applyLoadedPivotFieldModel(xlpp::PivotField& field,
                                const std::string& node,
                                std::size_t fieldIndex,
                                const std::vector<PivotSharedItem>* sharedItems = nullptr) {
    field.setFieldIndex(static_cast<int>(fieldIndex));
    field.setShowAll(pivotBoolAttribute(node, "showAll", false));
    field.setCompact(pivotBoolAttribute(node, "compact", true));
    field.setOutline(pivotBoolAttribute(node, "outline", true));
    field.setDefaultSubtotal(pivotBoolAttribute(node, "defaultSubtotal", true));
    for (const auto& subtotal : {"sum", "countA", "avg", "max", "min", "product", "count", "stdDev", "stdDevP", "var", "varP"}) {
        const auto attribute = std::string(subtotal) + "Subtotal";
        if (pivotBoolAttribute(node, attribute.c_str(), false)) field.addSubtotal(subtotal);
    }
    const auto sort = xlpp::internal::attribute(node, "sortType");
    if (sort == "ascending") field.setSortType(1);
    else if (sort == "descending") field.setSortType(2);
    const auto items = xlpp::internal::tags(node, "items");
    if (!items.empty()) {
        field.clearItems();
        for (const auto& itemNode : xlpp::internal::tags(items.front(), "item")) {
            xlpp::PivotFieldItem item;
            item.cacheIndex = pivotIntAttribute(itemNode, "x", -1);
            if (sharedItems && item.cacheIndex >= 0 && static_cast<std::size_t>(item.cacheIndex) < sharedItems->size()) {
                const auto& shared = (*sharedItems)[static_cast<std::size_t>(item.cacheIndex)];
                item.bindCacheValue(shared.value, publicPivotValueKind(shared.kind));
            }
            item.type = xlpp::internal::attribute(itemNode, "t");
            item.caption = xlpp::internal::attribute(itemNode, "n");
            item.hidden = pivotBoolAttribute(itemNode, "h", false);
            item.showDetails = pivotBoolAttribute(itemNode, "sd", true);
            item.formula = pivotBoolAttribute(itemNode, "f", false);
            item.missing = pivotBoolAttribute(itemNode, "m", false);
            item.rawXml = itemNode;
            field.addItem(std::move(item));
        }
    }
}

void loadPivotTables(xlpp::Worksheet& ws,
                     const std::string& sheetXml,
                     const xlpp::internal::ZipArchive& z,
                     const std::string& sheetPart) {
    const auto ownerBlocks = xlpp::internal::tags(sheetXml, "pivotTableParts");
    const auto relPart = xlpp::internal::RelationshipGraph::relationshipsPartForSource(sheetPart);
    if (!z.contains(relPart)) return;
    std::unordered_map<std::string, std::string> pivotTargets;
    std::vector<std::string> pivotOrder;
    for (const auto& rel : xlpp::internal::tags(z.get(relPart), "Relationship")) {
        if (xlpp::internal::attribute(rel, "Type").find("/pivotTable") == std::string::npos) continue;
        const auto id = xlpp::internal::attribute(rel, "Id");
        pivotTargets[id] = resolvePackagePart(sheetPart, xlpp::internal::attribute(rel, "Target"));
    }
    // Most Excel producers emit pivotTableParts owner nodes. LibreOffice can
    // legitimately expose the pivot only through the worksheet relationship
    // graph, so fall back to relationship order when the owner block is absent.
    if (!ownerBlocks.empty()) {
        for (const auto& partNode : xlpp::internal::tags(ownerBlocks.front(), "pivotTablePart")) {
            const auto rid = xlpp::internal::attribute(partNode, "r:id");
            if (pivotTargets.contains(rid)) pivotOrder.push_back(rid);
        }
    } else {
        for (const auto& rel : xlpp::internal::tags(z.get(relPart), "Relationship")) {
            if (xlpp::internal::attribute(rel, "Type").find("/pivotTable") == std::string::npos) continue;
            const auto rid = xlpp::internal::attribute(rel, "Id");
            if (pivotTargets.contains(rid)) pivotOrder.push_back(rid);
        }
    }

    for (const auto& rid : pivotOrder) {
        const auto targetIt = pivotTargets.find(rid);
        if (targetIt == pivotTargets.end() || !z.contains(targetIt->second)) continue;
        const auto& pivotPart = targetIt->second;
        const auto pivotXmlText = z.get(pivotPart);
        const auto roots = xlpp::internal::tags(pivotXmlText, "pivotTableDefinition");
        if (roots.empty()) continue;
        const auto& root = roots.front();
        xlpp::PivotTable pivot(xlpp::internal::attribute(root, "name"));
        const auto cacheId = std::max(1, pivotIntAttribute(root, "cacheId", 1));
        const auto cachePart = pivotRelationshipTarget(z, pivotPart, "/pivotCacheDefinition");
        pivot.cache() = loadPivotCache(z, cachePart, cacheId);
        const auto loadedSharedItems = pivotSharedItems(pivot.cache());

        const auto locations = xlpp::internal::tags(root, "location");
        if (!locations.empty()) pivot.setLocation(xlpp::internal::attribute(locations.front(), "ref"));
        pivot.setRowGrandTotals(pivotBoolAttribute(root, "rowGrandTotals", true));
        pivot.setColumnGrandTotals(pivotBoolAttribute(root, "colGrandTotals", true));
        pivot.setCompact(pivotBoolAttribute(root, "compact", true));
        pivot.setOutline(pivotBoolAttribute(root, "outline", true));
        pivot.setPreserveFormatting(pivotBoolAttribute(root, "preserveFormatting", true));
        pivot.setUseAutoFormatting(pivotBoolAttribute(root, "useAutoFormatting", true));
        pivot.setShowDrill(pivotBoolAttribute(root, "showDrill", true));
        pivot.setMultipleFieldFilters(pivotBoolAttribute(root, "multipleFieldFilters", false));
        const auto rootChartFormat = xlpp::internal::attribute(root, "chartFormat");
        if (!rootChartFormat.empty()) {
            try { pivot.setChartFormatIndex(static_cast<std::uint32_t>(std::stoul(rootChartFormat))); } catch (...) {}
        }
        const auto styles = xlpp::internal::tags(root, "pivotTableStyleInfo");
        if (!styles.empty()) {
            const auto& style = styles.front();
            const auto name = xlpp::internal::attribute(style, "name");
            if (!name.empty()) pivot.setStyleName(name);
            pivot.setShowRowHeaders(pivotBoolAttribute(style, "showRowHeaders", true));
            pivot.setShowColumnHeaders(pivotBoolAttribute(style, "showColHeaders", true));
            pivot.setShowRowStripes(pivotBoolAttribute(style, "showRowStripes", false));
            pivot.setShowColumnStripes(pivotBoolAttribute(style, "showColStripes", false));
        }

        const auto pivotFieldNodes = xlpp::internal::tags(root, "pivotField");
        auto fieldName = [&](int index) {
            return index >= 0 && static_cast<std::size_t>(index) < pivot.cache().fields().size()
                ? pivot.cache().fields()[static_cast<std::size_t>(index)]
                : "Field" + std::to_string(index + 1);
        };
        auto loadAxis = [&](const char* containerName, const char* childName, auto adder) {
            const auto containers = xlpp::internal::tags(root, containerName);
            if (containers.empty()) return;
            for (const auto& child : xlpp::internal::tags(containers.front(), childName)) {
                const auto index = pivotIntAttribute(child, "x", -1);
                if (index < 0) continue;
                auto& field = adder(fieldName(index));
                if (static_cast<std::size_t>(index) < pivotFieldNodes.size())
                    applyLoadedPivotFieldModel(field, pivotFieldNodes[static_cast<std::size_t>(index)], static_cast<std::size_t>(index),
                                               static_cast<std::size_t>(index) < loadedSharedItems.size() ? &loadedSharedItems[static_cast<std::size_t>(index)] : nullptr);
                else field.setFieldIndex(index);
            }
        };
        loadAxis("rowFields", "field", [&](std::string name) -> xlpp::PivotField& { return pivot.addRowField(std::move(name)); });
        loadAxis("colFields", "field", [&](std::string name) -> xlpp::PivotField& { return pivot.addColumnField(std::move(name)); });

        const auto pageContainers = xlpp::internal::tags(root, "pageFields");
        if (!pageContainers.empty()) {
            for (const auto& pageNode : xlpp::internal::tags(pageContainers.front(), "pageField")) {
                const auto index = pivotIntAttribute(pageNode, "fld", -1);
                if (index < 0) continue;
                auto& field = pivot.addPageField(fieldName(index));
                if (static_cast<std::size_t>(index) < pivotFieldNodes.size())
                    applyLoadedPivotFieldModel(field, pivotFieldNodes[static_cast<std::size_t>(index)], static_cast<std::size_t>(index),
                                               static_cast<std::size_t>(index) < loadedSharedItems.size() ? &loadedSharedItems[static_cast<std::size_t>(index)] : nullptr);
                else field.setFieldIndex(index);
                auto& setting = pivot.pageFieldSettings().back();
                setting.setFieldIndex(index);
                setting.setHierarchy(pivotIntAttribute(pageNode, "hier", -1));
                setting.setItem(pivotIntAttribute(pageNode, "item", -1));
                setting.setName(xlpp::internal::attribute(pageNode, "name"));
            }
        }

        const auto dataContainers = xlpp::internal::tags(root, "dataFields");
        if (!dataContainers.empty()) {
            for (const auto& dataNode : xlpp::internal::tags(dataContainers.front(), "dataField")) {
                const auto index = pivotIntAttribute(dataNode, "fld", -1);
                if (index < 0) continue;
                auto& field = pivot.addDataField(index);
                field.setName(fieldName(index));
                field.setDisplayName(xlpp::internal::attribute(dataNode, "name"));
                const auto subtotal = xlpp::internal::attribute(dataNode, "subtotal");
                if (!subtotal.empty()) { try { field.setSubtotal(subtotal); } catch (...) {} }
                const auto showDataAs = xlpp::internal::attribute(dataNode, "showDataAs");
                if (!showDataAs.empty()) { try { field.setShowDataAs(showDataAs); } catch (...) {} }
                field.setBaseField(pivotIntAttribute(dataNode, "baseField", -1));
                field.setBaseItem(pivotIntAttribute(dataNode, "baseItem", -1));
                field.setNumberFormatId(pivotIntAttribute(dataNode, "numFmtId", 0));
            }
        }
        const auto chartFormatContainers = xlpp::internal::tags(root, "chartFormats");
        if (!chartFormatContainers.empty()) {
            for (const auto& chartFormatNode : xlpp::internal::tags(chartFormatContainers.front(), "chartFormat")) {
                xlpp::PivotChartFormat format;
                format.chartIndex = static_cast<std::uint32_t>(std::max(0, pivotIntAttribute(chartFormatNode, "chart", 0)));
                format.formatId = static_cast<std::uint32_t>(std::max(0, pivotIntAttribute(chartFormatNode, "format", 0)));
                format.series = pivotBoolAttribute(chartFormatNode, "series", false);
                const auto pivotAreas = xlpp::internal::tags(chartFormatNode, "pivotArea");
                if (!pivotAreas.empty()) format.pivotAreaXml = pivotAreas.front();
                pivot.addChartFormat(std::move(format));
            }
        }

        const auto filterContainer = firstBalancedPivotElement(root, "filters");
        if (!filterContainer.empty()) {
            for (const auto& filterNode : xlpp::internal::tags(filterContainer, "filter")) {
                const auto fieldIndex = pivotIntAttribute(filterNode, "fld", -1);
                const auto type = xlpp::internal::attribute(filterNode, "type");
                // Nested AutoFilter criteria can also contain <filter> nodes;
                // those do not carry PivotFilter fld/type attributes.
                if (fieldIndex < 0 || type.empty()) continue;
                xlpp::PivotFilter filter;
                filter.fieldIndex = fieldIndex;
                filter.type = type;
                filter.id = static_cast<std::uint32_t>(std::max(0, pivotIntAttribute(filterNode, "id", 0)));
                filter.evaluationOrder = pivotIntAttribute(filterNode, "evalOrder", 0);
                filter.measureField = pivotIntAttribute(filterNode, "iMeasureFld", -1);
                filter.measureHierarchy = pivotIntAttribute(filterNode, "iMeasureHier", -1);
                filter.memberPropertyField = pivotIntAttribute(filterNode, "mpFld", -1);
                filter.name = xlpp::internal::attribute(filterNode, "name");
                filter.description = xlpp::internal::attribute(filterNode, "description");
                filter.stringValue1 = xlpp::internal::attribute(filterNode, "stringValue1");
                filter.stringValue2 = xlpp::internal::attribute(filterNode, "stringValue2");
                const auto autoFilters = xlpp::internal::tags(filterNode, "autoFilter");
                if (!autoFilters.empty()) filter.autoFilterXml = autoFilters.front();
                pivot.addFilter(std::move(filter));
            }
        }

        ws.addLoadedPivotTable(std::move(pivot));
    }
}

} // namespace internal
} // namespace xlpp

