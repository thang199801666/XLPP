#include "RelationshipGraph.h"
#include "../XML/XmlUtilities.h"
#include <algorithm>
#include <array>
#include <deque>
#include <sstream>
#include <unordered_map>
#include <unordered_set>

namespace xlpp::internal {
namespace {

std::string xmlEscapeLocal(const std::string& value) {
    std::string result;
    result.reserve(value.size());
    for (char c : value) {
        switch (c) {
        case '&': result += "&amp;"; break;
        case '<': result += "&lt;"; break;
        case '>': result += "&gt;"; break;
        case '"': result += "&quot;"; break;
        case '\'': result += "&apos;"; break;
        default: result += c; break;
        }
    }
    return result;
}

bool isRelationshipsPart(const std::string& name) {
    return name == "_rels/.rels" ||
        (name.size() > 5 && name.rfind(".rels") == name.size() - 5 && name.find("/_rels/") != std::string::npos);
}

bool isInfrastructurePart(const std::string& name) {
    return name == "[Content_Types].xml" || isRelationshipsPart(name);
}

std::string relationshipLabel(const xlpp::PreservedRelationship& relationship) {
    return (relationship.sourcePart.empty() ? std::string("/") : relationship.sourcePart) +
        "#" + relationship.id + " -> " + relationship.target;
}

std::string_view localName(std::string_view qualifiedName) {
    const auto colon = qualifiedName.find_last_of(':');
    return colon == std::string_view::npos ? qualifiedName : qualifiedName.substr(colon + 1);
}

std::vector<std::string> elementsByLocalName(std::string_view xml, std::string_view wanted) {
    std::vector<std::string> result;
    std::size_t position = 0;
    while ((position = xml.find('<', position)) != std::string_view::npos) {
        if (position + 1 >= xml.size()) break;
        const char lead = xml[position + 1];
        if (lead == '/' || lead == '?' || lead == '!') { ++position; continue; }

        std::size_t nameEnd = position + 1;
        while (nameEnd < xml.size()) {
            const char c = xml[nameEnd];
            if (c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '/' || c == '>') break;
            ++nameEnd;
        }
        const auto qualifiedName = xml.substr(position + 1, nameEnd - position - 1);
        if (localName(qualifiedName) != wanted) { position = nameEnd; continue; }

        bool quoted = false;
        char quote = 0;
        std::size_t openEnd = nameEnd;
        for (; openEnd < xml.size(); ++openEnd) {
            const char c = xml[openEnd];
            if (quoted) {
                if (c == quote) quoted = false;
            } else if (c == '"' || c == '\'') {
                quoted = true;
                quote = c;
            } else if (c == '>') {
                break;
            }
        }
        if (openEnd >= xml.size()) break;

        std::size_t beforeEnd = openEnd;
        while (beforeEnd > position && (xml[beforeEnd - 1] == ' ' || xml[beforeEnd - 1] == '\t' ||
               xml[beforeEnd - 1] == '\r' || xml[beforeEnd - 1] == '\n')) --beforeEnd;
        if (beforeEnd > position && xml[beforeEnd - 1] == '/') {
            result.emplace_back(xml.substr(position, openEnd - position + 1));
            position = openEnd + 1;
            continue;
        }

        const std::string closing = "</" + std::string(qualifiedName) + ">";
        const auto close = xml.find(closing, openEnd + 1);
        if (close == std::string_view::npos) {
            result.emplace_back(xml.substr(position, openEnd - position + 1));
            position = openEnd + 1;
            continue;
        }
        result.emplace_back(xml.substr(position, close + closing.size() - position));
        position = close + closing.size();
    }
    return result;
}

std::string attributeByLocalName(std::string_view element, std::string_view wanted) {
    const auto openEnd = element.find('>');
    if (openEnd == std::string_view::npos) return {};
    std::size_t position = element.find('<');
    if (position == std::string_view::npos) return {};
    ++position;
    while (position < openEnd && element[position] != ' ' && element[position] != '\t' &&
           element[position] != '\r' && element[position] != '\n' && element[position] != '>') ++position;

    while (position < openEnd) {
        while (position < openEnd && (element[position] == ' ' || element[position] == '\t' ||
               element[position] == '\r' || element[position] == '\n' || element[position] == '/')) ++position;
        if (position >= openEnd) break;
        const auto nameBegin = position;
        while (position < openEnd && element[position] != '=' && element[position] != ' ' &&
               element[position] != '\t' && element[position] != '\r' && element[position] != '\n') ++position;
        const auto nameEnd = position;
        while (position < openEnd && (element[position] == ' ' || element[position] == '\t' ||
               element[position] == '\r' || element[position] == '\n')) ++position;
        if (position >= openEnd || element[position] != '=') {
            while (position < openEnd && element[position] != ' ' && element[position] != '\t' &&
                   element[position] != '\r' && element[position] != '\n') ++position;
            continue;
        }
        ++position;
        while (position < openEnd && (element[position] == ' ' || element[position] == '\t' ||
               element[position] == '\r' || element[position] == '\n')) ++position;
        if (position >= openEnd || (element[position] != '"' && element[position] != '\'')) {
            if (position < openEnd) ++position;
            continue;
        }
        const char quote = element[position++];
        const auto valueBegin = position;
        const auto valueEnd = element.find(quote, position);
        if (valueEnd == std::string_view::npos || valueEnd > openEnd) return {};
        const auto attributeName = element.substr(nameBegin, nameEnd - nameBegin);
        if (localName(attributeName) == wanted)
            return xmlUnescape(element.substr(valueBegin, valueEnd - valueBegin));
        position = valueEnd + 1;
    }
    return {};
}

std::unordered_set<std::string> relationshipReferenceIds(std::string_view xml) {
    std::unordered_set<std::string> result;
    std::size_t position = 0;
    while ((position = xml.find('<', position)) != std::string_view::npos) {
        if (position + 1 >= xml.size()) break;
        const char lead = xml[position + 1];
        if (lead == '/' || lead == '?' || lead == '!') { ++position; continue; }

        std::size_t openEnd = position + 1;
        bool quoted = false;
        char quote = 0;
        for (; openEnd < xml.size(); ++openEnd) {
            const char c = xml[openEnd];
            if (quoted) {
                if (c == quote) quoted = false;
            } else if (c == '"' || c == '\'') {
                quoted = true;
                quote = c;
            } else if (c == '>') {
                break;
            }
        }
        if (openEnd >= xml.size()) break;

        auto tag = xml.substr(position, openEnd - position + 1);
        std::size_t cursor = 1;
        while (cursor < tag.size() && tag[cursor] != ' ' && tag[cursor] != '\t' &&
               tag[cursor] != '\r' && tag[cursor] != '\n' && tag[cursor] != '>') ++cursor;
        while (cursor < tag.size()) {
            while (cursor < tag.size() && (tag[cursor] == ' ' || tag[cursor] == '\t' ||
                   tag[cursor] == '\r' || tag[cursor] == '\n' || tag[cursor] == '/')) ++cursor;
            if (cursor >= tag.size() || tag[cursor] == '>') break;
            const auto nameBegin = cursor;
            while (cursor < tag.size() && tag[cursor] != '=' && tag[cursor] != ' ' &&
                   tag[cursor] != '\t' && tag[cursor] != '\r' && tag[cursor] != '\n') ++cursor;
            const auto nameEnd = cursor;
            while (cursor < tag.size() && (tag[cursor] == ' ' || tag[cursor] == '\t' ||
                   tag[cursor] == '\r' || tag[cursor] == '\n')) ++cursor;
            if (cursor >= tag.size() || tag[cursor] != '=') {
                while (cursor < tag.size() && tag[cursor] != ' ' && tag[cursor] != '\t' &&
                       tag[cursor] != '\r' && tag[cursor] != '\n' && tag[cursor] != '>') ++cursor;
                continue;
            }
            ++cursor;
            while (cursor < tag.size() && (tag[cursor] == ' ' || tag[cursor] == '\t' ||
                   tag[cursor] == '\r' || tag[cursor] == '\n')) ++cursor;
            if (cursor >= tag.size() || (tag[cursor] != '"' && tag[cursor] != '\'')) continue;
            const char valueQuote = tag[cursor++];
            const auto valueBegin = cursor;
            const auto valueEnd = tag.find(valueQuote, cursor);
            if (valueEnd == std::string_view::npos) break;

            const auto qualifiedName = tag.substr(nameBegin, nameEnd - nameBegin);
            const auto attributeLocalName = localName(qualifiedName);
            // Relationship-bearing DrawingML attributes are namespace-qualified
            // id/embed/link attributes.  Deliberately exclude unqualified id
            // attributes such as cNvPr/@id.
            if (qualifiedName.find(':') != std::string_view::npos &&
                (attributeLocalName == "id" || attributeLocalName == "embed" || attributeLocalName == "link")) {
                result.insert(xmlUnescape(tag.substr(valueBegin, valueEnd - valueBegin)));
            }
            cursor = valueEnd + 1;
        }
        position = openEnd + 1;
    }
    return result;
}

bool relationshipTypeEndsWith(const xlpp::PreservedRelationship& relationship, std::string_view suffix) {
    return relationship.type.size() >= suffix.size()
        && relationship.type.compare(relationship.type.size() - suffix.size(), suffix.size(), suffix) == 0;
}

const xlpp::PreservedRelationship* relationshipById(
    const std::vector<xlpp::PreservedRelationship>& relationships,
    const std::string& sourcePart,
    const std::string& id) {
    const auto it = std::find_if(relationships.begin(), relationships.end(), [&](const auto& relationship) {
        return relationship.sourcePart == sourcePart && relationship.id == id;
    });
    return it == relationships.end() ? nullptr : &*it;
}

struct ObjectInspectionResult {
    PackageObjectInventory inventory;
    std::vector<std::string> errors;
};

ObjectInspectionResult inspectObjectGraph(
    const ZipArchive& archive,
    const std::vector<xlpp::PreservedRelationship>& relationships,
    const std::unordered_set<std::string>& entries) {
    ObjectInspectionResult result;
    std::unordered_set<std::string> consumedObjectRelationships;
    const auto relationshipKey = [](const std::string& sourcePart, const std::string& id) {
        return sourcePart + "\n" + id;
    };

    const auto resolveOwnerReference = [&](const std::string& sourcePart,
                                           const std::string& id,
                                           std::string_view expectedTypeSuffix,
                                           const std::string& ownerLabel,
                                           bool allowExternal = false) -> std::string {
        if (id.empty()) {
            result.errors.push_back(ownerLabel + " has no relationship ID");
            return {};
        }
        const auto* relationship = relationshipById(relationships, sourcePart, id);
        if (!relationship) {
            result.errors.push_back(ownerLabel + " references missing relationship " + id + " from " +
                (sourcePart.empty() ? std::string("/") : sourcePart));
            return {};
        }
        if (!relationshipTypeEndsWith(*relationship, expectedTypeSuffix)) {
            result.errors.push_back(ownerLabel + " uses relationship " + id + " with unexpected type " + relationship->type);
            return {};
        }
        consumedObjectRelationships.insert(relationshipKey(sourcePart, id));
        if (relationship->targetMode == "External") {
            if (allowExternal) return relationship->target;
            result.errors.push_back(ownerLabel + " unexpectedly targets an external resource through " + id);
            return {};
        }
        const auto resolved = RelationshipGraph::resolveTarget(sourcePart, relationship->target);
        if (resolved.empty() || !entries.count(resolved)) {
            result.errors.push_back(ownerLabel + " relationship " + id + " resolves to missing part " + resolved);
            return {};
        }
        return resolved;
    };

    const auto officeDocument = std::find_if(relationships.begin(), relationships.end(), [](const auto& relationship) {
        return relationship.sourcePart.empty() && relationshipTypeEndsWith(relationship, "/officeDocument");
    });
    if (officeDocument == relationships.end()) {
        result.errors.push_back("Package root has no officeDocument relationship");
        return result;
    }
    const auto workbookPart = RelationshipGraph::resolveTarget({}, officeDocument->target);
    if (workbookPart.empty() || !archive.contains(workbookPart)) {
        result.errors.push_back("Root officeDocument relationship resolves to missing workbook part " + workbookPart);
        return result;
    }

    const auto workbookXml = archive.get(workbookPart);
    std::unordered_map<std::string, std::string> pivotCachePartById;
    std::vector<std::string> pivotCacheParts;
    std::unordered_set<std::string> uniquePivotCacheParts;
    for (const auto& node : elementsByLocalName(workbookXml, "pivotCache")) {
        const auto id = attributeByLocalName(node, "id");
        const auto target = resolveOwnerReference(workbookPart, id, "/pivotCacheDefinition", "Workbook pivotCache");
        if (target.empty()) continue;
        if (uniquePivotCacheParts.insert(target).second) {
            ++result.inventory.pivotCaches;
            pivotCacheParts.push_back(target);
        }
        const auto cacheId = attributeByLocalName(node, "cacheId");
        if (cacheId.empty()) {
            result.errors.push_back("Workbook pivotCache relationship " + id + " has no cacheId");
            continue;
        }
        const auto [it, inserted] = pivotCachePartById.emplace(cacheId, target);
        if (!inserted && it->second != target)
            result.errors.push_back("Workbook contains duplicate pivot cacheId " + cacheId + " for different cache definitions");
    }

    std::vector<std::string> worksheetParts;
    std::vector<std::string> chartsheetParts;
    for (const auto& node : elementsByLocalName(workbookXml, "sheet")) {
        const auto name = attributeByLocalName(node, "name");
        const auto id = attributeByLocalName(node, "id");
        const auto* relationship = relationshipById(relationships, workbookPart, id);
        if (!relationship) {
            result.errors.push_back("Workbook sheet" + (name.empty() ? std::string{} : " '" + name + "'") + " references missing relationship " + id);
            continue;
        }
        if (relationshipTypeEndsWith(*relationship, "/worksheet")) {
            const auto target = resolveOwnerReference(workbookPart, id, "/worksheet",
                "Workbook sheet" + (name.empty() ? std::string{} : " '" + name + "'"));
            if (target.empty()) continue;
            ++result.inventory.worksheets;
            worksheetParts.push_back(target);
        } else if (relationshipTypeEndsWith(*relationship, "/chartsheet")) {
            const auto target = resolveOwnerReference(workbookPart, id, "/chartsheet",
                "Workbook chartsheet" + (name.empty() ? std::string{} : " '" + name + "'"));
            if (target.empty()) continue;
            ++result.inventory.chartsheets;
            chartsheetParts.push_back(target);
        } else {
            result.errors.push_back("Workbook sheet" + (name.empty() ? std::string{} : " '" + name + "'") + " has unsupported relationship type " + relationship->type);
        }
    }

    std::unordered_set<std::string> externalLinkParts;
    for (const auto& node : elementsByLocalName(workbookXml, "externalReference")) {
        const auto id = attributeByLocalName(node, "id");
        const auto target = resolveOwnerReference(workbookPart, id, "/externalLink", "Workbook externalReference");
        if (target.empty() || !externalLinkParts.insert(target).second) continue;
        ++result.inventory.externalLinks;

        const auto externalXml = archive.get(target);
        for (const auto& relationshipId : relationshipReferenceIds(externalXml)) {
            const auto* relationship = relationshipById(relationships, target, relationshipId);
            if (!relationship) {
                result.errors.push_back("External link " + target + " references missing relationship " + relationshipId);
                continue;
            }
            consumedObjectRelationships.insert(relationshipKey(target, relationshipId));
            if (relationship->targetMode != "External") {
                const auto resolved = RelationshipGraph::resolveTarget(target, relationship->target);
                if (resolved.empty() || !entries.count(resolved))
                    result.errors.push_back("External link " + target + " relationship " + relationshipId +
                                            " resolves to missing part " + resolved);
            }
        }
    }

    for (const auto& cachePart : pivotCacheParts) {
        const auto cacheXml = archive.get(cachePart);
        const auto definitions = elementsByLocalName(cacheXml, "pivotCacheDefinition");
        if (definitions.empty()) {
            result.errors.push_back("Pivot cache part " + cachePart + " has no pivotCacheDefinition root");
            continue;
        }
        const auto recordsId = attributeByLocalName(definitions.front(), "id");
        if (!recordsId.empty())
            (void)resolveOwnerReference(cachePart, recordsId, "/pivotCacheRecords", "Pivot cache " + cachePart);
    }

    std::unordered_set<std::string> drawingParts;
    std::unordered_set<std::string> pivotTableParts;
    const auto inspectPivotTable = [&](const std::string& worksheetPart,
                                       const std::string& relationshipId,
                                       const std::string& ownerLabel) {
        const auto target = resolveOwnerReference(worksheetPart, relationshipId, "/pivotTable", ownerLabel);
        if (target.empty() || !pivotTableParts.insert(target).second) return;
        ++result.inventory.pivotTables;
        const auto pivotXml = archive.get(target);
        const auto definitions = elementsByLocalName(pivotXml, "pivotTableDefinition");
        if (definitions.empty()) {
            result.errors.push_back("Pivot table part " + target + " has no pivotTableDefinition root");
            return;
        }
        const auto cacheId = attributeByLocalName(definitions.front(), "cacheId");
        const auto cacheIt = pivotCachePartById.find(cacheId);
        if (cacheId.empty() || cacheIt == pivotCachePartById.end()) {
            result.errors.push_back("Pivot table " + target + " references unknown cacheId " + cacheId);
            return;
        }

        const xlpp::PreservedRelationship* matchingCacheRelationship = nullptr;
        std::size_t cacheRelationshipCount = 0;
        for (const auto& relationship : relationships) {
            if (relationship.sourcePart != target || !relationshipTypeEndsWith(relationship, "/pivotCacheDefinition")) continue;
            ++cacheRelationshipCount;
            const auto resolved = relationship.targetMode == "External"
                ? std::string{} : RelationshipGraph::resolveTarget(target, relationship.target);
            if (resolved == cacheIt->second) matchingCacheRelationship = &relationship;
        }
        if (!matchingCacheRelationship) {
            result.errors.push_back("Pivot table " + target + " cacheId " + cacheId +
                                    " is not linked to workbook cache part " + cacheIt->second);
        } else {
            consumedObjectRelationships.insert(relationshipKey(target, matchingCacheRelationship->id));
        }
        if (cacheRelationshipCount > 1)
            result.errors.push_back("Pivot table " + target + " has multiple pivotCacheDefinition relationships");
    };

    std::unordered_set<std::string> headerFooterDrawingParts;
    std::unordered_set<std::string> printerSettingsParts;
    for (const auto& chartsheetPart : chartsheetParts) {
        if (!archive.contains(chartsheetPart)) continue;
        const auto chartsheetXml = archive.get(chartsheetPart);
        const auto drawings = elementsByLocalName(chartsheetXml, "drawing");
        if (drawings.empty()) {
            result.errors.push_back("Chartsheet " + chartsheetPart + " has no drawing owner");
            continue;
        }
        for (const auto& node : drawings) {
            const auto target = resolveOwnerReference(chartsheetPart, attributeByLocalName(node, "id"),
                                                      "/drawing", "Chartsheet drawing in " + chartsheetPart);
            if (!target.empty()) drawingParts.insert(target);
        }
        for (const auto& node : elementsByLocalName(chartsheetXml, "legacyDrawingHF")) {
            const auto target = resolveOwnerReference(chartsheetPart, attributeByLocalName(node, "id"),
                                                      "/vmlDrawing", "Chartsheet legacyDrawingHF in " + chartsheetPart);
            if (!target.empty() && headerFooterDrawingParts.insert(target).second)
                ++result.inventory.headerFooterDrawings;
        }
        for (const auto& node : elementsByLocalName(chartsheetXml, "pageSetup")) {
            const auto relationshipId = attributeByLocalName(node, "id");
            if (relationshipId.empty()) continue;
            const auto target = resolveOwnerReference(chartsheetPart, relationshipId,
                                                      "/printerSettings", "Chartsheet pageSetup in " + chartsheetPart);
            if (!target.empty() && printerSettingsParts.insert(target).second)
                ++result.inventory.printerSettings;
        }
    }

    for (const auto& worksheetPart : worksheetParts) {
        if (!archive.contains(worksheetPart)) continue;
        const auto worksheetXml = archive.get(worksheetPart);
        for (const auto& node : elementsByLocalName(worksheetXml, "drawing")) {
            const auto target = resolveOwnerReference(worksheetPart, attributeByLocalName(node, "id"),
                                                      "/drawing", "Worksheet drawing in " + worksheetPart);
            if (!target.empty()) drawingParts.insert(target);
        }
        for (const auto& node : elementsByLocalName(worksheetXml, "legacyDrawing"))
            (void)resolveOwnerReference(worksheetPart, attributeByLocalName(node, "id"),
                                        "/vmlDrawing", "Worksheet legacyDrawing in " + worksheetPart);
        for (const auto& node : elementsByLocalName(worksheetXml, "legacyDrawingHF"))
            (void)resolveOwnerReference(worksheetPart, attributeByLocalName(node, "id"),
                                        "/vmlDrawing", "Worksheet legacyDrawingHF in " + worksheetPart);
        for (const auto& node : elementsByLocalName(worksheetXml, "tablePart")) {
            const auto tablePart = resolveOwnerReference(worksheetPart, attributeByLocalName(node, "id"),
                                                         "/table", "Worksheet tablePart in " + worksheetPart);
            if (tablePart.empty()) continue;
            const auto tableXml = archive.get(tablePart);
            if (elementsByLocalName(tableXml, "table").empty())
                result.errors.push_back("Table part " + tablePart + " has no table root");
            else
                ++result.inventory.tables;
        }
        // Legacy comments are relationship-owned rather than referenced by an
        // r:id node in worksheet XML. Consume the worksheet comments relation
        // implicitly, while validating that it targets a comments part.
        for (const auto& relationship : relationships) {
            if (relationship.sourcePart != worksheetPart || !relationshipTypeEndsWith(relationship, "/comments")) continue;
            consumedObjectRelationships.insert(relationshipKey(worksheetPart, relationship.id));
            if (relationship.targetMode == "External") {
                result.errors.push_back("Worksheet comments relationship " + relationship.id + " unexpectedly targets an external resource");
                continue;
            }
            const auto commentsPart = RelationshipGraph::resolveTarget(worksheetPart, relationship.target);
            if (commentsPart.empty() || !entries.count(commentsPart)) {
                result.errors.push_back("Worksheet comments relationship " + relationship.id + " resolves to missing part " + commentsPart);
                continue;
            }
            const auto commentsXml = archive.get(commentsPart);
            if (elementsByLocalName(commentsXml, "comments").empty()) {
                result.errors.push_back("Comments part " + commentsPart + " has no comments root");
                continue;
            }
            result.inventory.comments += elementsByLocalName(commentsXml, "comment").size();
        }
        for (const auto& node : elementsByLocalName(worksheetXml, "pivotTablePart"))
            inspectPivotTable(worksheetPart, attributeByLocalName(node, "id"),
                              "Worksheet pivotTablePart in " + worksheetPart);

        // LibreOffice can export a worksheet pivot relationship without the
        // optional pivotTableParts owner container.  Treat that relationship as
        // an implicit owner reference when the pivot/cache chain is otherwise
        // complete, while still requiring a valid cacheId cross-link.
        for (const auto& relationship : relationships) {
            if (relationship.sourcePart != worksheetPart || !relationshipTypeEndsWith(relationship, "/pivotTable")) continue;
            if (consumedObjectRelationships.count(relationshipKey(worksheetPart, relationship.id))) continue;
            inspectPivotTable(worksheetPart, relationship.id,
                              "Worksheet implicit pivot relationship in " + worksheetPart);
        }
    }

    result.inventory.drawings = drawingParts.size();
    for (const auto& drawingPart : drawingParts) {
        const auto drawingXml = archive.get(drawingPart);
        const auto genericReferences = relationshipReferenceIds(drawingXml);
        for (const auto& relationshipId : genericReferences) {
            const auto* relationship = relationshipById(relationships, drawingPart, relationshipId);
            if (!relationship) {
                result.errors.push_back("Drawing " + drawingPart + " references missing relationship " + relationshipId);
                continue;
            }
            consumedObjectRelationships.insert(relationshipKey(drawingPart, relationshipId));
        }

        const auto shapes = elementsByLocalName(drawingXml, "sp");
        result.inventory.shapes += shapes.size();
        for (const auto& shape : shapes)
            if (!elementsByLocalName(shape, "txBody").empty()) ++result.inventory.textBoxes;
        result.inventory.connectors += elementsByLocalName(drawingXml, "cxnSp").size();
        result.inventory.groups += elementsByLocalName(drawingXml, "grpSp").size();
        result.inventory.otherDrawingObjects += elementsByLocalName(drawingXml, "contentPart").size();

        for (const auto& frame : elementsByLocalName(drawingXml, "graphicFrame"))
            if (elementsByLocalName(frame, "chart").empty()) ++result.inventory.otherDrawingObjects;

        for (const auto& picture : elementsByLocalName(drawingXml, "pic")) {
            const auto blips = elementsByLocalName(picture, "blip");
            if (blips.empty()) {
                result.errors.push_back("Picture in " + drawingPart + " has no blip reference");
                continue;
            }
            const auto embedId = attributeByLocalName(blips.front(), "embed");
            if (!embedId.empty()) {
                if (!resolveOwnerReference(drawingPart, embedId, "/image", "Picture in " + drawingPart).empty())
                    ++result.inventory.images;
                continue;
            }
            const auto linkId = attributeByLocalName(blips.front(), "link");
            if (!linkId.empty()) {
                if (!resolveOwnerReference(drawingPart, linkId, "/image", "Linked picture in " + drawingPart, true).empty())
                    ++result.inventory.images;
                continue;
            }
            result.errors.push_back("Picture in " + drawingPart + " has neither embed nor link relationship");
        }
        for (const auto& chart : elementsByLocalName(drawingXml, "chart")) {
            const auto target = resolveOwnerReference(drawingPart, attributeByLocalName(chart, "id"),
                                                      "/chart", "Chart in " + drawingPart);
            if (!target.empty()) ++result.inventory.charts;
        }
    }

    const auto isObjectRelationship = [](const xlpp::PreservedRelationship& relationship) {
        static constexpr std::array<std::string_view, 12> suffixes{
            "/drawing", "/image", "/chart", "/table", "/comments", "/vmlDrawing", "/printerSettings",
            "/externalLink", "/externalLinkPath", "/pivotTable", "/pivotCacheDefinition", "/pivotCacheRecords"
        };
        return std::any_of(suffixes.begin(), suffixes.end(), [&](const auto suffix) {
            return relationshipTypeEndsWith(relationship, suffix);
        });
    };
    for (const auto& relationship : relationships) {
        const bool drawingOwnedRelationship = drawingParts.count(relationship.sourcePart) != 0;
        if (!drawingOwnedRelationship && !isObjectRelationship(relationship)) continue;
        if (consumedObjectRelationships.count(relationshipKey(relationship.sourcePart, relationship.id))) continue;
        result.errors.push_back((drawingOwnedRelationship ? "Drawing relationship " : "Object relationship ") +
                                relationshipLabel(relationship) + " is not referenced by its owner XML");
    }

    std::sort(result.errors.begin(), result.errors.end());
    result.errors.erase(std::unique(result.errors.begin(), result.errors.end()), result.errors.end());
    return result;
}

} // namespace

RelationshipGraph RelationshipGraph::fromArchive(const ZipArchive& archive) {
    RelationshipGraph graph;
    for (const auto& name : archive.entryNames()) graph.entries_.insert(name);
    if (archive.contains("[Content_Types].xml"))
        graph.contentTypesXml_ = archive.get("[Content_Types].xml");
    for (const auto& name : archive.entryNames()) {
        if (!isRelationshipsPart(name)) continue;
        const auto source = sourcePartForRelationshipsPart(name);
        auto parsed = parseRelationshipsXml(source, archive.get(name));
        graph.relationships_.insert(graph.relationships_.end(),
                                    std::make_move_iterator(parsed.begin()),
                                    std::make_move_iterator(parsed.end()));
    }
    const auto inspection = inspectObjectGraph(archive, graph.relationships_, graph.entries_);
    graph.objectInventory_ = inspection.inventory;
    graph.ownerReferenceErrors_ = inspection.errors;
    return graph;
}

std::vector<xlpp::PreservedRelationship> RelationshipGraph::relationshipsFrom(const std::string& sourcePart) const {
    std::vector<xlpp::PreservedRelationship> result;
    for (const auto& relationship : relationships_)
        if (relationship.sourcePart == sourcePart) result.push_back(relationship);
    return result;
}

RelationshipValidationReport RelationshipGraph::validate() const {
    RelationshipValidationReport report;
    report.ownerReferenceErrors = ownerReferenceErrors_;

    std::unordered_map<std::string, std::unordered_set<std::string>> idsBySource;
    for (const auto& relationship : relationships_) {
        const auto sourceLabel = relationship.sourcePart.empty() ? std::string("/") : relationship.sourcePart;
        if (relationship.id.empty())
            report.relationshipSyntaxErrors.push_back(sourceLabel + " has a Relationship with an empty Id");
        if (relationship.type.empty())
            report.relationshipSyntaxErrors.push_back(sourceLabel + "#" + relationship.id + " has an empty Type");
        if (relationship.target.empty())
            report.relationshipSyntaxErrors.push_back(sourceLabel + "#" + relationship.id + " has an empty Target");
        if (!relationship.targetMode.empty() && relationship.targetMode != "Internal" && relationship.targetMode != "External")
            report.relationshipSyntaxErrors.push_back(sourceLabel + "#" + relationship.id +
                                                      " has invalid TargetMode " + relationship.targetMode);
        auto& ids = idsBySource[relationship.sourcePart];
        if (!ids.insert(relationship.id).second)
            report.duplicateRelationshipIds.push_back(
                (relationship.sourcePart.empty() ? std::string("/") : relationship.sourcePart) +
                " has duplicate relationship ID " + relationship.id);
        if (relationship.targetMode == "External") continue;
        const auto resolved = resolveTarget(relationship.sourcePart, relationship.target);
        if (resolved.empty() || !entries_.count(resolved))
            report.danglingRelationships.push_back(relationshipLabel(relationship) + " (resolved: " + resolved + ")");
    }

    std::unordered_map<std::string, std::vector<const xlpp::PreservedRelationship*>> outgoing;
    for (const auto& relationship : relationships_)
        if (relationship.targetMode != "External") outgoing[relationship.sourcePart].push_back(&relationship);

    std::unordered_set<std::string> reachable;
    std::deque<std::string> queue;
    queue.push_back({});
    while (!queue.empty()) {
        const auto source = queue.front();
        queue.pop_front();
        const auto it = outgoing.find(source);
        if (it == outgoing.end()) continue;
        for (const auto* relationship : it->second) {
            const auto target = resolveTarget(source, relationship->target);
            if (!entries_.count(target)) continue;
            if (reachable.insert(target).second) queue.push_back(target);
        }
    }

    for (const auto& entry : entries_) {
        if (isInfrastructurePart(entry)) continue;
        if (!reachable.count(entry)) report.orphanedParts.push_back(entry);
    }

    if (contentTypesXml_.empty()) {
        report.contentTypeErrors.push_back("[Content_Types].xml is missing or empty");
    } else {
        std::unordered_map<std::string, std::string> defaults;
        std::unordered_map<std::string, std::string> overrides;
        for (const auto& node : elementsByLocalName(contentTypesXml_, "Default")) {
            const auto extension = attributeByLocalName(node, "Extension");
            const auto contentType = attributeByLocalName(node, "ContentType");
            if (extension.empty() || contentType.empty()) {
                report.contentTypeErrors.push_back("Malformed Default content-type declaration");
                continue;
            }
            if (!defaults.emplace(extension, contentType).second)
                report.contentTypeErrors.push_back("Duplicate Default content type for extension: " + extension);
        }
        for (const auto& node : elementsByLocalName(contentTypesXml_, "Override")) {
            auto partName = attributeByLocalName(node, "PartName");
            const auto contentType = attributeByLocalName(node, "ContentType");
            if (!partName.empty() && partName.front() == '/') partName.erase(partName.begin());
            if (partName.empty() || contentType.empty()) {
                report.contentTypeErrors.push_back("Malformed Override content-type declaration");
                continue;
            }
            if (!overrides.emplace(partName, contentType).second)
                report.contentTypeErrors.push_back("Duplicate Override content type for part: " + partName);
            if (!entries_.count(partName))
                report.contentTypeErrors.push_back("Content-type Override points to a missing part: " + partName);
        }
        for (const auto& entry : entries_) {
            if (entry == "[Content_Types].xml") continue;
            if (overrides.count(entry)) continue;
            const auto slash = entry.find_last_of('/');
            const auto dot = entry.find_last_of('.');
            const auto extension = dot == std::string::npos || (slash != std::string::npos && dot < slash)
                ? std::string{} : entry.substr(dot + 1);
            if (extension.empty() || !defaults.count(extension))
                report.contentTypeErrors.push_back("No content type declared for part: " + entry);
        }
    }
    std::sort(report.relationshipSyntaxErrors.begin(), report.relationshipSyntaxErrors.end());
    report.relationshipSyntaxErrors.erase(std::unique(report.relationshipSyntaxErrors.begin(), report.relationshipSyntaxErrors.end()),
                                          report.relationshipSyntaxErrors.end());
    std::sort(report.duplicateRelationshipIds.begin(), report.duplicateRelationshipIds.end());
    std::sort(report.danglingRelationships.begin(), report.danglingRelationships.end());
    std::sort(report.orphanedParts.begin(), report.orphanedParts.end());
    std::sort(report.contentTypeErrors.begin(), report.contentTypeErrors.end());
    std::sort(report.ownerReferenceErrors.begin(), report.ownerReferenceErrors.end());
    report.contentTypeErrors.erase(std::unique(report.contentTypeErrors.begin(), report.contentTypeErrors.end()),
                                   report.contentTypeErrors.end());
    report.ownerReferenceErrors.erase(std::unique(report.ownerReferenceErrors.begin(), report.ownerReferenceErrors.end()),
                                      report.ownerReferenceErrors.end());
    return report;
}

std::string RelationshipGraph::sourcePartForRelationshipsPart(const std::string& relationshipsPart) {
    if (relationshipsPart == "_rels/.rels") return {};
    const auto marker = relationshipsPart.rfind("/_rels/");
    if (marker == std::string::npos || relationshipsPart.size() < 5) return {};
    const auto directory = relationshipsPart.substr(0, marker + 1);
    auto file = relationshipsPart.substr(marker + 7);
    if (file.size() >= 5 && file.rfind(".rels") == file.size() - 5) file.resize(file.size() - 5);
    return directory + file;
}

std::string RelationshipGraph::relationshipsPartForSource(const std::string& sourcePart) {
    if (sourcePart.empty()) return "_rels/.rels";
    const auto slash = sourcePart.find_last_of('/');
    if (slash == std::string::npos) return "_rels/" + sourcePart + ".rels";
    return sourcePart.substr(0, slash + 1) + "_rels/" + sourcePart.substr(slash + 1) + ".rels";
}

std::string RelationshipGraph::resolveTarget(const std::string& sourcePart, const std::string& target) {
    if (target.empty()) return {};
    // OPC part names use forward slashes. A backslash or a URI scheme in an
    // Internal relationship is malformed and must not be silently converted
    // into a package path.
    if (target.find('\\') != std::string::npos) return {};
    // Reject any RFC-style URI scheme, not only the common "://" spelling.
    // Values such as "mailto:..." or "file:..." are External targets and
    // must never be interpreted as paths inside the ZIP package.
    const auto colon = target.find(':');
    if (colon != std::string::npos && colon != 0) {
        const auto slash = target.find('/');
        bool scheme = slash == std::string::npos || colon < slash;
        if (scheme) {
            const auto alpha = [](unsigned char ch) noexcept {
                return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
            };
            const auto schemeChar = [&](unsigned char ch) noexcept {
                return alpha(ch) || (ch >= '0' && ch <= '9') || ch == '+' || ch == '-' || ch == '.';
            };
            scheme = alpha(static_cast<unsigned char>(target.front()));
            for (std::size_t i = 1; scheme && i < colon; ++i)
                scheme = schemeChar(static_cast<unsigned char>(target[i]));
            if (scheme) return {};
        }
    }

    std::string combined;
    if (target.front() == '/') {
        combined = target.substr(1);
    } else {
        const auto slash = sourcePart.find_last_of('/');
        combined = (slash == std::string::npos ? std::string{} : sourcePart.substr(0, slash + 1)) + target;
    }

    std::vector<std::string> segments;
    std::size_t begin = 0;
    while (begin <= combined.size()) {
        const auto end = combined.find('/', begin);
        const auto segment = combined.substr(begin, end == std::string::npos ? std::string::npos : end - begin);
        if (segment == "..") {
            // Escaping above the package root is never a valid internal part URI.
            if (segments.empty()) return {};
            segments.pop_back();
        } else if (!segment.empty() && segment != ".") {
            segments.push_back(segment);
        }
        if (end == std::string::npos) break;
        begin = end + 1;
    }
    if (segments.empty()) return {};

    std::ostringstream result;
    for (std::size_t i = 0; i < segments.size(); ++i) {
        if (i) result << '/';
        result << segments[i];
    }
    return result.str();
}

std::vector<xlpp::PreservedRelationship> RelationshipGraph::parseRelationshipsXml(
    const std::string& sourcePart, const std::string& xml) {
    std::vector<xlpp::PreservedRelationship> result;
    for (const auto& node : elementsByLocalName(xml, "Relationship")) {
        xlpp::PreservedRelationship relationship;
        relationship.sourcePart = sourcePart;
        relationship.id = attributeByLocalName(node, "Id");
        relationship.type = attributeByLocalName(node, "Type");
        relationship.target = attributeByLocalName(node, "Target");
        relationship.targetMode = attributeByLocalName(node, "TargetMode");
        result.push_back(std::move(relationship));
    }
    return result;
}

std::string RelationshipGraph::serializeRelationships(
    const std::vector<xlpp::PreservedRelationship>& relationships, bool strictNamespace) {
    const char* ns = strictNamespace
        ? "http://purl.oclc.org/ooxml/package/relationships"
        : "http://schemas.openxmlformats.org/package/2006/relationships";
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << ns << "\">";
    for (const auto& relationship : relationships) {
        xml << "<Relationship Id=\"" << xmlEscapeLocal(relationship.id)
            << "\" Type=\"" << xmlEscapeLocal(relationship.type)
            << "\" Target=\"" << xmlEscapeLocal(relationship.target) << "\"";
        if (!relationship.targetMode.empty())
            xml << " TargetMode=\"" << xmlEscapeLocal(relationship.targetMode) << "\"";
        xml << "/>";
    }
    xml << "</Relationships>";
    return xml.str();
}

PackageDiffReport comparePackages(const ZipArchive& before, const ZipArchive& after) {
    PackageDiffReport report;
    const auto beforeNames = before.entryNames();
    const auto afterNames = after.entryNames();
    const std::unordered_set<std::string> beforeSet(beforeNames.begin(), beforeNames.end());
    const std::unordered_set<std::string> afterSet(afterNames.begin(), afterNames.end());
    for (const auto& name : afterNames) {
        if (!beforeSet.count(name)) report.addedParts.push_back(name);
        else if (before.get(name) != after.get(name)) report.changedParts.push_back(name);
    }
    for (const auto& name : beforeNames)
        if (!afterSet.count(name)) report.removedParts.push_back(name);
    std::sort(report.addedParts.begin(), report.addedParts.end());
    std::sort(report.removedParts.begin(), report.removedParts.end());
    std::sort(report.changedParts.begin(), report.changedParts.end());
    const auto beforeGraph = RelationshipGraph::fromArchive(before);
    const auto afterGraph = RelationshipGraph::fromArchive(after);
    report.beforeValidation = beforeGraph.validate();
    report.afterValidation = afterGraph.validate();
    report.beforeObjects = beforeGraph.objectInventory();
    report.afterObjects = afterGraph.objectInventory();
    const auto addRegression = [&](const char* label, std::size_t beforeCount, std::size_t afterCount) {
        if (afterCount < beforeCount)
            report.objectCountRegressions.push_back(std::string(label) + " decreased from " +
                std::to_string(beforeCount) + " to " + std::to_string(afterCount));
    };
    addRegression("Worksheet count", report.beforeObjects.worksheets, report.afterObjects.worksheets);
    addRegression("Chartsheet count", report.beforeObjects.chartsheets, report.afterObjects.chartsheets);
    addRegression("Header/footer drawing count", report.beforeObjects.headerFooterDrawings, report.afterObjects.headerFooterDrawings);
    addRegression("Printer settings count", report.beforeObjects.printerSettings, report.afterObjects.printerSettings);
    addRegression("Drawing count", report.beforeObjects.drawings, report.afterObjects.drawings);
    addRegression("Visible image count", report.beforeObjects.images, report.afterObjects.images);
    addRegression("Visible chart count", report.beforeObjects.charts, report.afterObjects.charts);
    addRegression("Shape count", report.beforeObjects.shapes, report.afterObjects.shapes);
    addRegression("Text box count", report.beforeObjects.textBoxes, report.afterObjects.textBoxes);
    addRegression("Connector count", report.beforeObjects.connectors, report.afterObjects.connectors);
    addRegression("Group count", report.beforeObjects.groups, report.afterObjects.groups);
    addRegression("Other drawing object count", report.beforeObjects.otherDrawingObjects, report.afterObjects.otherDrawingObjects);
    addRegression("Table count", report.beforeObjects.tables, report.afterObjects.tables);
    addRegression("Comment count", report.beforeObjects.comments, report.afterObjects.comments);
    addRegression("External link count", report.beforeObjects.externalLinks, report.afterObjects.externalLinks);
    addRegression("Pivot table count", report.beforeObjects.pivotTables, report.afterObjects.pivotTables);
    addRegression("Pivot cache count", report.beforeObjects.pivotCaches, report.afterObjects.pivotCaches);
    return report;
}

} // namespace xlpp::internal
