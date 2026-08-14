#include "WorkbookChartsheetPackage.h"
#include "WorkbookChartsheetIO.h"

#include <XLPP/Chart/Chartsheet.h>
#include <XLPP/Workbook/Workbook.h>
#include "../Packaging/RelationshipGraph.h"
#include "../Packaging/ZipArchive.h"
#include "../XML/XmlUtilities.h"

#include <algorithm>
#include <array>
#include <set>
#include <sstream>
#include <stdexcept>
#include <unordered_set>
#include <utility>

namespace xlpp::internal {

struct WorkbookChartsheetPackageAccess {
    static const std::string& sourcePart(const Chartsheet& sheet) noexcept { return sheet.sourcePart_; }
    static std::string& sourceXml(const Chartsheet& sheet) noexcept { return sheet.sourceXml_; }
    static const std::string& drawingRelationshipId(const Chartsheet& sheet) noexcept { return sheet.drawingRelationshipId_; }
    static const std::string& printerSettingsSourcePart(const Chartsheet& sheet) noexcept { return sheet.printerSettingsSourcePart_; }
    static const std::string& printerSettingsRelationshipId(const Chartsheet& sheet) noexcept { return sheet.printerSettingsRelationshipId_; }
    static const std::optional<std::string>& printerSettingsData(const Chartsheet& sheet) noexcept { return sheet.printerSettingsData_; }
    static bool printerSettingsDirty(const Chartsheet& sheet) noexcept { return sheet.printerSettingsDirty_; }
    static void markImportedSheetClean(const Chartsheet& sheet) noexcept { sheet.sheetDirty_ = false; }
    static void clearDirty(const Chartsheet& sheet) noexcept { sheet.clearDirty(); }
};

namespace {

std::string nsRelsPkg(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/package/relationships"
                  : "http://schemas.openxmlformats.org/package/2006/relationships";
}

std::string nsRelsDoc(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
                  : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
}

std::string resolvePackagePart(const std::string& basePart, std::string relativeTarget) {
    if (relativeTarget.empty()) return {};
    const bool absoluteTarget = relativeTarget.front() == '/';
    if (absoluteTarget) relativeTarget.erase(relativeTarget.begin());
    std::vector<std::string> segments;
    const auto slash = basePart.find_last_of('/');
    const std::string combined = absoluteTarget
        ? relativeTarget
        : (slash == std::string::npos ? std::string{} : basePart.substr(0, slash + 1)) + relativeTarget;
    std::size_t begin = 0;
    while (begin <= combined.size()) {
        const auto end = combined.find('/', begin);
        const auto segment = combined.substr(begin, end == std::string::npos ? std::string::npos : end - begin);
        if (segment == "..") { if (!segments.empty()) segments.pop_back(); }
        else if (!segment.empty() && segment != ".") segments.push_back(segment);
        if (end == std::string::npos) break;
        begin = end + 1;
    }
    std::ostringstream result;
    for (std::size_t i = 0; i < segments.size(); ++i) {
        if (i) result << '/';
        result << segments[i];
    }
    return result.str();
}

std::string relationshipKind(const PreservedRelationship& relationship) {
    const auto slash = relationship.type.find_last_of('/');
    if (slash != std::string::npos && slash + 1 < relationship.type.size())
        return relationship.type.substr(slash + 1);
    const auto& target = relationship.target;
    if (target.find("/drawings/") != std::string::npos || target.rfind("../drawings/", 0) == 0) return "drawing";
    if (target.find("/printerSettings/") != std::string::npos || target.rfind("../printerSettings/", 0) == 0) return "printerSettings";
    return {};
}

std::vector<PreservedRelationship> relationshipsForSource(
    const std::vector<PreservedRelationship>& relationships,
    const std::string& sourcePart) {
    std::vector<PreservedRelationship> result;
    for (const auto& relationship : relationships)
        if (relationship.sourcePart == sourcePart) result.push_back(relationship);
    return result;
}

bool sameRelationship(const PreservedRelationship& lhs, const PreservedRelationship& rhs) {
    return lhs.type == rhs.type && lhs.target == rhs.target && lhs.targetMode == rhs.targetMode;
}

std::string allocateRelationshipId(const std::set<std::string>& used) {
    for (std::size_t index = 1;; ++index) {
        const auto candidate = "rIdXLPP" + std::to_string(index);
        if (!used.count(candidate)) return candidate;
    }
}

bool replaceRelationshipReferenceInTag(std::string& ownerXml, std::string_view tagName,
                                       const std::string& oldId, const std::string& newId) {
    if (oldId == newId || oldId.empty()) return true;
    const auto tagStart = ownerXml.find("<" + std::string(tagName));
    if (tagStart == std::string::npos) return false;
    const auto tagEnd = ownerXml.find('>', tagStart);
    if (tagEnd == std::string::npos) return false;
    const std::array<std::string, 2> patterns{"r:id=\"" + oldId + "\"", "r:id='" + oldId + "'"};
    const std::array<std::string, 2> replacements{"r:id=\"" + newId + "\"", "r:id='" + newId + "'"};
    for (std::size_t p = 0; p < patterns.size(); ++p) {
        const auto position = ownerXml.find(patterns[p], tagStart);
        if (position != std::string::npos && position < tagEnd) {
            ownerXml.replace(position, patterns[p].size(), replacements[p]);
            return true;
        }
    }
    return false;
}

void replaceGeneratedRelationshipReference(std::string& ownerXml,
                                           const PreservedRelationship& relationship,
                                           const std::string& oldId,
                                           const std::string& newId) {
    const auto kind = relationshipKind(relationship);
    if (kind == "drawing") {
        if (!replaceRelationshipReferenceInTag(ownerXml, "drawing", oldId, newId))
            throw std::logic_error("Generated Chartsheet drawing relationship has no owner reference");
        return;
    }
    if (kind == "printerSettings") {
        if (!replaceRelationshipReferenceInTag(ownerXml, "pageSetup", oldId, newId))
            throw std::logic_error("Generated Chartsheet printerSettings relationship has no pageSetup owner reference");
        return;
    }
    throw std::logic_error("Unsupported generated Chartsheet relationship collision");
}

std::string mergeRelationshipsXml(
    const std::string& generatedXml,
    const std::vector<PreservedRelationship>& original,
    const std::function<bool(const PreservedRelationship&)>& preserve,
    bool strict,
    std::string* generatedOwnerXml) {
    auto generated = RelationshipGraph::parseRelationshipsXml({}, generatedXml);
    std::vector<PreservedRelationship> selected;
    for (const auto& relationship : original)
        if (preserve(relationship)) selected.push_back(relationship);

    std::set<std::string> used;
    for (const auto& relationship : selected) used.insert(relationship.id);
    for (const auto& relationship : generated) used.insert(relationship.id);

    for (auto& relationship : generated) {
        const auto collision = std::find_if(selected.begin(), selected.end(), [&](const auto& candidate) {
            return candidate.id == relationship.id && !sameRelationship(candidate, relationship);
        });
        if (collision == selected.end()) continue;
        const auto oldId = relationship.id;
        const auto newId = allocateRelationshipId(used);
        used.insert(newId);
        relationship.id = newId;
        if (generatedOwnerXml) replaceGeneratedRelationshipReference(*generatedOwnerXml, relationship, oldId, newId);
    }

    std::vector<PreservedRelationship> merged = std::move(generated);
    for (const auto& relationship : selected) {
        const auto duplicate = std::find_if(merged.begin(), merged.end(), [&](const auto& candidate) {
            return sameRelationship(candidate, relationship);
        });
        if (duplicate == merged.end()) merged.push_back(relationship);
    }
    return RelationshipGraph::serializeRelationships(merged, strict);
}

void suppressExclusivePartClosure(const std::string& rootPart,
                                  const std::vector<PreservedRelationship>& allRelationships,
                                  std::set<std::string>& suppressedPreservedParts) {
    std::unordered_set<std::string> closure;
    std::vector<std::string> stack{rootPart};
    while (!stack.empty()) {
        auto part = std::move(stack.back());
        stack.pop_back();
        if (part.empty() || !closure.insert(part).second) continue;
        for (const auto& relationship : allRelationships) {
            if (relationship.sourcePart != part || relationship.targetMode == "External") continue;
            const auto target = resolvePackagePart(part, relationship.target);
            if (!target.empty() && !closure.count(target)) stack.push_back(target);
        }
    }

    std::unordered_set<std::string> protectedParts;
    for (const auto& candidate : closure) {
        if (candidate == rootPart) continue;
        const bool externallyReferenced = std::any_of(allRelationships.begin(), allRelationships.end(), [&](const auto& relationship) {
            if (relationship.targetMode == "External" || closure.count(relationship.sourcePart)) return false;
            return resolvePackagePart(relationship.sourcePart, relationship.target) == candidate;
        });
        if (externallyReferenced) protectedParts.insert(candidate);
    }
    std::vector<std::string> protectStack(protectedParts.begin(), protectedParts.end());
    while (!protectStack.empty()) {
        auto part = std::move(protectStack.back());
        protectStack.pop_back();
        for (const auto& relationship : allRelationships) {
            if (relationship.sourcePart != part || relationship.targetMode == "External") continue;
            const auto target = resolvePackagePart(part, relationship.target);
            if (closure.count(target) && protectedParts.insert(target).second) protectStack.push_back(target);
        }
    }
    for (const auto& part : closure) {
        if (protectedParts.count(part)) continue;
        suppressedPreservedParts.insert(part);
        suppressedPreservedParts.insert(RelationshipGraph::relationshipsPartForSource(part));
    }
}

} // namespace

void writeChartsheetPackageParts(
    ZipArchive& archive,
    const std::deque<Chartsheet>& chartsheets,
    const std::vector<std::size_t>& chartsheetPartIds,
    const std::vector<std::size_t>& chartsheetDrawingIds,
    const std::vector<std::size_t>& chartsheetChartIds,
    const std::vector<std::size_t>& chartsheetPrinterSettingsIds,
    const std::vector<PreservedRelationship>& preservedRelationships,
    std::set<std::string>& suppressedPreservedParts,
    bool strict,
    const std::function<std::string(const Chart&, bool)>& chartSerializer) {

    for (std::size_t i = 0; i < chartsheets.size(); ++i) {
        auto& chartSheet = chartsheets[i];
        const auto originalRelationships = chartSheet.imported()
            ? relationshipsForSource(preservedRelationships, WorkbookChartsheetPackageAccess::sourcePart(chartSheet))
            : std::vector<PreservedRelationship>{};

        const auto suppressOldPrinterSettings = [&]() {
            const auto& source = WorkbookChartsheetPackageAccess::printerSettingsSourcePart(chartSheet);
            if (!source.empty()) suppressExclusivePartClosure(source, preservedRelationships, suppressedPreservedParts);
        };

        const auto generatedPrinterPartId = chartsheetPrinterSettingsIds.at(i);
        const bool generatePrinterSettings = chartSheet.hasPrinterSettings() && generatedPrinterPartId != 0;
        std::string generatedPrinterRelationshipId;
        std::string generatedPrinterTarget;
        if (generatePrinterSettings) {
            generatedPrinterRelationshipId = WorkbookChartsheetPackageAccess::printerSettingsRelationshipId(chartSheet).empty()
                ? "rIdPrinterSettings" : WorkbookChartsheetPackageAccess::printerSettingsRelationshipId(chartSheet);
            generatedPrinterTarget = "../printerSettings/printerSettings" + std::to_string(generatedPrinterPartId) + ".bin";
        }

        if (chartSheet.imported() && !chartSheet.chartDirty()) {
            if (!chartSheet.sheetDirty() && !WorkbookChartsheetPackageAccess::printerSettingsDirty(chartSheet)) continue;
            const auto& sourcePart = WorkbookChartsheetPackageAccess::sourcePart(chartSheet);
            if (sourcePart.empty()) throw std::logic_error("Imported chartsheet is missing its source part");

            auto sheetPartXml = serializeChartsheetXml(
                chartSheet, strict, WorkbookChartsheetPackageAccess::drawingRelationshipId(chartSheet), generatedPrinterRelationshipId);

            if (WorkbookChartsheetPackageAccess::printerSettingsDirty(chartSheet)) {
                suppressOldPrinterSettings();
                std::ostringstream generatedRelationships;
                generatedRelationships << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\""
                                       << nsRelsPkg(strict) << "\">";
                if (generatePrinterSettings)
                    generatedRelationships << "<Relationship Id=\"" << xmlEscape(generatedPrinterRelationshipId)
                                           << "\" Type=\"" << nsRelsDoc(strict) << "/printerSettings\" Target=\""
                                           << xmlEscape(generatedPrinterTarget) << "\"/>";
                generatedRelationships << "</Relationships>";
                const auto merged = mergeRelationshipsXml(
                    generatedRelationships.str(), originalRelationships,
                    [](const PreservedRelationship& relationship) { return relationshipKind(relationship) != "printerSettings"; },
                    strict, &sheetPartXml);
                archive.add(RelationshipGraph::relationshipsPartForSource(sourcePart), merged);
                if (generatePrinterSettings)
                    archive.add("xl/printerSettings/printerSettings" + std::to_string(generatedPrinterPartId) + ".bin",
                                *WorkbookChartsheetPackageAccess::printerSettingsData(chartSheet), false);
            }

            archive.add(sourcePart, sheetPartXml);
            WorkbookChartsheetPackageAccess::sourceXml(chartSheet) = std::move(sheetPartXml);
            continue;
        }

        if (!chartSheet.hasChart()) throw std::invalid_argument("Chartsheet has no chart: " + chartSheet.name());
        const auto partId = chartsheetPartIds.at(i);
        const auto drawingId = chartsheetDrawingIds.at(i);
        const auto chartId = chartsheetChartIds.at(i);
        if (partId == 0 || drawingId == 0 || chartId == 0)
            throw std::logic_error("Chartsheet package IDs are incomplete");

        if (chartSheet.imported() && !WorkbookChartsheetPackageAccess::sourcePart(chartSheet).empty()) {
            const auto& sourcePart = WorkbookChartsheetPackageAccess::sourcePart(chartSheet);
            suppressedPreservedParts.insert(sourcePart);
            suppressedPreservedParts.insert(RelationshipGraph::relationshipsPartForSource(sourcePart));
            for (const auto& relationship : originalRelationships) {
                if (relationshipKind(relationship) != "drawing" || relationship.targetMode == "External") continue;
                const auto drawingPart = resolvePackagePart(sourcePart, relationship.target);
                if (!drawingPart.empty()) suppressExclusivePartClosure(drawingPart, preservedRelationships, suppressedPreservedParts);
            }
            if (WorkbookChartsheetPackageAccess::printerSettingsDirty(chartSheet)) suppressOldPrinterSettings();
        }

        std::string sheetPartXml = serializeChartsheetXml(chartSheet, strict, "rId1", generatedPrinterRelationshipId);
        auto generatedRelationships = serializeChartsheetRelationshipsXml(drawingId, strict);
        if (generatePrinterSettings) {
            const auto close = generatedRelationships.rfind("</Relationships>");
            if (close == std::string::npos) throw std::logic_error("Invalid generated Chartsheet relationships XML");
            generatedRelationships.insert(close, "<Relationship Id=\"" + xmlEscape(generatedPrinterRelationshipId)
                + "\" Type=\"" + nsRelsDoc(strict) + "/printerSettings\" Target=\""
                + xmlEscape(generatedPrinterTarget) + "\"/>");
        }
        const auto mergedRelationships = mergeRelationshipsXml(
            generatedRelationships, originalRelationships,
            [&](const PreservedRelationship& relationship) {
                const auto kind = relationshipKind(relationship);
                if (kind == "drawing") return false;
                if (kind == "printerSettings" && WorkbookChartsheetPackageAccess::printerSettingsDirty(chartSheet)) return false;
                return true;
            }, strict, &sheetPartXml);

        archive.add("xl/chartsheets/sheet" + std::to_string(partId) + ".xml", std::move(sheetPartXml));
        archive.add("xl/chartsheets/_rels/sheet" + std::to_string(partId) + ".xml.rels", mergedRelationships);
        archive.add("xl/drawings/drawing" + std::to_string(drawingId) + ".xml", serializeChartsheetDrawingXml(strict));
        archive.add("xl/drawings/_rels/drawing" + std::to_string(drawingId) + ".xml.rels", serializeChartsheetDrawingRelationshipsXml(chartId, strict));
        archive.add("xl/charts/chart" + std::to_string(chartId) + ".xml", chartSerializer(chartSheet.chart(), strict));
        if (generatePrinterSettings)
            archive.add("xl/printerSettings/printerSettings" + std::to_string(generatedPrinterPartId) + ".bin",
                        *WorkbookChartsheetPackageAccess::printerSettingsData(chartSheet), false);

        if (chartSheet.imported()) WorkbookChartsheetPackageAccess::markImportedSheetClean(chartSheet);
        else WorkbookChartsheetPackageAccess::clearDirty(chartSheet);
    }
}

} // namespace xlpp::internal
