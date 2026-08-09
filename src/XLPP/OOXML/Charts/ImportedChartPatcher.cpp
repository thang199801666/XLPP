#include "OOXML/Charts/ImportedChartPatcher.h"
#include "OOXML/Charts/ImportedChartXmlEditor.h"
#include "OOXML/Charts/ChartSerializer.h"
#include "OOXML/Common/Namespaces.h"
#include "OOXML/Common/PackageRelationships.h"
#include "OOXML/Drawings/DrawingPackageSupport.h"
#include "OOXML/Drawings/WorkbookDrawingAccess.h"
#include "Package/Zip/ZipArchive.h"
#include "Package/Xml/XmlUtilities.h"
#include "Package/Opc/RelationshipGraph.h"
#include "Preservation/PartPolicy.h"
#include "Preservation/PackageClosure.h"

#include <algorithm>
#include <cstddef>
#include <set>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

using namespace xlpp::internal::ooxml::drawing_support;

namespace xlpp::internal::ooxml {

bool applyChartChangesToPreservedDrawing(xlpp::internal::ZipArchive& z,
                                         const xlpp::Worksheet& sheet,
                                         const std::string& sourceSheetPart,
                                         const std::string& sourceSheetXml,
                                         const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                         const std::vector<xlpp::PreservedPart>& preservedParts,
                                         std::size_t& nextChartId,
                                         std::set<std::string>& suppressedPreservedParts) {
    const auto& edits = xlpp::internal::WorkbookDrawingAccess::chartEdits(sheet);
    if (sheet.appendedChartCount() == 0 && edits.empty()) return true;
    if (sourceSheetPart.empty() || sourceSheetXml.empty()) return false;

    const auto sheetRelationships = relationshipsForSource(allRelationships, sourceSheetPart);
    std::vector<std::string> ownedDrawingParts;
    std::unordered_set<std::string> seenDrawingParts;
    for (const auto& drawingNode : xlpp::internal::tags(sourceSheetXml, "drawing")) {
        const auto relationshipId = xlpp::internal::attribute(drawingNode, "r:id");
        const auto relationship = std::find_if(sheetRelationships.begin(), sheetRelationships.end(), [&](const auto& candidate) {
            return candidate.id == relationshipId && relationshipKind(candidate) == "drawing"
                && candidate.targetMode != "External";
        });
        if (relationship == sheetRelationships.end()) continue;
        const auto drawingPart = resolvePackagePart(sourceSheetPart, relationship->target);
        if (!drawingPart.empty() && seenDrawingParts.insert(drawingPart).second) ownedDrawingParts.push_back(drawingPart);
    }
    if (ownedDrawingParts.empty()) return false;
    for (const auto& edit : edits)
        if (!seenDrawingParts.count(edit.sourceDrawingPart)) return false;

    std::string appendDrawingPart = ownedDrawingParts.front();
    const auto& charts = static_cast<const xlpp::Worksheet&>(sheet).charts();
    for (std::size_t index = 0; index < std::min(sheet.loadedChartCount(), charts.size()); ++index) {
        if (seenDrawingParts.count(charts[index].sourceDrawingPart())) {
            appendDrawingPart = charts[index].sourceDrawingPart();
            break;
        }
    }
    if (sheet.loadedChartCount() == 0) {
        const auto& images = static_cast<const xlpp::Worksheet&>(sheet).images();
        for (std::size_t index = 0; index < std::min(sheet.loadedImageCount(), images.size()); ++index) {
            if (seenDrawingParts.count(images[index].sourceDrawingPart())) {
                appendDrawingPart = images[index].sourceDrawingPart();
                break;
            }
        }
    }

    for (const auto& drawingPart : ownedDrawingParts) {
        const bool appendHere = sheet.appendedChartCount() > 0 && drawingPart == appendDrawingPart;
        const bool hasEditsHere = std::any_of(edits.begin(), edits.end(), [&](const auto& edit) {
            return edit.sourceDrawingPart == drawingPart;
        });
        if (!appendHere && !hasEditsHere) continue;

        std::string drawingXmlText;
        if (z.contains(drawingPart)) drawingXmlText = z.get(drawingPart);
        else {
            const auto* raw = findPreservedPart(preservedParts, drawingPart);
            if (!raw) return false;
            drawingXmlText = raw->data;
        }
        const bool sourceStrict = drawingXmlText.find("http://purl.oclc.org/ooxml/drawingml/") != std::string::npos;
        const auto drawingRelsPart = xlpp::internal::RelationshipGraph::relationshipsPartForSource(drawingPart);
        auto drawingRelationships = z.contains(drawingRelsPart)
            ? xlpp::internal::RelationshipGraph::parseRelationshipsXml(drawingPart, z.get(drawingRelsPart))
            : relationshipsForSource(allRelationships, drawingPart);
        bool relationshipsChanged = false;

        std::unordered_map<std::string, std::string> chartWorkingCopies;
        for (std::size_t editIndex = 0; editIndex < edits.size(); ++editIndex) {
            const auto& edit = edits[editIndex];
            if (edit.sourceDrawingPart != drawingPart) continue;
            auto relationship = std::find_if(drawingRelationships.begin(), drawingRelationships.end(), [&](const auto& candidate) {
                return candidate.id == edit.sourceRelationshipId;
            });
            if (relationship == drawingRelationships.end()) return false;

            if (edit.removed) {
                if (!removeImportedChartAnchor(drawingXmlText, edit.stableId, edit.sourceRelationshipId)) return false;
                if (!drawingReferencesChartRelationship(drawingXmlText, edit.sourceRelationshipId)) {
                    const bool shared = std::any_of(allRelationships.begin(), allRelationships.end(), [&](const auto& candidate) {
                        if (candidate.sourcePart == drawingPart && candidate.id == edit.sourceRelationshipId) return false;
                        return candidate.targetMode != "External" &&
                            resolvePackagePart(candidate.sourcePart, candidate.target) == edit.sourceChartPart;
                    });
                    drawingRelationships.erase(relationship);
                    relationshipsChanged = true;
                    if (!shared) xlpp::internal::preservation::suppressExclusivePartClosure(edit.sourceChartPart, allRelationships, suppressedPreservedParts);
                }
                continue;
            }

            if ((edit.moved || edit.resized) &&
                !patchImportedChartAnchor(drawingXmlText, edit.stableId, edit.sourceRelationshipId,
                                          edit.anchor, edit.moved, edit.resized)) return false;

            if (!importedChartEditRequiresXml(sheet, editIndex)) continue;
            auto chartIt = chartWorkingCopies.find(edit.sourceChartPart);
            if (chartIt == chartWorkingCopies.end()) {
                std::string chartXmlText;
                if (z.contains(edit.sourceChartPart)) chartXmlText = z.get(edit.sourceChartPart);
                else {
                    const auto* raw = findPreservedPart(preservedParts, edit.sourceChartPart);
                    if (!raw) return false;
                    chartXmlText = raw->data;
                }
                chartIt = chartWorkingCopies.emplace(edit.sourceChartPart, std::move(chartXmlText)).first;
            }
            if (!applyImportedChartXmlEdit(chartIt->second, sheet, editIndex)) return false;
        }

        if (appendHere) {
            std::set<std::string> usedRelationshipIds;
            for (const auto& relationship : drawingRelationships) usedRelationshipIds.insert(relationship.id);
            std::size_t objectId = maximumDrawingObjectId(drawingXmlText) + 1;
            std::string appendedAnchors;
            for (std::size_t index = sheet.loadedChartCount(); index < charts.size(); ++index) {
                const auto& chart = charts[index];
                const auto relationshipId = allocateRelationshipId(usedRelationshipIds);
                usedRelationshipIds.insert(relationshipId);
                const auto chartId = nextChartId++;
                xlpp::PreservedRelationship relationship;
                relationship.sourcePart = drawingPart;
                relationship.id = relationshipId;
                relationship.type = chart.modern() ? "http://schemas.microsoft.com/office/2014/relationships/chartEx" : nsRelsDoc(sourceStrict) + "/chart";
                relationship.target = "../charts/chart" + std::to_string(chartId) + ".xml";
                drawingRelationships.push_back(std::move(relationship));
                relationshipsChanged = true;
                appendedAnchors += appendedChartAnchorXml(chart, relationshipId, objectId++, index, sourceStrict);
                z.add("xl/charts/chart" + std::to_string(chartId) + ".xml", serializeChart(chart, sourceStrict));
            }
            if (!appendedAnchors.empty()) {
                const auto closing = drawingXmlText.rfind("</xdr:wsDr>") != std::string::npos
                    ? std::string("</xdr:wsDr>") : std::string("</wsDr>");
                const auto closingPosition = drawingXmlText.rfind(closing);
                if (closingPosition == std::string::npos) return false;
                drawingXmlText.insert(closingPosition, appendedAnchors);
            }
        }

        z.add(drawingPart, std::move(drawingXmlText));
        if (relationshipsChanged)
            z.add(drawingRelsPart, xlpp::internal::RelationshipGraph::serializeRelationships(drawingRelationships, sourceStrict));
        for (auto& [part, data] : chartWorkingCopies) z.add(part, std::move(data));
    }
    return true;
}


} // namespace xlpp::internal::ooxml
