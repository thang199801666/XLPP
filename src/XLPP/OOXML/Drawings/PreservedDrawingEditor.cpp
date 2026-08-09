#include "OOXML/Drawings/PreservedDrawingEditor.h"
#include "OOXML/Drawings/DrawingPackageSupport.h"
#include "OOXML/Drawings/WorkbookDrawingAccess.h"
#include "OOXML/Common/DrawingXmlSupport.h"
#include "OOXML/Common/Namespaces.h"
#include "OOXML/Common/PackageRelationships.h"
#include "Package/Xml/XmlUtilities.h"
#include "Package/Zip/ZipArchive.h"
#include "Package/Opc/RelationshipGraph.h"

#include <algorithm>
#include <set>
#include <unordered_set>

using xlpp::internal::ooxml::allocateRelationshipId;
using xlpp::internal::ooxml::relationshipKind;
using xlpp::internal::ooxml::relationshipsForSource;
using xlpp::internal::ooxml::resolvePackagePart;
using xlpp::internal::ooxml::nsRelsDoc;
using xlpp::internal::ooxml::partExtension;
using namespace xlpp::internal::ooxml::drawing_support;

namespace xlpp::internal::ooxml {

bool relationshipTargetsPart(const xlpp::PreservedRelationship& relationship, const std::string& part) {
    return relationship.targetMode != "External"
        && resolvePackagePart(relationship.sourcePart, relationship.target) == part;
}

bool hasOtherRelationshipToPart(const std::vector<xlpp::PreservedRelationship>& relationships,
                                const std::string& sourcePart,
                                const std::string& relationshipId,
                                const std::string& part) {
    return std::any_of(relationships.begin(), relationships.end(), [&](const auto& relationship) {
        if (relationship.sourcePart == sourcePart && relationship.id == relationshipId) return false;
        return relationshipTargetsPart(relationship, part);
    });
}

bool applyImageChangesToPreservedDrawing(xlpp::internal::ZipArchive& z,
                                         const xlpp::Worksheet& sheet,
                                         const std::string& sourceSheetPart,
                                         const std::string& sourceSheetXml,
                                         const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                         const std::vector<xlpp::PreservedPart>& preservedParts,
                                         std::size_t& nextMediaId,
                                         std::set<std::string>& suppressedPreservedParts) {
    const auto& edits = xlpp::internal::WorkbookDrawingAccess::imageEdits(sheet);
    if (sheet.appendedImageCount() == 0 && edits.empty()) return true;
    if (sourceSheetPart.empty()) return false;

    if (sourceSheetXml.empty()) return false;
    const auto sheetRelationships = relationshipsForSource(allRelationships, sourceSheetPart);

    // A normal worksheet owns one DrawingML part, but preserving producer data
    // means we must not assume that.  Resolve every explicit <drawing r:id>
    // owner node and patch only the drawing part that owns each imported image.
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
        if (!drawingPart.empty() && seenDrawingParts.insert(drawingPart).second)
            ownedDrawingParts.push_back(drawingPart);
    }
    if (ownedDrawingParts.empty()) return false;

    for (const auto& edit : edits)
        if (!seenDrawingParts.count(edit.sourceDrawingPart)) return false;

    // New images are appended to the first preserved drawing by default.  If
    // the sheet already has imported images, prefer the drawing that owns the
    // first one so append behavior stays stable for multi-drawing producers.
    std::string appendDrawingPart = ownedDrawingParts.front();
    for (std::size_t index = 0; index < std::min(sheet.loadedImageCount(), sheet.images().size()); ++index) {
        const auto& candidate = sheet.images()[index];
        if (seenDrawingParts.count(candidate.sourceDrawingPart())) {
            appendDrawingPart = candidate.sourceDrawingPart();
            break;
        }
    }

    for (const auto& drawingPart : ownedDrawingParts) {
        const bool appendHere = sheet.appendedImageCount() > 0 && drawingPart == appendDrawingPart;
        const bool hasEditsHere = std::any_of(edits.begin(), edits.end(), [&](const auto& edit) {
            return edit.sourceDrawingPart == drawingPart;
        });
        if (!appendHere && !hasEditsHere) continue; // keep untouched drawing bytes exactly as loaded

        const auto* rawDrawingPart = findPreservedPart(preservedParts, drawingPart);
        if (!rawDrawingPart) return false;
        auto drawingXmlText = rawDrawingPart->data;
        const bool sourceStrict = drawingXmlText.find("http://purl.oclc.org/ooxml/drawingml/") != std::string::npos;
        auto drawingRelationships = relationshipsForSource(allRelationships, drawingPart);

        for (const auto& edit : edits) {
            if (edit.sourceDrawingPart != drawingPart) continue;
            if (!patchImportedImageAnchor(drawingXmlText, edit.stableId, edit.sourceRelationshipId, edit.anchor,
                                          edit.moved, edit.resized, edit.removed)) return false;
            auto relationship = std::find_if(drawingRelationships.begin(), drawingRelationships.end(), [&](const auto& candidate) {
                return candidate.id == edit.sourceRelationshipId;
            });
            if (relationship == drawingRelationships.end()) return false;

            if (edit.removed) {
                if (!drawingReferencesRelationship(drawingXmlText, edit.sourceRelationshipId)) {
                    const bool shared = hasOtherRelationshipToPart(allRelationships, drawingPart,
                                                                   edit.sourceRelationshipId, edit.sourceMediaPart);
                    drawingRelationships.erase(relationship);
                    if (!shared) suppressedPreservedParts.insert(edit.sourceMediaPart);
                }
                continue;
            }
            if (edit.replaced) {
                const bool shared = hasOtherRelationshipToPart(allRelationships, drawingPart,
                                                               edit.sourceRelationshipId, edit.sourceMediaPart);
                const auto oldExtension = partExtension(edit.sourceMediaPart);
                if (!shared && oldExtension == edit.replacementExtension) {
                    const std::string bytes(reinterpret_cast<const char*>(edit.replacementBytes.data()), edit.replacementBytes.size());
                    z.add(edit.sourceMediaPart, bytes, false);
                } else {
                    const auto mediaId = nextMediaId++;
                    const auto newMediaPart = "xl/media/image" + std::to_string(mediaId) + "." + edit.replacementExtension;
                    relationship->target = "../media/image" + std::to_string(mediaId) + "." + edit.replacementExtension;
                    const std::string bytes(reinterpret_cast<const char*>(edit.replacementBytes.data()), edit.replacementBytes.size());
                    z.add(newMediaPart, bytes, false);
                    if (!shared) suppressedPreservedParts.insert(edit.sourceMediaPart);
                }
            }
        }

        if (appendHere) {
            std::set<std::string> usedRelationshipIds;
            for (const auto& relationship : drawingRelationships) usedRelationshipIds.insert(relationship.id);
            std::size_t objectId = maximumDrawingObjectId(drawingXmlText) + 1;
            std::string appendedAnchors;
            const auto& images = sheet.images();
            for (std::size_t index = sheet.loadedImageCount(); index < images.size(); ++index) {
                const auto& image = images[index];
                const auto relationshipId = allocateRelationshipId(usedRelationshipIds);
                usedRelationshipIds.insert(relationshipId);
                const auto mediaId = nextMediaId++;
                xlpp::PreservedRelationship relationship;
                relationship.sourcePart = drawingPart;
                relationship.id = relationshipId;
                relationship.type = nsRelsDoc(sourceStrict) + "/image";
                relationship.target = "../media/image" + std::to_string(mediaId) + "." + image.extension();
                drawingRelationships.push_back(std::move(relationship));
                appendedAnchors += appendedImageAnchorXml(image, relationshipId, objectId++, sourceStrict);
                const std::string bytes(reinterpret_cast<const char*>(image.bytes().data()), image.bytes().size());
                z.add("xl/media/image" + std::to_string(mediaId) + "." + image.extension(), bytes, false);
            }
            if (!appendedAnchors.empty()) {
                const auto closing = drawingXmlText.rfind("</xdr:wsDr>") != std::string::npos ? std::string("</xdr:wsDr>") : std::string("</wsDr>");
                const auto closingPosition = drawingXmlText.rfind(closing);
                if (closingPosition == std::string::npos) return false;
                drawingXmlText.insert(closingPosition, appendedAnchors);
            }
        }

        z.add(drawingPart, std::move(drawingXmlText));
        const auto slash = drawingPart.find_last_of('/');
        const auto drawingFile = drawingPart.substr(slash + 1);
        const auto drawingRelsPart = drawingPart.substr(0, slash + 1) + "_rels/" + drawingFile + ".rels";
        z.add(drawingRelsPart, xlpp::internal::RelationshipGraph::serializeRelationships(drawingRelationships, sourceStrict));
    }
    return true;
}


} // namespace xlpp::internal::ooxml
