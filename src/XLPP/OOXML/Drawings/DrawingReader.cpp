#include "OOXML/Drawings/DrawingReader.h"
#include "OOXML/Common/DrawingXmlSupport.h"
#include "OOXML/Common/PackageRelationships.h"
#include "Package/Xml/XmlUtilities.h"
#include "Package/Zip/ZipArchive.h"
#include <XLPP/Worksheet/Worksheet.h>

#include <algorithm>
#include <cstddef>
#include <string>

using xlpp::internal::ooxml::drawingTags;
using xlpp::internal::ooxml::drawingTagText;
using xlpp::internal::ooxml::drawingInteger;
using xlpp::internal::ooxml::partExtension;
using xlpp::internal::ooxml::parseDrawingMarker;
using xlpp::internal::ooxml::resolvePackagePart;
using xlpp::internal::ooxml::relationshipKind;
using xlpp::internal::ooxml::relationshipsForSource;

namespace xlpp::internal::ooxml {

void loadWorksheetImages(xlpp::Worksheet& ws, const std::string& sheetXml, const xlpp::internal::ZipArchive& z,
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

        std::unordered_map<std::string, std::string> imageRelationships;
        for (const auto& rel : xlpp::internal::tags(z.get(drawingRelsPart), "Relationship"))
            if (xlpp::internal::attribute(rel, "Type").find("/image") != std::string::npos)
                imageRelationships[xlpp::internal::attribute(rel, "Id")] = xlpp::internal::attribute(rel, "Target");

        const auto loadAnchors = [&](const char* prefixedName, const char* localName, xlpp::DrawingAnchorType type) {
            for (const auto& anchorNode : drawingTags(drawingXmlText, prefixedName, localName)) {
                const auto pictureNodes = drawingTags(anchorNode, "xdr:pic", "pic");
                if (pictureNodes.empty()) continue;
                const auto& pictureNode = pictureNodes.front();
                const auto blips = drawingTags(pictureNode, "a:blip", "blip");
                if (blips.empty()) continue;
                auto imageRelId = xlpp::internal::attribute(blips.front(), "r:embed");
                if (imageRelId.empty()) imageRelId = xlpp::internal::attribute(blips.front(), "r:link");
                const auto imageRelationship = imageRelationships.find(imageRelId);
                if (imageRelationship == imageRelationships.end()) continue;
                const auto mediaPart = resolvePackagePart(drawingPart, imageRelationship->second);
                if (!z.contains(mediaPart)) continue;
                const auto extension = partExtension(mediaPart);
                if (extension.empty()) continue;

                xlpp::DrawingAnchorInfo anchorInfo;
                anchorInfo.type = type;
                anchorInfo.editAs = xlpp::internal::attribute(anchorNode, "editAs");
                const auto fromNodes = drawingTags(anchorNode, "xdr:from", "from");
                if (!fromNodes.empty()) anchorInfo.from = parseDrawingMarker(fromNodes.front());
                const auto toNodes = drawingTags(anchorNode, "xdr:to", "to");
                if (!toNodes.empty()) anchorInfo.to = parseDrawingMarker(toNodes.front());
                const auto posNodes = drawingTags(anchorNode, "xdr:pos", "pos");
                if (!posNodes.empty()) {
                    anchorInfo.xEmu = drawingInteger(posNodes.front(), "xdr:x", "x");
                    anchorInfo.yEmu = drawingInteger(posNodes.front(), "xdr:y", "y");
                    const auto x = xlpp::internal::attribute(posNodes.front(), "x");
                    const auto y = xlpp::internal::attribute(posNodes.front(), "y");
                    if (!x.empty()) anchorInfo.xEmu = std::stoll(x);
                    if (!y.empty()) anchorInfo.yEmu = std::stoll(y);
                }
                const auto extNodes = drawingTags(anchorNode, "xdr:ext", "ext");
                if (!extNodes.empty()) {
                    const auto cx = xlpp::internal::attribute(extNodes.front(), "cx");
                    const auto cy = xlpp::internal::attribute(extNodes.front(), "cy");
                    if (!cx.empty()) anchorInfo.widthEmu = std::stoll(cx);
                    if (!cy.empty()) anchorInfo.heightEmu = std::stoll(cy);
                }
                // twoCellAnchor stores image extents in a:xfrm rather than an
                // anchor-level xdr:ext. Capture those values for inspection.
                if ((anchorInfo.widthEmu <= 0 || anchorInfo.heightEmu <= 0) && type == xlpp::DrawingAnchorType::TwoCell) {
                    const auto transformExt = drawingTags(pictureNode, "a:ext", "ext");
                    if (!transformExt.empty()) {
                        const auto cx = xlpp::internal::attribute(transformExt.front(), "cx");
                        const auto cy = xlpp::internal::attribute(transformExt.front(), "cy");
                        if (!cx.empty()) anchorInfo.widthEmu = std::stoll(cx);
                        if (!cy.empty()) anchorInfo.heightEmu = std::stoll(cy);
                    }
                }

                const auto nonVisual = drawingTags(pictureNode, "xdr:cNvPr", "cNvPr");
                std::string objectId;
                std::string objectName = "Image";
                if (!nonVisual.empty()) {
                    objectId = xlpp::internal::attribute(nonVisual.front(), "id");
                    const auto parsedName = xlpp::internal::attribute(nonVisual.front(), "name");
                    if (!parsedName.empty()) objectName = parsedName;
                }
                const auto anchorAddress = xlpp::CellReference{anchorInfo.from.row, anchorInfo.from.column}.address();
                const auto bytesText = z.get(mediaPart);
                xlpp::Image image(anchorAddress,
                    std::vector<unsigned char>(bytesText.begin(), bytesText.end()), extension);
                image.setName(objectName);
                image.setAnchorInfo(anchorInfo);
                image.setStableId(drawingPart + "#" + (objectId.empty() ? imageRelId : objectId));
                image.setSourceDrawingPart(drawingPart);
                image.setSourceMediaPart(mediaPart);
                image.setSourceRelationshipId(imageRelId);
                image.setImported(true);
                if (anchorInfo.widthEmu > 0) image.setWidthPixels(static_cast<double>(anchorInfo.widthEmu) / 9525.0);
                if (anchorInfo.heightEmu > 0) image.setHeightPixels(static_cast<double>(anchorInfo.heightEmu) / 9525.0);
                ws.addLoadedImage(std::move(image));
            }
        };

        loadAnchors("xdr:oneCellAnchor", "oneCellAnchor", xlpp::DrawingAnchorType::OneCell);
        loadAnchors("xdr:twoCellAnchor", "twoCellAnchor", xlpp::DrawingAnchorType::TwoCell);
        loadAnchors("xdr:absoluteAnchor", "absoluteAnchor", xlpp::DrawingAnchorType::Absolute);
    }
}




} // namespace xlpp::internal::ooxml
