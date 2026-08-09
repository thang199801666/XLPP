#include "OOXML/Drawings/DrawingPackageSupport.h"
#include "OOXML/Common/DrawingXmlSupport.h"
#include "OOXML/Common/Namespaces.h"
#include "Package/Xml/XmlUtilities.h"
#include "Package/Opc/RelationshipGraph.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <set>
#include <sstream>

using xlpp::internal::xmlEscape;
using xlpp::internal::ooxml::drawingTags;
using xlpp::internal::ooxml::nsRelsDoc;

namespace xlpp::internal::ooxml::drawing_support {

const xlpp::PreservedPart* findPreservedPart(const std::vector<xlpp::PreservedPart>& parts,
                                             const std::string& name) {
    const auto it = std::find_if(parts.begin(), parts.end(), [&](const auto& part) { return part.name == name; });
    return it == parts.end() ? nullptr : &*it;
}

std::size_t maximumDrawingObjectId(const std::string& drawingXmlText) {
    std::size_t maximum = 0;
    const auto inspect = [&](const std::vector<std::string>& nodes) {
        for (const auto& node : nodes) {
            const auto value = xlpp::internal::attribute(node, "id");
            if (value.empty() || !std::all_of(value.begin(), value.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
            maximum = std::max(maximum, static_cast<std::size_t>(std::stoull(value)));
        }
    };
    inspect(xlpp::internal::tags(drawingXmlText, "xdr:cNvPr"));
    inspect(xlpp::internal::tags(drawingXmlText, "cNvPr"));
    return maximum;
}

std::string appendedImageAnchorXml(const xlpp::Image& image,
                                   const std::string& relationshipId,
                                   std::size_t objectId,
                                   bool strict) {
    const auto ref = xlpp::CellReference::parse(image.anchor());
    const auto cx = static_cast<long long>(image.widthPixels() * 9525.0);
    const auto cy = static_cast<long long>(image.heightPixels() * 9525.0);
    const auto drawingNs = strict ? "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                                  : "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const auto drawingMainNs = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                      : "http://schemas.openxmlformats.org/drawingml/2006/main";
    std::ostringstream xml;
    xml << "<xdr:oneCellAnchor xmlns:xdr=\"" << drawingNs << "\" xmlns:a=\"" << drawingMainNs
        << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\"><xdr:from><xdr:col>" << (ref.column - 1)
        << "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" << (ref.row - 1)
        << "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"" << cx << "\" cy=\"" << cy
        << "\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" << objectId << "\" name=\"" << xmlEscape(image.name())
        << "\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"" << relationshipId
        << "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/>"
        << "</a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>";
    return xml.str();
}

std::string appendedChartAnchorXml(const xlpp::Chart& chart,
                                   const std::string& relationshipId,
                                   std::size_t objectId,
                                   std::size_t placementIndex,
                                   bool strict) {
    const auto widthEmu = static_cast<long long>(chart.width()) * 9525LL;
    const auto heightEmu = static_cast<long long>(chart.height()) * 9525LL;
    const auto drawingNs = strict ? "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                                  : "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const auto drawingMainNs = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                      : "http://schemas.openxmlformats.org/drawingml/2006/main";
    const auto chartNs = strict ? "http://purl.oclc.org/ooxml/drawingml/chart"
                                : "http://schemas.openxmlformats.org/drawingml/2006/chart";
    std::ostringstream xml;
    xml << "<xdr:oneCellAnchor xmlns:xdr=\"" << drawingNs << "\" xmlns:a=\"" << drawingMainNs
        << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" << (placementIndex * 20)
        << "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"" << widthEmu << "\" cy=\"" << heightEmu
        << "\"/><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" << objectId
        << "\" name=\"Chart " << objectId << "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
        << "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" << widthEmu << "\" cy=\"" << heightEmu << "\"/></xdr:xfrm><a:graphic>";
    if (chart.modern())
        xml << "<a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chart xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" r:id=\"" << relationshipId << "\"/></a:graphicData>";
    else
        xml << "<a:graphicData uri=\"" << chartNs << "\"><c:chart xmlns:c=\"" << chartNs << "\" r:id=\"" << relationshipId << "\"/></a:graphicData>";
    xml << "</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>";
    return xml.str();
}


bool replaceSimpleDrawingText(std::string& xml, const char* prefixed, const char* local, long long value) {
    const std::array<std::string, 2> names{prefixed, local};
    for (const auto& name : names) {
        if (name.empty()) continue;
        const auto open = "<" + name + ">";
        const auto close = "</" + name + ">";
        const auto begin = xml.find(open);
        if (begin == std::string::npos) continue;
        const auto textBegin = begin + open.size();
        const auto finish = xml.find(close, textBegin);
        if (finish == std::string::npos) continue;
        xml.replace(textBegin, finish - textBegin, std::to_string(value));
        return true;
    }
    return false;
}

bool replaceAttributeInNode(std::string& node, const std::string& attributeName, long long value) {
    const auto key = attributeName + "=\"";
    const auto begin = node.find(key);
    if (begin == std::string::npos) return false;
    const auto valueBegin = begin + key.size();
    const auto valueEnd = node.find('"', valueBegin);
    if (valueEnd == std::string::npos) return false;
    node.replace(valueBegin, valueEnd - valueBegin, std::to_string(value));
    return true;
}

bool patchFirstDrawingNodeAttributes(std::string& anchor,
                                     const char* prefixed,
                                     const char* local,
                                     long long first,
                                     long long second,
                                     const char* firstAttribute,
                                     const char* secondAttribute) {
    auto nodes = drawingTags(anchor, prefixed, local);
    if (nodes.empty()) return false;
    auto node = nodes.front();
    if (!replaceAttributeInNode(node, firstAttribute, first) ||
        !replaceAttributeInNode(node, secondAttribute, second)) return false;
    const auto position = anchor.find(nodes.front());
    if (position == std::string::npos) return false;
    anchor.replace(position, nodes.front().size(), node);
    return true;
}

bool patchDrawingMarker(std::string& anchor, const char* prefixed, const char* local,
                        const xlpp::DrawingMarker& marker, bool updateOffsets = false) {
    auto markers = drawingTags(anchor, prefixed, local);
    if (markers.empty()) return false;
    auto markerXml = markers.front();
    if (!replaceSimpleDrawingText(markerXml, "xdr:col", "col", static_cast<long long>(marker.column) - 1)) return false;
    if (!replaceSimpleDrawingText(markerXml, "xdr:row", "row", static_cast<long long>(marker.row) - 1)) return false;
    if (updateOffsets) {
        if (!replaceSimpleDrawingText(markerXml, "xdr:colOff", "colOff", marker.columnOffsetEmu)) return false;
        if (!replaceSimpleDrawingText(markerXml, "xdr:rowOff", "rowOff", marker.rowOffsetEmu)) return false;
    }
    // Normal moves deliberately retain producer-native sub-cell offsets. A
    // two-cell resize, however, must update the terminal offsets because those
    // offsets participate directly in the visible width/height.
    const auto position = anchor.find(markers.front());
    if (position == std::string::npos) return false;
    anchor.replace(position, markers.front().size(), markerXml);
    return true;
}

std::string drawingObjectIdFromStableId(const std::string& stableId) {
    const auto hash = stableId.rfind('#');
    return hash == std::string::npos ? std::string{} : stableId.substr(hash + 1);
}

bool anchorMatchesStableId(const std::string& anchor, const std::string& stableId) {
    const auto objectId = drawingObjectIdFromStableId(stableId);
    if (objectId.empty()) return false;
    const auto pictures = drawingTags(anchor, "xdr:pic", "pic");
    for (const auto& picture : pictures) {
        const auto properties = drawingTags(picture, "xdr:cNvPr", "cNvPr");
        for (const auto& property : properties)
            if (xlpp::internal::attribute(property, "id") == objectId) return true;
    }
    return false;
}

bool anchorReferencesRelationship(const std::string& anchor, const std::string& relationshipId) {
    const auto pictures = drawingTags(anchor, "xdr:pic", "pic");
    for (const auto& picture : pictures) {
        const auto blips = drawingTags(picture, "a:blip", "blip");
        for (const auto& blip : blips) {
            if (xlpp::internal::attribute(blip, "r:embed") == relationshipId ||
                xlpp::internal::attribute(blip, "r:link") == relationshipId) return true;
        }
    }
    return false;
}

std::string* findImageAnchorBlock(std::vector<std::string>& anchors,
                                  const std::string& stableId,
                                  const std::string& relationshipId) {
    for (auto& anchor : anchors)
        if (anchorMatchesStableId(anchor, stableId)) return &anchor;
    for (auto& anchor : anchors)
        if (anchorReferencesRelationship(anchor, relationshipId)) return &anchor;
    return nullptr;
}

bool drawingReferencesRelationship(const std::string& drawingXmlText, const std::string& relationshipId) {
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        for (const auto& anchor : drawingTags(drawingXmlText, prefixed, local))
            if (anchorReferencesRelationship(anchor, relationshipId)) return true;
    }
    return false;
}

bool patchImportedImageAnchor(std::string& drawingXmlText,
                              const std::string& stableId,
                              const std::string& relationshipId,
                              const xlpp::DrawingAnchorInfo& anchorInfo,
                              bool moved,
                              bool resized,
                              bool remove) {
    std::vector<std::string> anchors;
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        auto family = drawingTags(drawingXmlText, prefixed, local);
        anchors.insert(anchors.end(), std::make_move_iterator(family.begin()), std::make_move_iterator(family.end()));
    }
    auto* found = findImageAnchorBlock(anchors, stableId, relationshipId);
    if (!found) return false;
    const auto original = *found;
    auto patched = original;
    if (remove) {
        const auto position = drawingXmlText.find(original);
        if (position == std::string::npos) return false;
        drawingXmlText.erase(position, original.size());
        return true;
    }

    if (moved) {
        if (anchorInfo.type == xlpp::DrawingAnchorType::Absolute) {
            if (!patchFirstDrawingNodeAttributes(patched, "xdr:pos", "pos",
                                                 anchorInfo.xEmu, anchorInfo.yEmu, "x", "y")) return false;
        } else {
            if (!patchDrawingMarker(patched, "xdr:from", "from", anchorInfo.from)) return false;
            if (anchorInfo.type == xlpp::DrawingAnchorType::TwoCell &&
                !patchDrawingMarker(patched, "xdr:to", "to", anchorInfo.to)) return false;
        }
    }
    if (resized) {
        if (anchorInfo.type == xlpp::DrawingAnchorType::TwoCell) {
            // Two-cell anchors derive their visible geometry primarily from
            // from/to. Patch the terminal marker as well as a:xfrm/a:ext so
            // Excel/LibreOffice do not normalize the requested size back to
            // the old cell span on the next save.
            if (!patchDrawingMarker(patched, "xdr:to", "to", anchorInfo.to, true)) return false;
            auto pictures = drawingTags(patched, "xdr:pic", "pic");
            if (pictures.empty()) return false;
            auto picture = pictures.front();
            if (!patchFirstDrawingNodeAttributes(picture, "a:ext", "ext",
                                                 anchorInfo.widthEmu, anchorInfo.heightEmu, "cx", "cy")) return false;
            const auto picturePosition = patched.find(pictures.front());
            if (picturePosition == std::string::npos) return false;
            patched.replace(picturePosition, pictures.front().size(), picture);
        } else {
            if (!patchFirstDrawingNodeAttributes(patched, "xdr:ext", "ext",
                                                 anchorInfo.widthEmu, anchorInfo.heightEmu, "cx", "cy")) return false;
        }
    }
    const auto position = drawingXmlText.find(original);
    if (position == std::string::npos) return false;
    drawingXmlText.replace(position, original.size(), patched);
    return true;
}


bool chartAnchorMatchesStableId(const std::string& anchor, const std::string& stableId) {
    const auto objectId = drawingObjectIdFromStableId(stableId);
    if (objectId.empty()) return false;
    const auto frames = drawingTags(anchor, "xdr:graphicFrame", "graphicFrame");
    for (const auto& frame : frames) {
        const auto properties = drawingTags(frame, "xdr:cNvPr", "cNvPr");
        for (const auto& property : properties)
            if (xlpp::internal::attribute(property, "id") == objectId) return true;
    }
    return false;
}

bool chartAnchorReferencesRelationship(const std::string& anchor, const std::string& relationshipId) {
    const auto frames = drawingTags(anchor, "xdr:graphicFrame", "graphicFrame");
    for (const auto& frame : frames) {
        const auto charts = drawingTags(frame, "c:chart", "chart");
        for (const auto& chart : charts)
            if (xlpp::internal::attribute(chart, "r:id") == relationshipId) return true;
    }
    return false;
}

bool drawingReferencesChartRelationship(const std::string& drawingXmlText, const std::string& relationshipId) {
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        for (const auto& anchor : drawingTags(drawingXmlText, prefixed, local))
            if (chartAnchorReferencesRelationship(anchor, relationshipId)) return true;
    }
    return false;
}

bool removeImportedChartAnchor(std::string& drawingXmlText,
                               const std::string& stableId,
                               const std::string& relationshipId) {
    std::vector<std::string> anchors;
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        auto family = drawingTags(drawingXmlText, prefixed, local);
        anchors.insert(anchors.end(), std::make_move_iterator(family.begin()), std::make_move_iterator(family.end()));
    }
    for (const auto& candidate : anchors) {
        if (!chartAnchorMatchesStableId(candidate, stableId) &&
            !chartAnchorReferencesRelationship(candidate, relationshipId)) continue;
        const auto position = drawingXmlText.find(candidate);
        if (position == std::string::npos) return false;
        drawingXmlText.erase(position, candidate.size());
        return true;
    }
    return false;
}

bool patchImportedChartAnchor(std::string& drawingXmlText,
                              const std::string& stableId,
                              const std::string& relationshipId,
                              const xlpp::DrawingAnchorInfo& anchorInfo,
                              bool moved,
                              bool resized) {
    std::vector<std::string> anchors;
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        auto family = drawingTags(drawingXmlText, prefixed, local);
        anchors.insert(anchors.end(), std::make_move_iterator(family.begin()), std::make_move_iterator(family.end()));
    }
    std::string* found = nullptr;
    for (auto& candidate : anchors) {
        if (chartAnchorMatchesStableId(candidate, stableId)) { found = &candidate; break; }
    }
    if (!found) {
        for (auto& candidate : anchors) {
            if (chartAnchorReferencesRelationship(candidate, relationshipId)) { found = &candidate; break; }
        }
    }
    if (!found) return false;

    const auto original = *found;
    auto patched = original;
    if (moved) {
        if (anchorInfo.type == xlpp::DrawingAnchorType::Absolute) {
            if (!patchFirstDrawingNodeAttributes(patched, "xdr:pos", "pos",
                                                 anchorInfo.xEmu, anchorInfo.yEmu, "x", "y")) return false;
        } else {
            if (!patchDrawingMarker(patched, "xdr:from", "from", anchorInfo.from)) return false;
            if (anchorInfo.type == xlpp::DrawingAnchorType::TwoCell &&
                !patchDrawingMarker(patched, "xdr:to", "to", anchorInfo.to)) return false;
        }
    }
    if (resized) {
        if (anchorInfo.type == xlpp::DrawingAnchorType::TwoCell) {
            if (!patchDrawingMarker(patched, "xdr:to", "to", anchorInfo.to, true)) return false;
            auto frames = drawingTags(patched, "xdr:graphicFrame", "graphicFrame");
            if (!frames.empty()) {
                auto frame = frames.front();
                const auto extNodes = drawingTags(frame, "a:ext", "ext");
                if (!extNodes.empty()) {
                    if (!patchFirstDrawingNodeAttributes(frame, "a:ext", "ext",
                                                         anchorInfo.widthEmu, anchorInfo.heightEmu, "cx", "cy")) return false;
                    const auto framePosition = patched.find(frames.front());
                    if (framePosition == std::string::npos) return false;
                    patched.replace(framePosition, frames.front().size(), frame);
                }
            }
        } else {
            if (!patchFirstDrawingNodeAttributes(patched, "xdr:ext", "ext",
                                                 anchorInfo.widthEmu, anchorInfo.heightEmu, "cx", "cy")) return false;
        }
    }
    const auto position = drawingXmlText.find(original);
    if (position == std::string::npos) return false;
    drawingXmlText.replace(position, original.size(), patched);
    return true;
}


} // namespace xlpp::internal::ooxml::drawing_support
