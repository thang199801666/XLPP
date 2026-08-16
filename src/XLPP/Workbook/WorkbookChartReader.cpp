#include "WorkbookChartReader.h"
#include "WorkbookChartSerializer.h"
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Cell/Cell.h>
#include <XLPP/Cell/DateTime.h>
#include "../XML/XmlUtilities.h"
#include "../XML/NumericParsing.h"
#include "../Packaging/ZipArchive.h"
#include <algorithm>
#include <array>
#include <cctype>
#include <cmath>
#include <cstdlib>
#include <limits>
#include <memory>
#include <optional>
#include <sstream>
#include <stdexcept>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>
#include <cstdint>

namespace xlpp {
namespace internal {

std::string resolvePackagePart(const std::string& basePart, std::string relativeTarget) {
    if (relativeTarget.empty()) return {};
    const bool absoluteTarget = relativeTarget.front() == '/';
    if (absoluteTarget) relativeTarget.erase(relativeTarget.begin());
    std::vector<std::string> segments;
    const auto slash = basePart.find_last_of('/');
    std::string combined = absoluteTarget
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
    for (std::size_t i = 0; i < segments.size(); ++i) { if (i) result << '/'; result << segments[i]; }
    return result.str();
}
long long drawingInteger(const std::string& xml, const char* prefixed, const char* local, long long fallback = 0) {
    const auto value = drawingTagText(xml, prefixed, local);
    if (value.empty()) return fallback;
    try { return std::stoll(value); } catch (...) { return fallback; }
}

std::string partExtension(const std::string& part) {
    const auto slash = part.find_last_of('/');
    const auto dot = part.find_last_of('.');
    if (dot == std::string::npos || (slash != std::string::npos && dot < slash)) return {};
    auto extension = part.substr(dot + 1);
    std::transform(extension.begin(), extension.end(), extension.begin(), [](unsigned char c) {
        return static_cast<char>(std::tolower(c));
    });
    if (extension == "jpeg") extension = "jpg";
    return extension;
}

xlpp::DrawingMarker parseDrawingMarker(const std::string& markerXml) {
    xlpp::DrawingMarker marker;
    marker.column = static_cast<std::size_t>(std::max<long long>(0, drawingInteger(markerXml, "xdr:col", "col"))) + 1;
    marker.row = static_cast<std::size_t>(std::max<long long>(0, drawingInteger(markerXml, "xdr:row", "row"))) + 1;
    marker.columnOffsetEmu = drawingInteger(markerXml, "xdr:colOff", "colOff");
    marker.rowOffsetEmu = drawingInteger(markerXml, "xdr:rowOff", "rowOff");
    return marker;
}

void loadImages(xlpp::Worksheet& ws, const std::string& sheetXml, const xlpp::internal::ZipArchive& z,
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
                    long long parsedX = 0, parsedY = 0;
                    if (!x.empty() && xlpp::internal::tryParseIntegerExact(x, parsedX)) anchorInfo.xEmu = parsedX;
                    if (!y.empty() && xlpp::internal::tryParseIntegerExact(y, parsedY)) anchorInfo.yEmu = parsedY;
                }
                const auto extNodes = drawingTags(anchorNode, "xdr:ext", "ext");
                if (!extNodes.empty()) {
                    const auto cx = xlpp::internal::attribute(extNodes.front(), "cx");
                    const auto cy = xlpp::internal::attribute(extNodes.front(), "cy");
                    long long parsedCx = 0, parsedCy = 0;
                    if (!cx.empty() && xlpp::internal::tryParseIntegerExact(cx, parsedCx)) anchorInfo.widthEmu = parsedCx;
                    if (!cy.empty() && xlpp::internal::tryParseIntegerExact(cy, parsedCy)) anchorInfo.heightEmu = parsedCy;
                }
                // twoCellAnchor stores image extents in a:xfrm rather than an
                // anchor-level xdr:ext. Capture those values for inspection.
                if ((anchorInfo.widthEmu <= 0 || anchorInfo.heightEmu <= 0) && type == xlpp::DrawingAnchorType::TwoCell) {
                    const auto transformExt = drawingTags(pictureNode, "a:ext", "ext");
                    if (!transformExt.empty()) {
                        const auto cx = xlpp::internal::attribute(transformExt.front(), "cx");
                        const auto cy = xlpp::internal::attribute(transformExt.front(), "cy");
                        long long parsedCx = 0, parsedCy = 0;
                        if (!cx.empty() && xlpp::internal::tryParseIntegerExact(cx, parsedCx)) anchorInfo.widthEmu = parsedCx;
                        if (!cy.empty() && xlpp::internal::tryParseIntegerExact(cy, parsedCy)) anchorInfo.heightEmu = parsedCy;
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
                // Read rotation/flip transform from the drawing a:xfrm.
                const auto xfrmNodes = drawingTags(anchorNode, "a:xfrm", "xfrm");
                if (!xfrmNodes.empty()) {
                    const auto rotText = xlpp::internal::attribute(xfrmNodes.front(), "rot");
                    long long rot = 0;
                    if (!rotText.empty() && xlpp::internal::tryParseIntegerExact(rotText, rot))
                        image.setRotation(static_cast<double>(rot) / 60000.0);
                    image.setFlipHorizontal(xlpp::internal::attribute(xfrmNodes.front(), "flipH") == "1");
                    image.setFlipVertical(xlpp::internal::attribute(xfrmNodes.front(), "flipV") == "1");
                }
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


struct ParsedChartPlotNode {
    std::size_t position{0};
    xlpp::Chart::Type type{xlpp::Chart::Type::Bar};
    std::string xml;
};

std::vector<ParsedChartPlotNode> chartPlotNodesInOrder(const std::string& chartXmlText) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return {};
    const auto& plotArea = plotAreas.front();
    std::vector<ParsedChartPlotNode> result;
    const auto collect = [&](const char* prefixed, const char* local, xlpp::Chart::Type type) {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, prefixed, local)) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            result.push_back({position, type, node});
            cursor = position + node.size();
        }
    };
    collect("c:barChart", "barChart", xlpp::Chart::Type::Bar);
    collect("c:lineChart", "lineChart", xlpp::Chart::Type::Line);
    collect("c:pieChart", "pieChart", xlpp::Chart::Type::Pie);
    collect("c:scatterChart", "scatterChart", xlpp::Chart::Type::Scatter);
    collect("c:doughnutChart", "doughnutChart", xlpp::Chart::Type::Doughnut);
    collect("c:radarChart", "radarChart", xlpp::Chart::Type::Radar);
    collect("c:areaChart", "areaChart", xlpp::Chart::Type::Area);
    collect("c:bubbleChart", "bubbleChart", xlpp::Chart::Type::Bubble);
    collect("c:stockChart", "stockChart", xlpp::Chart::Type::Stock);
    {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, "c:ofPieChart", "ofPieChart")) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            const auto kinds = drawingTags(node, "c:ofPieType", "ofPieType");
            const auto value = kinds.empty() ? std::string{"pie"} : xlpp::internal::attribute(kinds.front(), "val");
            result.push_back({position, value == "bar" ? xlpp::Chart::Type::BarOfPie : xlpp::Chart::Type::PieOfPie, node});
            cursor = position + node.size();
        }
    }
    collect("c:bar3DChart", "bar3DChart", xlpp::Chart::Type::Bar3D);
    collect("c:line3DChart", "line3DChart", xlpp::Chart::Type::Line3D);
    collect("c:area3DChart", "area3DChart", xlpp::Chart::Type::Area3D);
    collect("c:pie3DChart", "pie3DChart", xlpp::Chart::Type::Pie3D);
    collect("c:surfaceChart", "surfaceChart", xlpp::Chart::Type::Surface);
    collect("c:surface3DChart", "surface3DChart", xlpp::Chart::Type::Surface3D);
    std::sort(result.begin(), result.end(), [](const auto& a, const auto& b) { return a.position < b.position; });
    return result;
}

xlpp::Chart::Grouping parseChartGroupingNode(const std::string& chartNode) {
    const auto groupingNodes = drawingTags(chartNode, "c:grouping", "grouping");
    if (groupingNodes.empty()) return xlpp::Chart::Grouping::Standard;
    const auto value = xlpp::internal::attribute(groupingNodes.front(), "val");
    if (value == "stacked") return xlpp::Chart::Grouping::Stacked;
    if (value == "percentStacked") return xlpp::Chart::Grouping::PercentStacked;
    if (value == "clustered") return xlpp::Chart::Grouping::Clustered;
    return xlpp::Chart::Grouping::Standard;
}

std::vector<std::uint64_t> chartAxisIds(const std::string& chartNode) {
    std::vector<std::uint64_t> result;
    for (const auto& node : drawingTags(chartNode, "c:axId", "axId")) {
        const auto value = xlpp::internal::attribute(node, "val");
        if (value.empty()) continue;
        try { result.push_back(std::stoull(value)); } catch (...) {}
    }
    return result;
}

bool chartBoolValue(const std::string& container, const char* prefixed, const char* local, bool fallback) {
    const auto nodes = drawingTags(container, prefixed, local);
    if (nodes.empty()) return fallback;
    const auto value = xlpp::internal::attribute(nodes.front(), "val");
    return value == "1" || value == "true" || value == "True";
}

std::vector<xlpp::ChartColorTransform> parseChartColorTransformsInOrder(const std::string& container) {
    using K = xlpp::ChartColorTransform::Kind;
    std::vector<std::pair<std::size_t, xlpp::ChartColorTransform>> ordered;
    for (const auto& transform : std::array<std::tuple<const char*, const char*, K>, 9>{
             std::tuple{"a:alpha", "alpha", K::Alpha},
             std::tuple{"a:alphaMod", "alphaMod", K::AlphaMod},
             std::tuple{"a:alphaOff", "alphaOff", K::AlphaOff},
             std::tuple{"a:tint", "tint", K::Tint},
             std::tuple{"a:shade", "shade", K::Shade},
             std::tuple{"a:lumMod", "lumMod", K::LumMod},
             std::tuple{"a:lumOff", "lumOff", K::LumOff},
             std::tuple{"a:satMod", "satMod", K::SatMod},
             std::tuple{"a:satOff", "satOff", K::SatOff}}) {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(container, std::get<0>(transform), std::get<1>(transform))) {
            const auto position = container.find(node, cursor);
            if (position == std::string::npos) continue;
            cursor = position + node.size();
            try {
                ordered.push_back({position, {std::get<2>(transform), std::stoi(xlpp::internal::attribute(node, "val"))}});
            } catch (...) {}
        }
    }
    std::sort(ordered.begin(), ordered.end(), [](const auto& lhs, const auto& rhs) { return lhs.first < rhs.first; });
    std::vector<xlpp::ChartColorTransform> result;
    result.reserve(ordered.size());
    for (auto& item : ordered) result.push_back(std::move(item.second));
    return result;
}

xlpp::ChartColor parseChartColor(const std::string& container) {
    using Kind = xlpp::ChartColor::Kind;
    for (const auto& entry : std::array<std::tuple<const char*, const char*, Kind>, 4>{
             std::tuple{"a:srgbClr", "srgbClr", Kind::SRgb},
             std::tuple{"a:schemeClr", "schemeClr", Kind::Scheme},
             std::tuple{"a:sysClr", "sysClr", Kind::System},
             std::tuple{"a:prstClr", "prstClr", Kind::Preset}}) {
        const auto nodes = drawingTags(container, std::get<0>(entry), std::get<1>(entry));
        if (nodes.empty()) continue;
        const auto value = xlpp::internal::attribute(nodes.front(), "val");
        if (value.empty()) continue;
        xlpp::ChartColor color;
        color.kind = std::get<2>(entry);
        color.value = value;
        const auto& colorXml = nodes.front();
        color.transforms = parseChartColorTransformsInOrder(colorXml);
        return color;
    }
    return {};
}


xlpp::ChartLineFormat parseChartLineFormat(const std::string& container);
xlpp::ChartFillFormat parseChartFillFormat(const std::string& container);

xlpp::ChartThemeEffectStyle parseThemeEffectStyle(const std::string& container) {
    xlpp::ChartThemeEffectStyle style;
    const auto effectLists = drawingTags(container, "a:effectLst", "effectLst");
    const auto scope = effectLists.empty() ? container : effectLists.front();
    const auto parseShadowGeometry = [](const std::string& node, double& blur, double& distance, double& direction) {
        try { blur = std::stod(xlpp::internal::attribute(node, "blurRad")) / 12700.0; } catch (...) {}
        try { distance = std::stod(xlpp::internal::attribute(node, "dist")) / 12700.0; } catch (...) {}
        try { direction = std::stod(xlpp::internal::attribute(node, "dir")) / 60000.0; } catch (...) {}
    };
    const auto shadows = drawingTags(scope, "a:outerShdw", "outerShdw");
    if (!shadows.empty()) {
        style.present = style.outerShadow = true;
        style.outerShadowColor = parseChartColor(shadows.front());
        parseShadowGeometry(shadows.front(), style.outerShadowBlurPoints,
                            style.outerShadowDistancePoints, style.outerShadowDirectionDegrees);
    }
    const auto innerShadows = drawingTags(scope, "a:innerShdw", "innerShdw");
    if (!innerShadows.empty()) {
        style.present = style.innerShadow = true;
        style.innerShadowColor = parseChartColor(innerShadows.front());
        parseShadowGeometry(innerShadows.front(), style.innerShadowBlurPoints,
                            style.innerShadowDistancePoints, style.innerShadowDirectionDegrees);
    }
    const auto glows = drawingTags(scope, "a:glow", "glow");
    if (!glows.empty()) {
        style.present = style.glow = true;
        style.glowColor = parseChartColor(glows.front());
        try { style.glowRadiusPoints = std::stod(xlpp::internal::attribute(glows.front(), "rad")) / 12700.0; } catch (...) {}
    }
    const auto softEdges = drawingTags(scope, "a:softEdge", "softEdge");
    if (!softEdges.empty()) {
        style.present = style.softEdge = true;
        try { style.softEdgeRadiusPoints = std::stod(xlpp::internal::attribute(softEdges.front(), "rad")) / 12700.0; } catch (...) {}
    }
    const auto reflections = drawingTags(scope, "a:reflection", "reflection");
    if (!reflections.empty()) {
        style.present = style.reflection = true;
        try { style.reflectionBlurPoints = std::stod(xlpp::internal::attribute(reflections.front(), "blurRad")) / 12700.0; } catch (...) {}
        try { style.reflectionDistancePoints = std::stod(xlpp::internal::attribute(reflections.front(), "dist")) / 12700.0; } catch (...) {}
        try { style.reflectionDirectionDegrees = std::stod(xlpp::internal::attribute(reflections.front(), "dir")) / 60000.0; } catch (...) {}
    }
    const auto blurs = drawingTags(scope, "a:blur", "blur");
    if (!blurs.empty()) {
        style.present = style.blur = true;
        try { style.blurRadiusPoints = std::stod(xlpp::internal::attribute(blurs.front(), "rad")) / 12700.0; } catch (...) {}
        const auto grow = xlpp::internal::attribute(blurs.front(), "grow");
        style.blurGrow = grow == "1" || grow == "true" || grow == "True";
    }
    return style;
}

xlpp::ChartSeriesCache parseChartSeriesCache(const std::string& container) {
    xlpp::ChartSeriesCache result;
    auto caches = drawingTags(container, "c:numCache", "numCache");
    if (!caches.empty()) result.numeric = true;
    else { caches = drawingTags(container, "c:strCache", "strCache"); result.numeric = false; }
    if (caches.empty()) return result;
    result.present = true;
    const auto& cache = caches.front();
    result.formatCode = drawingTagText(cache, "c:formatCode", "formatCode");
    const auto counts = drawingTags(cache, "c:ptCount", "ptCount");
    if (!counts.empty()) { try { result.pointCount = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(counts.front(), "val"))); } catch (...) {} }
    for (const auto& point : drawingTags(cache, "c:pt", "pt")) {
        const auto idx = xlpp::internal::attribute(point, "idx");
        if (idx.empty()) continue;
        try { result.points.push_back({static_cast<std::size_t>(std::stoull(idx)), drawingTagText(point, "c:v", "v")}); } catch (...) {}
    }
    return result;
}

xlpp::ChartThemePalette parseChartThemePalette(const xlpp::internal::ZipArchive& z) {
    xlpp::ChartThemePalette palette;
    if (!z.contains("xl/theme/theme1.xml")) return palette;
    const auto xml = z.get("xl/theme/theme1.xml");
    for (const auto& name : std::array<const char*, 12>{"dk1","lt1","dk2","lt2","accent1","accent2","accent3","accent4","accent5","accent6","hlink","folHlink"}) {
        const auto prefixed = std::string("a:") + name;
        const auto nodes = drawingTags(xml, prefixed.c_str(), name);
        if (nodes.empty()) continue;
        std::string value;
        const auto srgb = drawingTags(nodes.front(), "a:srgbClr", "srgbClr");
        if (!srgb.empty()) value = xlpp::internal::attribute(srgb.front(), "val");
        if (value.empty()) {
            const auto sys = drawingTags(nodes.front(), "a:sysClr", "sysClr");
            if (!sys.empty()) { value = xlpp::internal::attribute(sys.front(), "lastClr"); if (value.empty()) value = xlpp::internal::attribute(sys.front(), "val"); }
        }
        if (!value.empty()) palette.colors.push_back({name, value});
    }
    const auto fontSchemes = drawingTags(xml, "a:fontScheme", "fontScheme");
    if (!fontSchemes.empty()) {
        palette.fontScheme.present = true;
        palette.fontScheme.name = xlpp::internal::attribute(fontSchemes.front(), "name");
        const auto major = drawingTags(fontSchemes.front(), "a:majorFont", "majorFont");
        if (!major.empty()) {
            const auto latin = drawingTags(major.front(), "a:latin", "latin");
            if (!latin.empty()) palette.fontScheme.majorLatinTypeface = xlpp::internal::attribute(latin.front(), "typeface");
        }
        const auto minor = drawingTags(fontSchemes.front(), "a:minorFont", "minorFont");
        if (!minor.empty()) {
            const auto latin = drawingTags(minor.front(), "a:latin", "latin");
            if (!latin.empty()) palette.fontScheme.minorLatinTypeface = xlpp::internal::attribute(latin.front(), "typeface");
        }
    }
    const auto fmtSchemes = drawingTags(xml, "a:fmtScheme", "fmtScheme");
    if (!fmtSchemes.empty()) {
        auto& effects = palette.effectScheme; effects.present = true; effects.name = xlpp::internal::attribute(fmtSchemes.front(), "name");
        const auto materializeFillsInOrder = [&](const std::string& list) {
            std::vector<std::pair<std::size_t, xlpp::ChartFillFormat>> ordered;
            for (const auto* tag : {"solidFill", "gradFill", "pattFill", "noFill"}) {
                std::size_t cursor = 0;
                for (const auto& node : drawingTags(list, (std::string("a:") + tag).c_str(), tag)) {
                    const auto position = list.find(node, cursor);
                    if (position == std::string::npos) continue;
                    ordered.push_back({position, parseChartFillFormat(node)});
                    cursor = position + node.size();
                }
            }
            std::sort(ordered.begin(), ordered.end(), [](const auto& lhs, const auto& rhs) { return lhs.first < rhs.first; });
            std::vector<xlpp::ChartFillFormat> result; result.reserve(ordered.size());
            for (auto& entry : ordered) result.push_back(std::move(entry.second));
            return result;
        };
        const auto fills = drawingTags(fmtSchemes.front(), "a:fillStyleLst", "fillStyleLst");
        if (!fills.empty()) {
            effects.fillStyles = materializeFillsInOrder(fills.front());
            effects.fillStyleCount = effects.fillStyles.size();
        }
        const auto lines = drawingTags(fmtSchemes.front(), "a:lnStyleLst", "lnStyleLst");
        if (!lines.empty()) {
            for (const auto& node : drawingTags(lines.front(), "a:ln", "ln"))
                effects.lineStyles.push_back(parseChartLineFormat(node));
            effects.lineStyleCount = effects.lineStyles.size();
        }
        const auto effectStyles = drawingTags(fmtSchemes.front(), "a:effectStyleLst", "effectStyleLst");
        if (!effectStyles.empty()) {
            for (const auto& node : drawingTags(effectStyles.front(), "a:effectStyle", "effectStyle"))
                effects.effectStyles.push_back(parseThemeEffectStyle(node));
            effects.effectStyleCount = effects.effectStyles.size();
        }
        const auto bgFills = drawingTags(fmtSchemes.front(), "a:bgFillStyleLst", "bgFillStyleLst");
        if (!bgFills.empty()) {
            effects.backgroundFillStyles = materializeFillsInOrder(bgFills.front());
            effects.backgroundFillStyleCount = effects.backgroundFillStyles.size();
        }
    }
    palette.present = !palette.colors.empty() || palette.fontScheme.present || palette.effectScheme.present;
    return palette;
}

std::vector<xlpp::ChartColorTransform> parseChartStyleColorTransforms(const std::string& container) {
    return parseChartColorTransformsInOrder(container);
}

xlpp::ChartStyleReference parseChartStyleReference(const std::string& container,
                                                   const char* prefixed, const char* local) {
    xlpp::ChartStyleReference result;
    const auto refs = drawingTags(container, prefixed, local);
    if (refs.empty()) return result;
    const auto& node = refs.front();
    result.present = true;
    try { result.index = std::stoi(xlpp::internal::attribute(node, "idx")); } catch (...) { result.index = -1; }
    result.modifiers = xlpp::internal::attribute(node, "mods");
    result.color = parseChartColor(node);
    const auto styleColors = drawingTags(node, "cs:styleClr", "styleClr");
    if (!styleColors.empty()) {
        result.styleColor = true;
        result.styleColorValue = xlpp::internal::attribute(styleColors.front(), "val");
        result.styleColorTransforms = parseChartStyleColorTransforms(styleColors.front());
    }
    return result;
}

xlpp::ChartStyleRule parseChartStyleRule(const std::string& target, const std::string& node) {
    xlpp::ChartStyleRule rule;
    rule.target = target;
    rule.modifiers = xlpp::internal::attribute(node, "mods");
    rule.lineReference = parseChartStyleReference(node, "cs:lnRef", "lnRef");
    rule.fillReference = parseChartStyleReference(node, "cs:fillRef", "fillRef");
    rule.effectReference = parseChartStyleReference(node, "cs:effectRef", "effectRef");
    const auto widthScales = drawingTags(node, "cs:lineWidthScale", "lineWidthScale");
    if (!widthScales.empty()) {
        auto text = drawingTagText(node, "cs:lineWidthScale", "lineWidthScale");
        if (text.empty()) text = xlpp::internal::attribute(widthScales.front(), "val");
        try { rule.lineWidthScale = std::stod(text); rule.hasLineWidthScale = true; } catch (...) {}
    }
    const auto fontRefs = drawingTags(node, "cs:fontRef", "fontRef");
    if (!fontRefs.empty()) {
        const auto& font = fontRefs.front();
        rule.fontIndex = xlpp::internal::attribute(font, "idx");
        rule.fontModifiers = xlpp::internal::attribute(font, "mods");
        rule.fontColor = parseChartColor(font);
        const auto styleColors = drawingTags(font, "cs:styleClr", "styleClr");
        if (!styleColors.empty()) {
            rule.fontStyleColor = true;
            rule.fontStyleColorValue = xlpp::internal::attribute(styleColors.front(), "val");
            rule.fontStyleColorTransforms = parseChartStyleColorTransforms(styleColors.front());
        }
    }
    const auto shapeProperties = drawingTags(node, "cs:spPr", "spPr");
    if (!shapeProperties.empty()) {
        rule.shapeFill = parseChartFillFormat(shapeProperties.front());
        rule.shapeLine = parseChartLineFormat(shapeProperties.front());
    }
    return rule;
}

xlpp::ChartStyleResources parseChartStyleResources(const xlpp::internal::ZipArchive& z, const std::string& chartPart) {
    xlpp::ChartStyleResources resources;
    const auto slash = chartPart.find_last_of('/');
    if (slash == std::string::npos) return resources;
    const auto file = chartPart.substr(slash + 1);
    const auto relsPart = chartPart.substr(0, slash + 1) + "_rels/" + file + ".rels";
    if (!z.contains(relsPart)) return resources;
    for (const auto& rel : xlpp::internal::tags(z.get(relsPart), "Relationship")) {
        const auto type = xlpp::internal::attribute(rel, "Type");
        const auto target = xlpp::internal::attribute(rel, "Target");
        if (target.empty()) continue;
        if (type.find("/chartStyle") != std::string::npos) { resources.chartStylePresent = true; resources.chartStylePart = resolvePackagePart(chartPart, target); }
        else if (type.find("/chartColorStyle") != std::string::npos) { resources.colorStylePresent = true; resources.colorStylePart = resolvePackagePart(chartPart, target); }
    }
    const auto parseId = [](const std::string& node) -> int {
        const auto id = xlpp::internal::attribute(node, "id");
        if (id.empty()) return -1;
        try { return std::stoi(id); } catch (...) { return -1; }
    };
    if (resources.chartStylePresent && z.contains(resources.chartStylePart)) {
        const auto xml = z.get(resources.chartStylePart);
        const auto roots = drawingTags(xml, "cs:chartStyle", "chartStyle");
        if (!roots.empty()) {
            resources.chartStyleId = parseId(roots.front());
            const auto& root = roots.front();
            for (const auto* target : std::array<const char*, 29>{
                     "axisTitle", "categoryAxis", "chartArea", "dataLabel", "dataLabelCallout",
                     "dataPoint", "dataPoint3D", "dataPointLine", "dataPointMarker", "dataPointWireframe",
                     "dataTable", "downBar", "dropLine", "errorBar", "floor", "gridlineMajor",
                     "gridlineMinor", "hiLoLine", "leaderLine", "legend", "plotArea", "plotArea3D",
                     "seriesAxis", "seriesLine", "title", "trendline", "trendlineLabel", "upBar", "valueAxis"}) {
                const auto prefixed = std::string("cs:") + target;
                const auto nodes = drawingTags(root, prefixed.c_str(), target);
                if (!nodes.empty()) resources.chartStyleRules.push_back(parseChartStyleRule(target, nodes.front()));
            }
            // wall appears after valueAxis in CT_ChartStyle; keep it separate so
            // the target list above remains readable and order-independent.
            const auto walls = drawingTags(root, "cs:wall", "wall");
            if (!walls.empty()) resources.chartStyleRules.push_back(parseChartStyleRule("wall", walls.front()));
            const auto markerLayouts = drawingTags(root, "cs:dataPointMarkerLayout", "dataPointMarkerLayout");
            if (!markerLayouts.empty()) {
                resources.markerLayout.present = true;
                resources.markerLayout.symbol = xlpp::internal::attribute(markerLayouts.front(), "symbol");
                try { resources.markerLayout.size = std::stoi(xlpp::internal::attribute(markerLayouts.front(), "size")); } catch (...) {}
            }
        }
    }
    if (resources.colorStylePresent && z.contains(resources.colorStylePart)) {
        const auto xml = z.get(resources.colorStylePart);
        const auto roots = drawingTags(xml, "cs:colorStyle", "colorStyle");
        if (!roots.empty()) {
            resources.colorStyleId = parseId(roots.front());
            resources.colorStyleMethod = xlpp::internal::attribute(roots.front(), "meth");
        }
        std::vector<std::pair<std::size_t, xlpp::ChartColor>> orderedColors;
        for (const auto& entry : std::array<std::pair<const char*, const char*>, 4>{
                 std::pair{"a:schemeClr", "schemeClr"},
                 std::pair{"a:srgbClr", "srgbClr"},
                 std::pair{"a:sysClr", "sysClr"},
                 std::pair{"a:prstClr", "prstClr"}}) {
            std::size_t cursor = 0;
            for (const auto& node : drawingTags(xml, entry.first, entry.second)) {
                const auto position = xml.find(node, cursor);
                if (position == std::string::npos) continue;
                cursor = position + node.size();
                auto color = parseChartColor(node);
                if (color.present()) orderedColors.push_back({position, std::move(color)});
            }
        }
        std::sort(orderedColors.begin(), orderedColors.end(), [](const auto& lhs, const auto& rhs) { return lhs.first < rhs.first; });
        for (auto& entry : orderedColors) resources.colorStyleColors.push_back(std::move(entry.second));
    }
    return resources;
}

xlpp::ChartLineFormat parseChartLineFormat(const std::string& container) {
    xlpp::ChartLineFormat format;
    const auto lines = drawingTags(container, "a:ln", "ln");
    if (lines.empty()) return format;
    const auto& line = lines.front();
    format.present = true;
    const auto width = xlpp::internal::attribute(line, "w");
    if (!width.empty()) {
        try { format.widthPoints = std::stod(width) / 12700.0; } catch (...) {}
    }
    format.cap = xlpp::internal::attribute(line, "cap");
    format.compound = xlpp::internal::attribute(line, "cmpd");
    format.noFill = !drawingTags(line, "a:noFill", "noFill").empty();
    const auto fills = drawingTags(line, "a:solidFill", "solidFill");
    if (!fills.empty()) format.color = parseChartColor(fills.front());
    const auto dashes = drawingTags(line, "a:prstDash", "prstDash");
    if (!dashes.empty()) format.dash = xlpp::internal::attribute(dashes.front(), "val");
    const auto custom = drawingTags(line, "a:custDash", "custDash");
    if (!custom.empty()) {
        for (const auto& ds : drawingTags(custom.front(), "a:ds", "ds")) {
            double d = 0.0, sp = 0.0;
            if (xlpp::internal::tryParseDoubleExact(xlpp::internal::attribute(ds, "d"), d)
                && xlpp::internal::tryParseDoubleExact(xlpp::internal::attribute(ds, "sp"), sp)) {
                format.customDash.push_back({d / 1000.0, sp / 1000.0});
            }
        }
    }
    if (!drawingTags(line, "a:round", "round").empty()) format.join = "round";
    else if (!drawingTags(line, "a:bevel", "bevel").empty()) format.join = "bevel";
    else if (!drawingTags(line, "a:miter", "miter").empty()) format.join = "miter";
    return format;
}

xlpp::ChartFillFormat parseChartFillFormat(const std::string& container) {
    xlpp::ChartFillFormat format;
    auto scope = container;
    for (const auto& line : drawingTags(scope, "a:ln", "ln")) {
        const auto position = scope.find(line);
        if (position != std::string::npos) scope.erase(position, line.size());
    }
    if (!drawingTags(scope, "a:noFill", "noFill").empty()) {
        format.present = true;
        format.noFill = true;
        format.kind = xlpp::ChartFillFormat::Kind::NoFill;
        return format;
    }
    const auto fills = drawingTags(scope, "a:solidFill", "solidFill");
    if (!fills.empty()) {
        format.present = true;
        format.kind = xlpp::ChartFillFormat::Kind::Solid;
        format.color = parseChartColor(fills.front());
        return format;
    }
    const auto gradients = drawingTags(scope, "a:gradFill", "gradFill");
    if (!gradients.empty()) {
        format.present = true;
        format.kind = xlpp::ChartFillFormat::Kind::Gradient;
        for (const auto& gs : drawingTags(gradients.front(), "a:gs", "gs")) {
            xlpp::ChartGradientStop stop;
            try { stop.position = std::stoi(xlpp::internal::attribute(gs, "pos")); } catch (...) {}
            stop.color = parseChartColor(gs);
            if (stop.color.present()) format.gradientStops.push_back(std::move(stop));
        }
        const auto linear = drawingTags(gradients.front(), "a:lin", "lin");
        if (!linear.empty()) {
            try { format.gradientAngleDegrees = std::stod(xlpp::internal::attribute(linear.front(), "ang")) / 60000.0; } catch (...) {}
        }
        return format;
    }
    const auto patterns = drawingTags(scope, "a:pattFill", "pattFill");
    if (!patterns.empty()) {
        format.present = true;
        format.kind = xlpp::ChartFillFormat::Kind::Pattern;
        format.pattern = xlpp::internal::attribute(patterns.front(), "prst");
        const auto foreground = drawingTags(patterns.front(), "a:fgClr", "fgClr");
        if (!foreground.empty()) format.foregroundColor = parseChartColor(foreground.front());
        const auto background = drawingTags(patterns.front(), "a:bgClr", "bgClr");
        if (!background.empty()) format.backgroundColor = parseChartColor(background.front());
    }
    return format;
}

xlpp::ChartMarkerFormat parseChartMarkerFormat(const std::string& seriesXml) {
    xlpp::ChartMarkerFormat format;
    const auto markers = drawingTags(seriesXml, "c:marker", "marker");
    if (markers.empty()) return format;
    const auto& marker = markers.front();
    format.present = true;
    const auto symbols = drawingTags(marker, "c:symbol", "symbol");
    if (!symbols.empty()) format.symbol = xlpp::internal::attribute(symbols.front(), "val");
    const auto sizes = drawingTags(marker, "c:size", "size");
    if (!sizes.empty()) { try { format.size = std::stoi(xlpp::internal::attribute(sizes.front(), "val")); } catch (...) {} }
    const auto spPr = drawingTags(marker, "c:spPr", "spPr");
    if (!spPr.empty()) {
        format.line = parseChartLineFormat(spPr.front());
        format.fill = parseChartFillFormat(spPr.front());
    }
    return format;
}

std::string seriesDirectSpPr(const std::string& seriesXml) {
    const auto candidates = drawingTags(seriesXml, "c:spPr", "spPr");
    if (candidates.empty()) return {};
    std::vector<std::string> nested;
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 5>{
             std::pair{"c:marker", "marker"}, {"c:dPt", "dPt"}, {"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}, {"c:errBars", "errBars"}}) {
        const auto nodes = drawingTags(seriesXml, pair.first, pair.second);
        nested.insert(nested.end(), nodes.begin(), nodes.end());
    }
    for (const auto& candidate : candidates) {
        if (std::none_of(nested.begin(), nested.end(), [&](const auto& node) { return node.find(candidate) != std::string::npos; }))
            return candidate;
    }
    return {};
}

xlpp::ChartRichText parseChartRichText(const std::string& owner) {
    xlpp::ChartRichText result;
    const auto richNodes = drawingTags(owner, "c:rich", "rich");
    if (richNodes.empty()) return result;
    result.present = true;
    const auto& rich = richNodes.front();
    for (const auto& runXml : drawingTags(rich, "a:r", "r")) {
        xlpp::ChartTextRun run;
        run.text = drawingTagText(runXml, "a:t", "t");
        const auto properties = drawingTags(runXml, "a:rPr", "rPr");
        if (!properties.empty()) {
            const auto& rPr = properties.front();
            const auto bold = xlpp::internal::attribute(rPr, "b");
            const auto italic = xlpp::internal::attribute(rPr, "i");
            run.bold = bold == "1" || bold == "true";
            run.italic = italic == "1" || italic == "true";
            const auto size = xlpp::internal::attribute(rPr, "sz");
            if (!size.empty()) { try { run.fontSizePoints = std::stod(size) / 100.0; } catch (...) {} }
            const auto latin = drawingTags(rPr, "a:latin", "latin");
            if (!latin.empty()) run.typeface = xlpp::internal::attribute(latin.front(), "typeface");
            run.color = parseChartColor(rPr);
        }
        result.runs.push_back(std::move(run));
    }
    if (result.runs.empty()) {
        const auto text = drawingTagText(rich, "a:t", "t");
        if (!text.empty()) result.runs.push_back({text});
    }
    return result;
}

xlpp::ChartTextStyle parseChartTextStyle(const std::string& owner) {
    xlpp::ChartTextStyle style;
    const auto txPr = drawingTags(owner, "c:txPr", "txPr");
    if (txPr.empty()) return style;
    std::string properties;
    const auto defaults = drawingTags(txPr.front(), "a:defRPr", "defRPr");
    if (!defaults.empty()) properties = defaults.front();
    else {
        const auto ends = drawingTags(txPr.front(), "a:endParaRPr", "endParaRPr");
        if (!ends.empty()) properties = ends.front();
        else {
            const auto runs = drawingTags(txPr.front(), "a:rPr", "rPr");
            if (!runs.empty()) properties = runs.front();
        }
    }
    if (properties.empty()) { style.present = true; return style; }
    style.present = true;
    const auto bold = xlpp::internal::attribute(properties, "b");
    const auto italic = xlpp::internal::attribute(properties, "i");
    style.bold = bold == "1" || bold == "true";
    style.italic = italic == "1" || italic == "true";
    const auto size = xlpp::internal::attribute(properties, "sz");
    if (!size.empty()) { try { style.fontSizePoints = std::stod(size) / 100.0; } catch (...) {} }
    const auto latin = drawingTags(properties, "a:latin", "latin");
    if (!latin.empty()) style.typeface = xlpp::internal::attribute(latin.front(), "typeface");
    style.color = parseChartColor(properties);
    return style;
}

std::vector<xlpp::ChartDataPointFormat> parseChartDataPoints(const std::string& seriesXml) {
    std::vector<xlpp::ChartDataPointFormat> result;
    for (const auto& node : drawingTags(seriesXml, "c:dPt", "dPt")) {
        xlpp::ChartDataPointFormat point;
        const auto indices = drawingTags(node, "c:idx", "idx");
        if (!indices.empty()) { try { point.index = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(indices.front(), "val"))); } catch (...) {} }
        const auto spPr = drawingTags(node, "c:spPr", "spPr");
        if (!spPr.empty()) {
            point.line = parseChartLineFormat(spPr.front());
            point.fill = parseChartFillFormat(spPr.front());
        }
        point.marker = parseChartMarkerFormat(node);
        result.push_back(std::move(point));
    }
    return result;
}

xlpp::ChartDataLabelPoint parseChartDataLabelPoint(const std::string& xml) {
    xlpp::ChartDataLabelPoint point;
    const auto idx = drawingTags(xml, "c:idx", "idx");
    if (!idx.empty()) { try { point.index = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(idx.front(), "val"))); } catch (...) {} }
    point.deleted = chartBoolValue(xml, "c:delete", "delete");
    point.showLegendKey = chartBoolValue(xml, "c:showLegendKey", "showLegendKey");
    point.showValue = chartBoolValue(xml, "c:showVal", "showVal");
    point.showCategoryName = chartBoolValue(xml, "c:showCatName", "showCatName");
    point.showSeriesName = chartBoolValue(xml, "c:showSerName", "showSerName");
    point.showPercent = chartBoolValue(xml, "c:showPercent", "showPercent");
    point.showBubbleSize = chartBoolValue(xml, "c:showBubbleSize", "showBubbleSize");
    point.showLeaderLines = chartBoolValue(xml, "c:showLeaderLines", "showLeaderLines");
    const auto positions = drawingTags(xml, "c:dLblPos", "dLblPos");
    if (!positions.empty()) point.position = xlpp::internal::attribute(positions.front(), "val");
    point.separator = drawingTagText(xml, "c:separator", "separator");
    point.richText = parseChartRichText(xml);
    return point;
}

xlpp::Chart::DataLabels parseChartDataLabels(const std::string& plotXml, bool directPlotChild) {
    xlpp::Chart::DataLabels labels;
    auto nodes = drawingTags(plotXml, "c:dLbls", "dLbls");
    if (directPlotChild && !nodes.empty()) {
        const auto seriesNodes = drawingTags(plotXml, "c:ser", "ser");
        nodes.erase(std::remove_if(nodes.begin(), nodes.end(), [&](const std::string& node) {
            return std::any_of(seriesNodes.begin(), seriesNodes.end(), [&](const std::string& series) {
                return series.find(node) != std::string::npos;
            });
        }), nodes.end());
    }
    if (nodes.empty()) return labels;
    labels.present = true;
    const auto& xml = nodes.front();
    auto aggregateXml = xml;
    for (const auto& pointNode : drawingTags(xml, "c:dLbl", "dLbl")) {
        labels.points.push_back(parseChartDataLabelPoint(pointNode));
        const auto position = aggregateXml.find(pointNode);
        if (position != std::string::npos) aggregateXml.erase(position, pointNode.size());
    }
    labels.showLegendKey = chartBoolValue(aggregateXml, "c:showLegendKey", "showLegendKey");
    labels.showValue = chartBoolValue(aggregateXml, "c:showVal", "showVal");
    labels.showCategoryName = chartBoolValue(aggregateXml, "c:showCatName", "showCatName");
    labels.showSeriesName = chartBoolValue(aggregateXml, "c:showSerName", "showSerName");
    labels.showPercent = chartBoolValue(aggregateXml, "c:showPercent", "showPercent");
    labels.showBubbleSize = chartBoolValue(aggregateXml, "c:showBubbleSize", "showBubbleSize");
    labels.showLeaderLines = chartBoolValue(aggregateXml, "c:showLeaderLines", "showLeaderLines");
    const auto leaderLines = drawingTags(aggregateXml, "c:leaderLines", "leaderLines");
    if (!leaderLines.empty()) {
        labels.hasLeaderLines = true;
        const auto spPr = drawingTags(leaderLines.front(), "c:spPr", "spPr");
        if (!spPr.empty()) labels.leaderLineFormat = parseChartLineFormat(spPr.front());
    }
    const auto positions = drawingTags(aggregateXml, "c:dLblPos", "dLblPos");
    if (!positions.empty()) labels.position = xlpp::internal::attribute(positions.front(), "val");
    labels.separator = drawingTagText(aggregateXml, "c:separator", "separator");
    return labels;
}

xlpp::ChartSeries::TrendlineType parseTrendlineTypeValue(const std::string& value) {
    using T = xlpp::ChartSeries::TrendlineType;
    if (value == "exp") return T::Exponential;
    if (value == "log") return T::Logarithmic;
    if (value == "poly") return T::Polynomial;
    if (value == "power") return T::Power;
    if (value == "movingAvg") return T::MovingAverage;
    return T::Linear;
}

std::vector<xlpp::ChartSeries::Trendline> parseChartTrendlines(const std::string& seriesXml) {
    std::vector<xlpp::ChartSeries::Trendline> result;
    for (const auto& node : drawingTags(seriesXml, "c:trendline", "trendline")) {
        xlpp::ChartSeries::Trendline trendline;
        const auto types = drawingTags(node, "c:trendlineType", "trendlineType");
        if (!types.empty()) trendline.type = parseTrendlineTypeValue(xlpp::internal::attribute(types.front(), "val"));
        const auto order = drawingTags(node, "c:order", "order");
        if (!order.empty()) { try { trendline.order = std::stoi(xlpp::internal::attribute(order.front(), "val")); } catch (...) {} }
        const auto period = drawingTags(node, "c:period", "period");
        if (!period.empty()) { try { trendline.period = std::stoi(xlpp::internal::attribute(period.front(), "val")); } catch (...) {} }
        const auto forward = drawingTags(node, "c:forward", "forward");
        if (!forward.empty()) { try { trendline.forward = std::stod(xlpp::internal::attribute(forward.front(), "val")); } catch (...) {} }
        const auto backward = drawingTags(node, "c:backward", "backward");
        if (!backward.empty()) { try { trendline.backward = std::stod(xlpp::internal::attribute(backward.front(), "val")); } catch (...) {} }
        trendline.displayEquation = chartBoolValue(node, "c:dispEq", "dispEq");
        trendline.displayRSquared = chartBoolValue(node, "c:dispRSqr", "dispRSqr");
        const auto spPr = drawingTags(node, "c:spPr", "spPr");
        if (!spPr.empty()) trendline.lineFormat = parseChartLineFormat(spPr.front());
        result.push_back(std::move(trendline));
    }
    return result;
}

xlpp::ChartSeries::ErrorBarDirection parseErrorBarDirectionValue(const std::string& value) {
    return value == "x" ? xlpp::ChartSeries::ErrorBarDirection::X : xlpp::ChartSeries::ErrorBarDirection::Y;
}

xlpp::ChartSeries::ErrorBarType parseErrorBarTypeValue(const std::string& value) {
    using T = xlpp::ChartSeries::ErrorBarType;
    if (value == "plus") return T::Plus;
    if (value == "minus") return T::Minus;
    return T::Both;
}

xlpp::ChartSeries::ErrorValueType parseErrorValueTypeValue(const std::string& value) {
    using T = xlpp::ChartSeries::ErrorValueType;
    if (value == "percentage") return T::Percentage;
    if (value == "stdDev") return T::StandardDeviation;
    if (value == "stdErr") return T::StandardError;
    if (value == "cust") return T::Custom;
    return T::FixedValue;
}

std::vector<xlpp::ChartSeries::ErrorBars> parseChartErrorBars(const std::string& seriesXml) {
    std::vector<xlpp::ChartSeries::ErrorBars> result;
    for (const auto& node : drawingTags(seriesXml, "c:errBars", "errBars")) {
        xlpp::ChartSeries::ErrorBars bars;
        const auto dirs = drawingTags(node, "c:errDir", "errDir");
        if (!dirs.empty()) bars.direction = parseErrorBarDirectionValue(xlpp::internal::attribute(dirs.front(), "val"));
        const auto types = drawingTags(node, "c:errBarType", "errBarType");
        if (!types.empty()) bars.barType = parseErrorBarTypeValue(xlpp::internal::attribute(types.front(), "val"));
        const auto valueTypes = drawingTags(node, "c:errValType", "errValType");
        if (!valueTypes.empty()) bars.valueType = parseErrorValueTypeValue(xlpp::internal::attribute(valueTypes.front(), "val"));
        bars.noEndCap = chartBoolValue(node, "c:noEndCap", "noEndCap");
        const auto values = drawingTags(node, "c:val", "val");
        if (!values.empty()) { try { bars.value = std::stod(xlpp::internal::attribute(values.front(), "val")); } catch (...) {} }
        const auto plus = drawingTags(node, "c:plus", "plus");
        if (!plus.empty()) bars.plusReference = drawingTagText(plus.front(), "c:f", "f");
        const auto minus = drawingTags(node, "c:minus", "minus");
        if (!minus.empty()) bars.minusReference = drawingTagText(minus.front(), "c:f", "f");
        const auto spPr = drawingTags(node, "c:spPr", "spPr");
        if (!spPr.empty()) bars.lineFormat = parseChartLineFormat(spPr.front());
        result.push_back(std::move(bars));
    }
    return result;
}

xlpp::ChartDataTable parseChartDataTable(const std::string& plotArea) {
    xlpp::ChartDataTable table;
    const auto nodes = drawingTags(plotArea, "c:dTable", "dTable");
    if (nodes.empty()) return table;
    table.present = true;
    const auto& xml = nodes.front();
    table.showHorizontalBorder = chartBoolValue(xml, "c:showHorzBorder", "showHorzBorder");
    table.showVerticalBorder = chartBoolValue(xml, "c:showVertBorder", "showVertBorder");
    table.showOutline = chartBoolValue(xml, "c:showOutline", "showOutline");
    table.showLegendKeys = chartBoolValue(xml, "c:showKeys", "showKeys");
    const auto spPr = drawingTags(xml, "c:spPr", "spPr");
    if (!spPr.empty()) {
        table.line = parseChartLineFormat(spPr.front());
        table.fill = parseChartFillFormat(spPr.front());
    }
    table.textStyle = parseChartTextStyle(xml);
    return table;
}

xlpp::ChartLineFormat parsePlotLineObject(const std::string& plotXml, const char* prefixed, const char* local, bool& present) {
    present = false;
    const auto nodes = drawingTags(plotXml, prefixed, local);
    if (nodes.empty()) return {};
    present = true;
    const auto spPr = drawingTags(nodes.front(), "c:spPr", "spPr");
    return spPr.empty() ? xlpp::ChartLineFormat{} : parseChartLineFormat(spPr.front());
}

xlpp::ChartUpDownBars parseChartUpDownBars(const std::string& plotXml) {
    xlpp::ChartUpDownBars bars;
    const auto nodes = drawingTags(plotXml, "c:upDownBars", "upDownBars");
    if (nodes.empty()) return bars;
    bars.present = true;
    const auto& xml = nodes.front();
    const auto gaps = drawingTags(xml, "c:gapWidth", "gapWidth");
    if (!gaps.empty()) { try { bars.gapWidth = std::stoi(xlpp::internal::attribute(gaps.front(), "val")); } catch (...) {} }
    const auto parseBar = [&](const char* prefixed, const char* local, xlpp::ChartFillFormat& fill, xlpp::ChartLineFormat& line) {
        const auto barNodes = drawingTags(xml, prefixed, local);
        if (barNodes.empty()) return;
        const auto spPr = drawingTags(barNodes.front(), "c:spPr", "spPr");
        if (!spPr.empty()) { fill = parseChartFillFormat(spPr.front()); line = parseChartLineFormat(spPr.front()); }
    };
    parseBar("c:upBars", "upBars", bars.upFill, bars.upLine);
    parseBar("c:downBars", "downBars", bars.downFill, bars.downLine);
    return bars;
}

std::vector<xlpp::Chart::Plot> parseChartPlots(const std::string& chartXmlText) {
    std::vector<xlpp::Chart::Plot> plots;
    std::size_t firstSeries = 0;
    for (const auto& parsed : chartPlotNodesInOrder(chartXmlText)) {
        xlpp::Chart::Plot plot;
        plot.type = parsed.type;
        plot.grouping = parseChartGroupingNode(parsed.xml);
        plot.axisIds = chartAxisIds(parsed.xml);
        plot.firstSeries = firstSeries;
        plot.seriesCount = drawingTags(parsed.xml, "c:ser", "ser").size();
        plot.dataLabels = parseChartDataLabels(parsed.xml, true);
        plot.dropLinesFormat = parsePlotLineObject(parsed.xml, "c:dropLines", "dropLines", plot.hasDropLines);
        plot.highLowLinesFormat = parsePlotLineObject(parsed.xml, "c:hiLowLines", "hiLowLines", plot.hasHighLowLines);
        plot.upDownBars = parseChartUpDownBars(parsed.xml);
        const auto gapDepth = drawingTags(parsed.xml, "c:gapDepth", "gapDepth");
        if (!gapDepth.empty()) { try { plot.gapDepth=std::stoi(xlpp::internal::attribute(gapDepth.front(),"val")); plot.hasGapDepth=true; } catch (...) {} }
        const auto wireframe = drawingTags(parsed.xml, "c:wireframe", "wireframe");
        if (!wireframe.empty()) { plot.hasWireframe=true; const auto value=xlpp::internal::attribute(wireframe.front(),"val"); plot.wireframe=value=="1"||value=="true"||value=="True"; }
        const auto shape = drawingTags(parsed.xml, "c:shape", "shape");
        if (!shape.empty()) plot.shape=xlpp::internal::attribute(shape.front(),"val");
        const auto firstSlice = drawingTags(parsed.xml, "c:firstSliceAng", "firstSliceAng");
        if (!firstSlice.empty()) { try { plot.firstSliceAngle=std::stoi(xlpp::internal::attribute(firstSlice.front(),"val")); plot.hasFirstSliceAngle=true; } catch (...) {} }
        const auto holeSize = drawingTags(parsed.xml, "c:holeSize", "holeSize");
        if (!holeSize.empty()) { try { plot.holeSize=std::stoi(xlpp::internal::attribute(holeSize.front(),"val")); plot.hasHoleSize=true; } catch (...) {} }
        const auto radarStyle = drawingTags(parsed.xml, "c:radarStyle", "radarStyle");
        if (!radarStyle.empty()) plot.radarStyle=xlpp::internal::attribute(radarStyle.front(),"val");
        if (parsed.type == xlpp::Chart::Type::PieOfPie || parsed.type == xlpp::Chart::Type::BarOfPie) {
            auto& options=plot.projectedPie; options.present=true; options.ofPieType=parsed.type==xlpp::Chart::Type::BarOfPie?"bar":"pie";
            const auto gapWidth=drawingTags(parsed.xml,"c:gapWidth","gapWidth"); if(!gapWidth.empty()) { try { options.gapWidth=std::stoi(xlpp::internal::attribute(gapWidth.front(),"val")); } catch (...) {} }
            const auto splitType=drawingTags(parsed.xml,"c:splitType","splitType"); if(!splitType.empty()) options.splitType=xlpp::internal::attribute(splitType.front(),"val");
            const auto splitPos=drawingTags(parsed.xml,"c:splitPos","splitPos"); if(!splitPos.empty()) { try { options.splitPosition=std::stod(xlpp::internal::attribute(splitPos.front(),"val")); options.hasSplitPosition=true; } catch (...) {} }
            const auto secondSize=drawingTags(parsed.xml,"c:secondPieSize","secondPieSize"); if(!secondSize.empty()) { try { options.secondPlotSize=std::stoi(xlpp::internal::attribute(secondSize.front(),"val")); } catch (...) {} }
            const auto custom=drawingTags(parsed.xml,"c:custSplit","custSplit"); if(!custom.empty()) for(const auto& pt:drawingTags(custom.front(),"c:secondPiePt","secondPiePt")) { try { options.customSplitPoints.push_back(std::stoi(xlpp::internal::attribute(pt,"val"))); } catch (...) {} }
            options.seriesLinesFormat=parsePlotLineObject(parsed.xml,"c:serLines","serLines",options.hasSeriesLines);
        }
        firstSeries += plot.seriesCount;
        plots.push_back(std::move(plot));
    }
    return plots;
}

xlpp::ChartManualLayout parseChartManualLayout(const std::string& owner) {
    xlpp::ChartManualLayout layout;
    const auto layouts = drawingTags(owner, "c:layout", "layout");
    if (layouts.empty()) return layout;
    const auto manuals = drawingTags(layouts.front(), "c:manualLayout", "manualLayout");
    if (manuals.empty()) return layout;
    layout.present = true;
    const auto& xml = manuals.front();
    const auto readString=[&](const char* p,const char* l){ const auto n=drawingTags(xml,p,l); return n.empty()?std::string{}:xlpp::internal::attribute(n.front(),"val"); };
    const auto readDouble=[&](const char* p,const char* l,bool& has,double& value){ const auto n=drawingTags(xml,p,l); if(n.empty()) return; const auto v=xlpp::internal::attribute(n.front(),"val"); if(v.empty()) return; try{ value=std::stod(v); has=true; }catch(...){} };
    layout.target=readString("c:layoutTarget","layoutTarget");
    layout.xMode=readString("c:xMode","xMode"); layout.yMode=readString("c:yMode","yMode");
    layout.widthMode=readString("c:wMode","wMode"); layout.heightMode=readString("c:hMode","hMode");
    readDouble("c:x","x",layout.hasX,layout.x); readDouble("c:y","y",layout.hasY,layout.y);
    readDouble("c:w","w",layout.hasWidth,layout.width); readDouble("c:h","h",layout.hasHeight,layout.height);
    return layout;
}

xlpp::ChartView3D parseChartView3D(const std::string& chartXmlText) {
    xlpp::ChartView3D view;
    const auto nodes = drawingTags(chartXmlText, "c:view3D", "view3D");
    if (nodes.empty()) return view;
    view.present = true;
    const auto& xml = nodes.front();
    const auto readInt = [&](const char* prefixed, const char* local, bool& has, int& out) {
        const auto values = drawingTags(xml, prefixed, local);
        if (values.empty()) return;
        const auto value = xlpp::internal::attribute(values.front(), "val");
        if (value.empty()) return;
        try { out = std::stoi(value); has = true; } catch (...) {}
    };
    readInt("c:rotX", "rotX", view.hasRotationX, view.rotationX);
    readInt("c:rotY", "rotY", view.hasRotationY, view.rotationY);
    readInt("c:hPercent", "hPercent", view.hasHeightPercent, view.heightPercent);
    readInt("c:depthPercent", "depthPercent", view.hasDepthPercent, view.depthPercent);
    readInt("c:perspective", "perspective", view.hasPerspective, view.perspective);
    const auto rAng = drawingTags(xml, "c:rAngAx", "rAngAx");
    if (!rAng.empty()) {
        view.hasRightAngleAxes = true;
        const auto value = xlpp::internal::attribute(rAng.front(), "val");
        view.rightAngleAxes = value.empty() || value == "1" || value == "true" || value == "True";
    }
    return view;
}

xlpp::ChartWallFormat parseChartWallFormat(const std::string& chartXmlText, const char* prefixed, const char* local) {
    xlpp::ChartWallFormat wall;
    const auto nodes = drawingTags(chartXmlText, prefixed, local);
    if (nodes.empty()) return wall;
    wall.present = true;
    const auto& xml = nodes.front();
    const auto thickness = drawingTags(xml, "c:thickness", "thickness");
    if (!thickness.empty()) {
        const auto value = xlpp::internal::attribute(thickness.front(), "val");
        int parsed = 0;
        if (!value.empty() && xlpp::internal::tryParseIntegerExact(value, parsed)) {
            wall.thickness = parsed;
            wall.hasThickness = true;
        }
    }
    const auto spPr = drawingTags(xml, "c:spPr", "spPr");
    if (!spPr.empty()) {
        wall.line = parseChartLineFormat(spPr.front());
        wall.fill = parseChartFillFormat(spPr.front());
    }
    return wall;
}

std::string axisDirectSpPr(const std::string& axisXml) {
    const auto candidates=drawingTags(axisXml,"c:spPr","spPr");
    if(candidates.empty()) return {};
    std::vector<std::string> nested;
    for(const auto& pair: std::array<std::pair<const char*,const char*>,5>{{{"c:title","title"},{"c:majorGridlines","majorGridlines"},{"c:minorGridlines","minorGridlines"},{"c:txPr","txPr"},{"c:extLst","extLst"}}}) {
        const auto nodes=drawingTags(axisXml,pair.first,pair.second); nested.insert(nested.end(),nodes.begin(),nodes.end());
    }
    for(const auto& candidate:candidates)
        if(std::none_of(nested.begin(),nested.end(),[&](const auto& node){ return node.find(candidate)!=std::string::npos; })) return candidate;
    return {};
}

std::vector<xlpp::Chart::Axis> parseChartAxes(const std::string& chartXmlText,
                                               const std::vector<xlpp::Chart::Plot>& plots) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return {};
    const auto& plotArea = plotAreas.front();
    struct AxisNode { std::size_t position; xlpp::Chart::AxisKind kind; std::string xml; };
    std::vector<AxisNode> nodes;
    const auto collect = [&](const char* prefixed, const char* local, xlpp::Chart::AxisKind kind) {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, prefixed, local)) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            nodes.push_back({position, kind, node});
            cursor = position + node.size();
        }
    };
    collect("c:catAx", "catAx", xlpp::Chart::AxisKind::Category);
    collect("c:valAx", "valAx", xlpp::Chart::AxisKind::Value);
    collect("c:dateAx", "dateAx", xlpp::Chart::AxisKind::Date);
    collect("c:serAx", "serAx", xlpp::Chart::AxisKind::Series);
    std::sort(nodes.begin(), nodes.end(), [](const auto& a, const auto& b) { return a.position < b.position; });

    std::set<std::uint64_t> primaryAxisIds;
    if (!plots.empty()) primaryAxisIds.insert(plots.front().axisIds.begin(), plots.front().axisIds.end());
    std::set<std::uint64_t> laterPlotAxisIds;
    for (std::size_t i = 1; i < plots.size(); ++i)
        laterPlotAxisIds.insert(plots[i].axisIds.begin(), plots[i].axisIds.end());

    std::vector<xlpp::Chart::Axis> axes;
    for (const auto& parsed : nodes) {
        xlpp::Chart::Axis axis;
        axis.kind = parsed.kind;
        const auto ids = drawingTags(parsed.xml, "c:axId", "axId");
        if (!ids.empty()) {
            const auto value = xlpp::internal::attribute(ids.front(), "val");
            try { if (!value.empty()) axis.id = std::stoull(value); } catch (...) {}
        }
        const auto crosses = drawingTags(parsed.xml, "c:crossAx", "crossAx");
        if (!crosses.empty()) {
            const auto value = xlpp::internal::attribute(crosses.front(), "val");
            try { if (!value.empty()) axis.crossAxisId = std::stoull(value); } catch (...) {}
        }
        const auto positions = drawingTags(parsed.xml, "c:axPos", "axPos");
        if (!positions.empty()) axis.position = xlpp::internal::attribute(positions.front(), "val");
        const auto scalings = drawingTags(parsed.xml, "c:scaling", "scaling");
        if (!scalings.empty()) {
            const auto& scaling = scalings.front();
            const auto minimum = drawingTags(scaling, "c:min", "min");
            if (!minimum.empty()) { try { axis.scaling.minimum=std::stod(xlpp::internal::attribute(minimum.front(),"val")); axis.scaling.hasMinimum=true; } catch (...) {} }
            const auto maximum = drawingTags(scaling, "c:max", "max");
            if (!maximum.empty()) { try { axis.scaling.maximum=std::stod(xlpp::internal::attribute(maximum.front(),"val")); axis.scaling.hasMaximum=true; } catch (...) {} }
            const auto logBase = drawingTags(scaling, "c:logBase", "logBase");
            if (!logBase.empty()) { try { axis.scaling.logBase=std::stod(xlpp::internal::attribute(logBase.front(),"val")); axis.scaling.hasLogBase=true; } catch (...) {} }
            const auto orientation = drawingTags(scaling, "c:orientation", "orientation");
            if (!orientation.empty()) axis.scaling.reverseOrder = xlpp::internal::attribute(orientation.front(),"val") == "maxMin";
        }
        const auto titles = drawingTags(parsed.xml, "c:title", "title");
        if (!titles.empty()) {
            axis.title = drawingTagText(titles.front(), "a:t", "t");
            if (axis.title.empty()) axis.title = drawingTagText(titles.front(), "c:v", "v");
            if (axis.title.empty()) axis.title = drawingTagText(titles.front(), "c:f", "f");
            axis.titleRichText = parseChartRichText(titles.front());
            if (axis.title.empty() && axis.titleRichText.present) axis.title = axis.titleRichText.plainText();
        }
        const auto numFmt = drawingTags(parsed.xml, "c:numFmt", "numFmt");
        if (!numFmt.empty()) {
            axis.numberFormat = xlpp::internal::attribute(numFmt.front(), "formatCode");
            const auto linked = xlpp::internal::attribute(numFmt.front(), "sourceLinked");
            axis.numberFormatSourceLinked = linked.empty() || linked == "1" || linked == "true";
        }
        const auto majorTick = drawingTags(parsed.xml, "c:majorTickMark", "majorTickMark");
        if (!majorTick.empty()) axis.majorTickMark = xlpp::internal::attribute(majorTick.front(), "val");
        const auto minorTick = drawingTags(parsed.xml, "c:minorTickMark", "minorTickMark");
        if (!minorTick.empty()) axis.minorTickMark = xlpp::internal::attribute(minorTick.front(), "val");
        const auto tickPos = drawingTags(parsed.xml, "c:tickLblPos", "tickLblPos");
        if (!tickPos.empty()) axis.tickLabelPosition = xlpp::internal::attribute(tickPos.front(), "val");
        const auto majorUnit = drawingTags(parsed.xml, "c:majorUnit", "majorUnit");
        if (!majorUnit.empty()) { try { axis.majorUnit = std::stod(xlpp::internal::attribute(majorUnit.front(), "val")); axis.hasMajorUnit = true; } catch (...) {} }
        const auto minorUnit = drawingTags(parsed.xml, "c:minorUnit", "minorUnit");
        if (!minorUnit.empty()) { try { axis.minorUnit = std::stod(xlpp::internal::attribute(minorUnit.front(), "val")); axis.hasMinorUnit = true; } catch (...) {} }
        const auto crossesNode = drawingTags(parsed.xml, "c:crosses", "crosses");
        if (!crossesNode.empty()) axis.crosses = xlpp::internal::attribute(crossesNode.front(), "val");
        const auto crossesAtNode = drawingTags(parsed.xml, "c:crossesAt", "crossesAt");
        if (!crossesAtNode.empty()) { try { axis.crossesAt=std::stod(xlpp::internal::attribute(crossesAtNode.front(),"val")); axis.hasCrossesAt=true; } catch (...) {} }
        const auto crossBetweenNode = drawingTags(parsed.xml, "c:crossBetween", "crossBetween");
        if (!crossBetweenNode.empty()) axis.crossBetween = xlpp::internal::attribute(crossBetweenNode.front(), "val");
        const auto displayUnits = drawingTags(parsed.xml, "c:dispUnits", "dispUnits");
        if (!displayUnits.empty()) {
            axis.displayUnits.present = true;
            const auto builtIn = drawingTags(displayUnits.front(), "c:builtInUnit", "builtInUnit");
            if (!builtIn.empty()) axis.displayUnits.builtInUnit = xlpp::internal::attribute(builtIn.front(), "val");
            const auto custom = drawingTags(displayUnits.front(), "c:custUnit", "custUnit");
            if (!custom.empty()) { try { axis.displayUnits.customUnit=std::stod(xlpp::internal::attribute(custom.front(),"val")); axis.displayUnits.hasCustomUnit=true; } catch (...) {} }
            const auto labels = drawingTags(displayUnits.front(), "c:dispUnitsLbl", "dispUnitsLbl");
            if (!labels.empty()) { axis.displayUnits.showLabel=true; axis.displayUnits.labelRichText=parseChartRichText(labels.front()); }
        }
        const auto axisSpPr = axisDirectSpPr(parsed.xml);
        if (!axisSpPr.empty()) axis.lineFormat = parseChartLineFormat(axisSpPr);
        const auto majorGrid = drawingTags(parsed.xml, "c:majorGridlines", "majorGridlines");
        axis.hasMajorGridlines = !majorGrid.empty();
        if (!majorGrid.empty()) { const auto spPr=drawingTags(majorGrid.front(),"c:spPr","spPr"); if(!spPr.empty()) axis.majorGridlineFormat=parseChartLineFormat(spPr.front()); }
        const auto minorGrid = drawingTags(parsed.xml, "c:minorGridlines", "minorGridlines");
        axis.hasMinorGridlines = !minorGrid.empty();
        if (!minorGrid.empty()) { const auto spPr=drawingTags(minorGrid.front(),"c:spPr","spPr"); if(!spPr.empty()) axis.minorGridlineFormat=parseChartLineFormat(spPr.front()); }
        axis.secondary = axis.id != 0 && primaryAxisIds.find(axis.id) == primaryAxisIds.end() &&
            laterPlotAxisIds.find(axis.id) != laterPlotAxisIds.end();
        axes.push_back(std::move(axis));
    }
    return axes;
}

std::string axisTitleById(const std::vector<xlpp::Chart::Axis>& axes, std::uint64_t axisId) {
    const auto it = std::find_if(axes.begin(), axes.end(), [&](const auto& axis) { return axis.id == axisId; });
    return it == axes.end() ? std::string{} : it->title;
}

xlpp::Chart::Type parseChartType(const std::string& chartXmlText) {
    using Type = xlpp::Chart::Type;
    if (!drawingTags(chartXmlText, "c:lineChart", "lineChart").empty()) return Type::Line;
    if (!drawingTags(chartXmlText, "c:pieChart", "pieChart").empty()) return Type::Pie;
    if (!drawingTags(chartXmlText, "c:scatterChart", "scatterChart").empty()) return Type::Scatter;
    if (!drawingTags(chartXmlText, "c:doughnutChart", "doughnutChart").empty()) return Type::Doughnut;
    if (!drawingTags(chartXmlText, "c:radarChart", "radarChart").empty()) return Type::Radar;
    if (!drawingTags(chartXmlText, "c:areaChart", "areaChart").empty()) return Type::Area;
    if (!drawingTags(chartXmlText, "c:bubbleChart", "bubbleChart").empty()) return Type::Bubble;
    if (!drawingTags(chartXmlText, "c:stockChart", "stockChart").empty()) return Type::Stock;
    const auto projectedPie = drawingTags(chartXmlText, "c:ofPieChart", "ofPieChart");
    if (!projectedPie.empty()) {
        const auto kinds=drawingTags(projectedPie.front(),"c:ofPieType","ofPieType");
        return !kinds.empty() && xlpp::internal::attribute(kinds.front(),"val")=="bar" ? Type::BarOfPie : Type::PieOfPie;
    }
    if (!drawingTags(chartXmlText, "c:bar3DChart", "bar3DChart").empty()) return Type::Bar3D;
    if (!drawingTags(chartXmlText, "c:line3DChart", "line3DChart").empty()) return Type::Line3D;
    if (!drawingTags(chartXmlText, "c:area3DChart", "area3DChart").empty()) return Type::Area3D;
    if (!drawingTags(chartXmlText, "c:pie3DChart", "pie3DChart").empty()) return Type::Pie3D;
    if (!drawingTags(chartXmlText, "c:surfaceChart", "surfaceChart").empty()) return Type::Surface;
    if (!drawingTags(chartXmlText, "c:surface3DChart", "surface3DChart").empty()) return Type::Surface3D;
    return Type::Bar;
}

xlpp::Chart::Grouping parseChartGrouping(const std::string& chartXmlText, xlpp::Chart::Type type) {
    const char* prefixed = "c:barChart";
    const char* local = "barChart";
    switch (type) {
        case xlpp::Chart::Type::Line: prefixed = "c:lineChart"; local = "lineChart"; break;
        case xlpp::Chart::Type::Area: prefixed = "c:areaChart"; local = "areaChart"; break;
        case xlpp::Chart::Type::Bar3D: prefixed = "c:bar3DChart"; local = "bar3DChart"; break;
        case xlpp::Chart::Type::Line3D: prefixed = "c:line3DChart"; local = "line3DChart"; break;
        case xlpp::Chart::Type::Area3D: prefixed = "c:area3DChart"; local = "area3DChart"; break;
        default: break;
    }
    const auto chartTypeNodes = drawingTags(chartXmlText, prefixed, local);
    if (chartTypeNodes.empty()) return xlpp::Chart::Grouping::Standard;
    const auto groupingNodes = drawingTags(chartTypeNodes.front(), "c:grouping", "grouping");
    if (groupingNodes.empty()) return xlpp::Chart::Grouping::Standard;
    const auto value = xlpp::internal::attribute(groupingNodes.front(), "val");
    if (value == "stacked") return xlpp::Chart::Grouping::Stacked;
    if (value == "percentStacked") return xlpp::Chart::Grouping::PercentStacked;
    if (value == "clustered") return xlpp::Chart::Grouping::Clustered;
    return xlpp::Chart::Grouping::Standard;
}

std::string chartTitleText(const std::string& chartXmlText) {
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return {};
    const auto& chartNode = chartNodes.front();
    auto plotPosition = chartNode.find("<c:plotArea");
    if (plotPosition == std::string::npos) plotPosition = chartNode.find("<plotArea");
    for (const auto& titleNode : drawingTags(chartNode, "c:title", "title")) {
        const auto titlePosition = chartNode.find(titleNode);
        if (plotPosition != std::string::npos && titlePosition > plotPosition) continue;
        auto value = drawingTagText(titleNode, "a:t", "t");
        if (value.empty()) value = drawingTagText(titleNode, "c:v", "v");
        if (value.empty()) value = drawingTagText(titleNode, "c:f", "f");
        return value;
    }
    return {};
}

std::string axisTitleText(const std::string& chartXmlText, const char* prefixedAxis, const char* localAxis, std::size_t axisIndex) {
    const auto axes = drawingTags(chartXmlText, prefixedAxis, localAxis);
    if (axisIndex >= axes.size()) return {};
    const auto titles = drawingTags(axes[axisIndex], "c:title", "title");
    if (titles.empty()) return {};
    auto value = drawingTagText(titles.front(), "a:t", "t");
    if (value.empty()) value = drawingTagText(titles.front(), "c:v", "v");
    if (value.empty()) value = drawingTagText(titles.front(), "c:f", "f");
    return value;
}

xlpp::DrawingAnchorInfo parseChartAnchorInfo(const std::string& anchorNode,
                                             xlpp::DrawingAnchorType type,
                                             const std::string& graphicFrame) {
    xlpp::DrawingAnchorInfo info;
    info.type = type;
    info.editAs = xlpp::internal::attribute(anchorNode, "editAs");
    const auto fromNodes = drawingTags(anchorNode, "xdr:from", "from");
    if (!fromNodes.empty()) info.from = parseDrawingMarker(fromNodes.front());
    const auto toNodes = drawingTags(anchorNode, "xdr:to", "to");
    if (!toNodes.empty()) info.to = parseDrawingMarker(toNodes.front());
    const auto posNodes = drawingTags(anchorNode, "xdr:pos", "pos");
    if (!posNodes.empty()) {
        const auto x = xlpp::internal::attribute(posNodes.front(), "x");
        const auto y = xlpp::internal::attribute(posNodes.front(), "y");
        long long parsed = 0;
        if (!x.empty() && xlpp::internal::tryParseIntegerExact(x, parsed)) info.xEmu = parsed;
        if (!y.empty() && xlpp::internal::tryParseIntegerExact(y, parsed)) info.yEmu = parsed;
    }
    const auto anchorExt = drawingTags(anchorNode, "xdr:ext", "ext");
    if (!anchorExt.empty()) {
        const auto cx = xlpp::internal::attribute(anchorExt.front(), "cx");
        const auto cy = xlpp::internal::attribute(anchorExt.front(), "cy");
        long long parsed = 0;
        if (!cx.empty() && xlpp::internal::tryParseIntegerExact(cx, parsed)) info.widthEmu = parsed;
        if (!cy.empty() && xlpp::internal::tryParseIntegerExact(cy, parsed)) info.heightEmu = parsed;
    }
    if ((info.widthEmu <= 0 || info.heightEmu <= 0) && type == xlpp::DrawingAnchorType::TwoCell) {
        const auto frameExt = drawingTags(graphicFrame, "a:ext", "ext");
        if (!frameExt.empty()) {
            const auto cx = xlpp::internal::attribute(frameExt.front(), "cx");
            const auto cy = xlpp::internal::attribute(frameExt.front(), "cy");
            long long parsed = 0;
            if (!cx.empty() && xlpp::internal::tryParseIntegerExact(cx, parsed)) info.widthEmu = parsed;
            if (!cy.empty() && xlpp::internal::tryParseIntegerExact(cy, parsed)) info.heightEmu = parsed;
        }
    }
    return info;
}

} // namespace internal
} // namespace xlpp

