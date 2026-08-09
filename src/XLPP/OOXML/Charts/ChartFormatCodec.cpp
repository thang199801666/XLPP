#include "OOXML/Charts/ChartFormatCodec.h"
#include "OOXML/Charts/ChartMutationSupport.h"
#include "OOXML/Charts/ChartXmlSupport.h"
#include "OOXML/Common/DrawingXmlSupport.h"
#include "Package/Xml/XmlUtilities.h"
#include <XLPP/Chart/Chart.h>
#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <iomanip>
#include <limits>
#include <sstream>
#include <string>
#include <utility>
#include <vector>


namespace xlpp::internal::ooxml {
const char* chartColorTransformTag(xlpp::ChartColorTransform::Kind kind) {
    using Kind = xlpp::ChartColorTransform::Kind;
    switch (kind) {
    case Kind::Alpha: return "alpha";
    case Kind::AlphaMod: return "alphaMod";
    case Kind::AlphaOff: return "alphaOff";
    case Kind::Tint: return "tint";
    case Kind::Shade: return "shade";
    case Kind::LumMod: return "lumMod";
    case Kind::LumOff: return "lumOff";
    case Kind::SatMod: return "satMod";
    case Kind::SatOff: return "satOff";
    }
    return "alpha";
}

std::string chartColorElement(const xlpp::ChartColor& color) {
    using Kind = xlpp::ChartColor::Kind;
    if (!color.present()) return {};
    const char* tag = "srgbClr";
    switch (color.kind) {
    case Kind::SRgb: tag = "srgbClr"; break;
    case Kind::Scheme: tag = "schemeClr"; break;
    case Kind::System: tag = "sysClr"; break;
    case Kind::Preset: tag = "prstClr"; break;
    case Kind::Unknown: tag = "srgbClr"; break;
    case Kind::None: return {};
    }
    if (color.transforms.empty())
        return "<a:" + std::string(tag) + " val=\"" + xmlEscape(color.value) + "\"/>";
    std::string xml = "<a:" + std::string(tag) + " val=\"" + xmlEscape(color.value) + "\">";
    for (const auto& transform : color.transforms)
        xml += "<a:" + std::string(chartColorTransformTag(transform.kind)) + " val=\"" + std::to_string(transform.value) + "\"/>";
    xml += "</a:" + std::string(tag) + ">";
    return xml;
}

std::string chartSolidFillXml(const xlpp::ChartColor& color, bool declareNamespace) {
    if (!color.present()) return {};
    return "<a:solidFill" + std::string(declareNamespace ? " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" : "") +
           ">" + chartColorElement(color) + "</a:solidFill>";
}

std::string chartGradientFillXml(const xlpp::ChartFillFormat& format, bool declareNamespace = false) {
    if (format.gradientStops.empty()) return {};
    std::string xml = "<a:gradFill" + std::string(declareNamespace ? " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" : "") + "><a:gsLst>";
    for (const auto& stop : format.gradientStops)
        if (stop.color.present())
            xml += "<a:gs pos=\"" + std::to_string(std::clamp(stop.position, 0, 100000)) + "\">" +
                   chartColorElement(stop.color) + "</a:gs>";
    xml += "</a:gsLst>";
    if (std::isfinite(format.gradientAngleDegrees) && std::abs(format.gradientAngleDegrees) > 1e-12) {
        const auto angle = static_cast<long long>(std::llround(format.gradientAngleDegrees * 60000.0));
        xml += "<a:lin ang=\"" + std::to_string(angle) + "\" scaled=\"1\"/>";
    }
    xml += "</a:gradFill>";
    return xml;
}

std::string chartPatternFillXml(const xlpp::ChartFillFormat& format, bool declareNamespace = false) {
    if (format.pattern.empty()) return {};
    std::string xml = "<a:pattFill" + std::string(declareNamespace ? " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" : "") +
                      " prst=\"" + xmlEscape(format.pattern) + "\">";
    if (format.foregroundColor.present()) xml += "<a:fgClr>" + chartColorElement(format.foregroundColor) + "</a:fgClr>";
    if (format.backgroundColor.present()) xml += "<a:bgClr>" + chartColorElement(format.backgroundColor) + "</a:bgClr>";
    xml += "</a:pattFill>";
    return xml;
}

std::string directSpPrFillNode(const std::string& spPr, const char* prefixed, const char* local) {
    const auto lines = drawingTags(spPr, "a:ln", "ln");
    for (const auto& candidate : drawingTags(spPr, prefixed, local)) {
        if (std::none_of(lines.begin(), lines.end(), [&](const auto& line) { return line.find(candidate) != std::string::npos; }))
            return candidate;
    }
    return {};
}

bool patchChartLineFormatInSpPr(std::string& spPr, const xlpp::ChartLineFormat& format) {
    auto lines = drawingTags(spPr, "a:ln", "ln");
    std::string line;
    if (!lines.empty()) line = lines.front();
    else line = "<a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"></a:ln>";

    if (format.widthPoints > 0.0) {
        const auto widthEmu = static_cast<long long>(std::llround(format.widthPoints * 12700.0));
        if (!patchOpeningTagAttribute(line, "w", std::to_string(widthEmu))) return false;
    }
    if (!patchOpeningTagAttribute(line, "cap", format.cap, true)) return false;
    if (!patchOpeningTagAttribute(line, "cmpd", format.compound, true)) return false;

    removeDrawingChild(line, "a:noFill", "noFill");
    removeDrawingChild(line, "a:solidFill", "solidFill");
    if (format.noFill) {
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        line.insert(close, "<a:noFill/>");
    } else if (format.color.present()) {
        const auto dashes = drawingTags(line, "a:prstDash", "prstDash");
        const auto customDashes = drawingTags(line, "a:custDash", "custDash");
        std::size_t insertion = line.rfind("</");
        if (!dashes.empty()) insertion = line.find(dashes.front());
        else if (!customDashes.empty()) insertion = line.find(customDashes.front());
        if (insertion == std::string::npos) return false;
        line.insert(insertion, chartSolidFillXml(format.color));
    }

    removeDrawingChild(line, "a:prstDash", "prstDash");
    removeDrawingChild(line, "a:custDash", "custDash");
    if (!format.customDash.empty()) {
        std::string custom = "<a:custDash>";
        for (const auto& stop : format.customDash) {
            const auto d = static_cast<long long>(std::llround(std::max(0.0, stop.dash) * 1000.0));
            const auto sp = static_cast<long long>(std::llround(std::max(0.0, stop.space) * 1000.0));
            custom += "<a:ds d=\"" + std::to_string(d) + "\" sp=\"" + std::to_string(sp) + "\"/>";
        }
        custom += "</a:custDash>";
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        line.insert(close, custom);
    } else if (!format.dash.empty()) {
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        line.insert(close, "<a:prstDash val=\"" + xmlEscape(format.dash) + "\"/>");
    }

    removeDrawingChild(line, "a:round", "round");
    removeDrawingChild(line, "a:bevel", "bevel");
    removeDrawingChild(line, "a:miter", "miter");
    if (!format.join.empty()) {
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        if (format.join == "round") line.insert(close, "<a:round/>");
        else if (format.join == "bevel") line.insert(close, "<a:bevel/>");
        else if (format.join == "miter") line.insert(close, "<a:miter/>");
        else return false;
    }

    if (!lines.empty()) {
        const auto position = spPr.find(lines.front());
        if (position == std::string::npos) return false;
        spPr.replace(position, lines.front().size(), line);
    } else {
        const auto close = spPr.rfind("</");
        if (close == std::string::npos) return false;
        spPr.insert(close, line);
    }
    return true;
}

bool patchChartFillFormatInSpPr(std::string& spPr, const xlpp::ChartFillFormat& format) {
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
             {"a:noFill", "noFill"}, {"a:solidFill", "solidFill"}, {"a:gradFill", "gradFill"}, {"a:pattFill", "pattFill"}}}) {
        const auto existing = directSpPrFillNode(spPr, pair.first, pair.second);
        if (!existing.empty()) {
            const auto position = spPr.find(existing);
            if (position != std::string::npos) spPr.erase(position, existing.size());
        }
    }
    std::string generated;
    const auto kind = format.noFill ? xlpp::ChartFillFormat::Kind::NoFill : format.kind;
    if (kind == xlpp::ChartFillFormat::Kind::NoFill)
        generated = "<a:noFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>";
    else if (kind == xlpp::ChartFillFormat::Kind::Gradient)
        generated = chartGradientFillXml(format, true);
    else if (kind == xlpp::ChartFillFormat::Kind::Pattern)
        generated = chartPatternFillXml(format, true);
    else if (format.color.present())
        generated = chartSolidFillXml(format.color, true);
    if (!generated.empty()) {
        const auto lines = drawingTags(spPr, "a:ln", "ln");
        const auto insertion = !lines.empty() ? spPr.find(lines.front()) : spPr.rfind("</");
        if (insertion == std::string::npos) return false;
        spPr.insert(insertion, generated);
    }
    return true;
}

bool ensureChartSpPr(std::string& owner, std::string& spPr, const std::string& beforeXml) {
    if (!spPr.empty()) return true;
    const bool prefixed = owner.find("<c:") != std::string::npos;
    const auto c = prefixed ? "c:" : "";
    spPr = "<" + std::string(c) + "spPr></" + std::string(c) + "spPr>";
    std::size_t insertion = std::string::npos;
    if (!beforeXml.empty()) insertion = owner.find(beforeXml);
    if (insertion == std::string::npos) insertion = owner.rfind("</");
    if (insertion == std::string::npos) return false;
    owner.insert(insertion, spPr);
    return true;
}

bool patchNestedLineFormat(std::string& owner, const xlpp::ChartLineFormat& format) {
    auto spPrNodes = drawingTags(owner, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if (spPr.empty() && !ensureChartSpPr(owner, spPr)) return false;
    auto patched = spPr;
    if (!patchChartLineFormatInSpPr(patched, format)) return false;
    const auto position = owner.find(spPr);
    if (position == std::string::npos) return false;
    owner.replace(position, spPr.size(), patched);
    return true;
}

bool patchShapeOwnerFormat(std::string& owner, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    auto spPrNodes = drawingTags(owner, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if (spPr.empty() && !ensureChartSpPr(owner, spPr)) return false;
    auto patched = spPr;
    if (line && line->present && !patchChartLineFormatInSpPr(patched, *line)) return false;
    if (fill && fill->present && !patchChartFillFormatInSpPr(patched, *fill)) return false;
    const auto position = owner.find(spPr);
    if (position == std::string::npos) return false;
    owner.replace(position, spPr.size(), patched);
    return true;
}

bool patchMarkerFormatInOwner(std::string& owner, const xlpp::ChartMarkerFormat& format) {
    auto markers = drawingTags(owner, "c:marker", "marker");
    std::string marker;
    if (!markers.empty()) marker = markers.front();
    else {
        const bool prefixed = owner.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        marker = "<" + std::string(c) + "marker></" + std::string(c) + "marker>";
    }
    if (!format.symbol.empty() && !patchOrInsertValChild(marker, "c:symbol", "symbol", format.symbol)) return false;
    if (format.size > 0 && !patchOrInsertValChild(marker, "c:size", "size", std::to_string(format.size))) return false;
    auto spPrNodes = drawingTags(marker, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if ((format.line.present || format.fill.present) && spPr.empty()) {
        if (!ensureChartSpPr(marker, spPr)) return false;
    }
    if (!spPr.empty()) {
        auto patched = spPr;
        if (format.line.present && !patchChartLineFormatInSpPr(patched, format.line)) return false;
        if (format.fill.present && !patchChartFillFormatInSpPr(patched, format.fill)) return false;
        const auto position = marker.find(spPr);
        if (position == std::string::npos) return false;
        marker.replace(position, spPr.size(), patched);
    }
    if (!markers.empty()) {
        const auto position = owner.find(markers.front());
        if (position == std::string::npos) return false;
        owner.replace(position, markers.front().size(), marker);
    } else {
        const auto close = owner.rfind("</");
        if (close == std::string::npos) return false;
        owner.insert(close, marker);
    }
    return true;
}

} // namespace xlpp::internal::ooxml
