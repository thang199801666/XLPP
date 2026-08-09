#include "OOXML/Charts/ImportedChartEditorParts.h"
#include "OOXML/Charts/ChartFormatCodec.h"
#include "OOXML/Charts/ChartSerializer.h"
#include "OOXML/Charts/ChartMutationSupport.h"
#include "OOXML/Charts/ChartXmlSupport.h"
#include "OOXML/Common/DrawingXmlSupport.h"
#include "OOXML/Common/Namespaces.h"
#include "OOXML/Drawings/WorkbookDrawingAccess.h"
#include "Package/Xml/XmlUtilities.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <iomanip>
#include <limits>
#include <sstream>
#include <string>
#include <tuple>
#include <type_traits>
#include <utility>
#include <vector>

namespace xlpp::internal::ooxml {

bool replaceSimpleElementText(std::string& xml, const char* prefixed, const char* local, const std::string& value) {
    const auto nodes = drawingTags(xml, prefixed, local);
    if (nodes.empty()) return false;
    const auto& original = nodes.front();
    const auto openEnd = original.find('>');
    const auto closeBegin = original.rfind("</");
    if (openEnd == std::string::npos || closeBegin == std::string::npos || closeBegin < openEnd) return false;
    auto patched = original;
    patched.replace(openEnd + 1, closeBegin - openEnd - 1, xmlEscape(value));
    const auto position = xml.find(original);
    if (position == std::string::npos) return false;
    xml.replace(position, original.size(), patched);
    return true;
}

void eraseChartCacheBlocks(std::string& xml) {
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 2>{
            std::pair{"c:strCache", "strCache"}, std::pair{"c:numCache", "numCache"}}) {
        for (;;) {
            const auto nodes = drawingTags(xml, prefixed, local);
            if (nodes.empty()) break;
            const auto position = xml.find(nodes.front());
            if (position == std::string::npos) break;
            xml.erase(position, nodes.front().size());
        }
    }
}

bool patchSeriesReferenceContainer(std::string& seriesXml,
                                   const char* prefixedContainer,
                                   const char* localContainer,
                                   const std::string& reference) {
    const auto containers = drawingTags(seriesXml, prefixedContainer, localContainer);
    if (containers.empty()) return false;
    const auto original = containers.front();
    auto patched = original;
    if (!replaceSimpleElementText(patched, "c:f", "f", reference)) return false;
    eraseChartCacheBlocks(patched);
    const auto position = seriesXml.find(original);
    if (position == std::string::npos) return false;
    seriesXml.replace(position, original.size(), patched);
    return true;
}

std::string generatedChartTitleXml(const std::string& title, bool prefixed, bool strict) {
    const auto c = prefixed ? "c:" : "";
    const auto drawingMain = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                    : "http://schemas.openxmlformats.org/drawingml/2006/main";
    std::ostringstream xml;
    xml << '<' << c << "title><" << c << "tx><" << c << "rich xmlns:a=\"" << drawingMain
        << "\"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(title)
        << "</a:t></a:r></a:p></" << c << "rich></" << c << "tx><" << c
        << "overlay val=\"0\"/></" << c << "title>";
    return xml.str();
}


std::string chartRichTextTxXml(const xlpp::ChartRichText& richText, bool prefixed, bool strict) {
    const auto c = prefixed ? "c:" : "";
    const auto drawingMain = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                    : "http://schemas.openxmlformats.org/drawingml/2006/main";
    std::string xml = "<" + std::string(c) + "tx><" + std::string(c) + "rich xmlns:a=\"" + drawingMain +
                      "\"><a:bodyPr/><a:lstStyle/><a:p>";
    for (const auto& run : richText.runs) {
        xml += "<a:r>";
        if (run.bold || run.italic || run.fontSizePoints > 0.0 || !run.typeface.empty() || run.color.present()) {
            xml += "<a:rPr";
            if (run.bold) xml += " b=\"1\"";
            if (run.italic) xml += " i=\"1\"";
            if (run.fontSizePoints > 0.0)
                xml += " sz=\"" + std::to_string(static_cast<long long>(std::llround(run.fontSizePoints * 100.0))) + "\"";
            xml += ">";
            if (run.color.present()) xml += chartSolidFillXml(run.color, false);
            if (!run.typeface.empty()) xml += "<a:latin typeface=\"" + xmlEscape(run.typeface) + "\"/>";
            xml += "</a:rPr>";
        }
        xml += "<a:t>" + xmlEscape(run.text) + "</a:t></a:r>";
    }
    xml += "</a:p></" + std::string(c) + "rich></" + std::string(c) + "tx>";
    return xml;
}

bool patchImportedChartTitleRichText(std::string& chartXmlText, const xlpp::ChartRichText& richText) {
    if (!richText.present || richText.runs.empty()) return false;
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chartNode = originalChart;
    auto plotPosition = chartNode.find("<c:plotArea");
    if (plotPosition == std::string::npos) plotPosition = chartNode.find("<plotArea");
    if (plotPosition == std::string::npos) return false;
    std::string titleNode;
    for (const auto& candidate : drawingTags(chartNode, "c:title", "title")) {
        const auto position = chartNode.find(candidate);
        if (position < plotPosition) { titleNode = candidate; break; }
    }
    const bool prefixed = chartNode.find("<c:chart") != std::string::npos;
    const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
    const auto txXml = chartRichTextTxXml(richText, prefixed, strict);
    if (titleNode.empty()) {
        const auto c = prefixed ? "c:" : "";
        const auto generated = "<" + std::string(c) + "title>" + txXml + "<" + std::string(c) +
                               "overlay val=\"0\"/></" + std::string(c) + "title>";
        chartNode.insert(plotPosition, generated);
    } else {
        auto patchedTitle = titleNode;
        const auto txNodes = drawingTags(patchedTitle, "c:tx", "tx");
        if (!txNodes.empty()) {
            const auto position = patchedTitle.find(txNodes.front());
            if (position == std::string::npos) return false;
            patchedTitle.replace(position, txNodes.front().size(), txXml);
        } else {
            const auto close = patchedTitle.rfind("</");
            if (close == std::string::npos) return false;
            patchedTitle.insert(close, txXml);
        }
        const auto position = chartNode.find(titleNode);
        if (position == std::string::npos) return false;
        chartNode.replace(position, titleNode.size(), patchedTitle);
    }
    const auto position = chartXmlText.find(originalChart);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalChart.size(), chartNode);
    return true;
}

bool patchImportedChartTitle(std::string& chartXmlText, const std::string& title) {
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chartNode = originalChart;
    auto plotPosition = chartNode.find("<c:plotArea");
    if (plotPosition == std::string::npos) plotPosition = chartNode.find("<plotArea");
    if (plotPosition == std::string::npos) return false;

    std::string titleNode;
    for (const auto& candidate : drawingTags(chartNode, "c:title", "title")) {
        const auto position = chartNode.find(candidate);
        if (position < plotPosition) { titleNode = candidate; break; }
    }
    if (title.empty()) {
        if (!titleNode.empty()) {
            const auto position = chartNode.find(titleNode);
            if (position == std::string::npos) return false;
            chartNode.erase(position, titleNode.size());
        }
    } else if (!titleNode.empty()) {
        auto patchedTitle = titleNode;
        if (!replaceSimpleElementText(patchedTitle, "a:t", "t", title) &&
            !replaceSimpleElementText(patchedTitle, "c:v", "v", title)) {
            const auto prefixed = chartNode.find("<c:chart") != std::string::npos;
            const auto strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
            patchedTitle = generatedChartTitleXml(title, prefixed, strict);
        }
        const auto position = chartNode.find(titleNode);
        if (position == std::string::npos) return false;
        chartNode.replace(position, titleNode.size(), patchedTitle);
    } else {
        const auto prefixed = chartNode.find("<c:chart") != std::string::npos;
        const auto strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        chartNode.insert(plotPosition, generatedChartTitleXml(title, prefixed, strict));
    }
    const auto position = chartXmlText.find(originalChart);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalChart.size(), chartNode);
    return true;
}

bool patchAxisTitleNode(std::string& axis, const std::string& chartXmlText, const std::string& title) {
    const auto titleNodes = drawingTags(axis, "c:title", "title");
    if (title.empty()) {
        if (!titleNodes.empty()) {
            const auto position = axis.find(titleNodes.front());
            if (position == std::string::npos) return false;
            axis.erase(position, titleNodes.front().size());
        }
    } else if (!titleNodes.empty()) {
        auto patched = titleNodes.front();
        if (!replaceSimpleElementText(patched, "a:t", "t", title) &&
            !replaceSimpleElementText(patched, "c:v", "v", title)) {
            const bool prefixed = axis.find("<c:") != std::string::npos;
            const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
            patched = generatedChartTitleXml(title, prefixed, strict);
        }
        const auto position = axis.find(titleNodes.front());
        if (position == std::string::npos) return false;
        axis.replace(position, titleNodes.front().size(), patched);
    } else {
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        const auto generated = generatedChartTitleXml(title, prefixed, strict);
        std::size_t insertion = std::string::npos;
        for (const auto& candidate : std::array<std::pair<const char*, const char*>, 7>{
                 std::pair{"c:numFmt", "numFmt"}, std::pair{"c:majorTickMark", "majorTickMark"},
                 std::pair{"c:minorTickMark", "minorTickMark"}, std::pair{"c:tickLblPos", "tickLblPos"},
                 std::pair{"c:spPr", "spPr"}, std::pair{"c:txPr", "txPr"},
                 std::pair{"c:crossAx", "crossAx"}}) {
            const auto nodes = drawingTags(axis, candidate.first, candidate.second);
            if (nodes.empty()) continue;
            const auto position = axis.find(nodes.front());
            if (position != std::string::npos) insertion = std::min(insertion, position);
        }
        if (insertion == std::string::npos) {
            const auto close = axis.rfind("</");
            if (close == std::string::npos) return false;
            insertion = close;
        }
        axis.insert(insertion, generated);
    }
    return true;
}

bool patchImportedAxisTitle(std::string& chartXmlText,
                            const char* prefixedAxis,
                            const char* localAxis,
                            std::size_t axisIndex,
                            const std::string& title) {
    const auto axes = drawingTags(chartXmlText, prefixedAxis, localAxis);
    if (axisIndex >= axes.size()) return false;
    const auto originalAxis = axes[axisIndex];
    auto axis = originalAxis;
    if (!patchAxisTitleNode(axis, chartXmlText, title)) return false;
    const auto position = chartXmlText.find(originalAxis);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalAxis.size(), axis);
    return true;
}

bool patchImportedAxisTitleById(std::string& chartXmlText, std::uint64_t axisId, const std::string& title) {
    if (axisId == 0) return false;
    const auto tryAxis = [&](const char* prefixedAxis, const char* localAxis) -> bool {
        for (const auto& originalAxis : drawingTags(chartXmlText, prefixedAxis, localAxis)) {
            const auto ids = drawingTags(originalAxis, "c:axId", "axId");
            if (ids.empty()) continue;
            const auto value = xlpp::internal::attribute(ids.front(), "val");
            std::uint64_t parsed = 0;
            try { if (!value.empty()) parsed = std::stoull(value); } catch (...) { continue; }
            if (parsed != axisId) continue;
            auto axis = originalAxis;
            if (!patchAxisTitleNode(axis, chartXmlText, title)) return false;
            const auto position = chartXmlText.find(originalAxis);
            if (position == std::string::npos) return false;
            chartXmlText.replace(position, originalAxis.size(), axis);
            return true;
        }
        return false;
    };
    return tryAxis("c:catAx", "catAx") || tryAxis("c:valAx", "valAx") ||
           tryAxis("c:dateAx", "dateAx") || tryAxis("c:serAx", "serAx");
}

bool patchImportedChartStyle(std::string& chartXmlText, const std::string& style) {
    if (style.empty()) return false;
    const auto existing = drawingTags(chartXmlText, "c:style", "style");
    if (!existing.empty()) {
        auto node = existing.front();
        if (!patchOpeningTagAttribute(node, "val", style, false)) return false;
        const auto pos = chartXmlText.find(existing.front());
        if (pos == std::string::npos) return false;
        chartXmlText.replace(pos, existing.front().size(), node);
        return true;
    }
    const auto charts = drawingTags(chartXmlText, "c:chart", "chart");
    if (charts.empty()) return false;
    const auto pos = chartXmlText.find(charts.front());
    if (pos == std::string::npos) return false;
    const bool prefixed = chartXmlText.find("<c:chartSpace") != std::string::npos;
    chartXmlText.insert(pos, std::string("<") + (prefixed ? "c:" : "") + "style val=\"" + xmlEscape(style) + "\"/>");
    return true;
}


bool patchAxisNodeById(std::string& chartXmlText, std::uint64_t axisId,
                       const std::function<bool(std::string&)>& patcher) {
    if (axisId == 0) return false;
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
             {"c:catAx", "catAx"}, {"c:valAx", "valAx"}, {"c:dateAx", "dateAx"}, {"c:serAx", "serAx"}}}) {
        for (const auto& originalAxis : drawingTags(chartXmlText, pair.first, pair.second)) {
            const auto ids = drawingTags(originalAxis, "c:axId", "axId");
            if (ids.empty()) continue;
            std::uint64_t parsed = 0;
            try { const auto value=xlpp::internal::attribute(ids.front(),"val"); if(!value.empty()) parsed=std::stoull(value); } catch (...) { continue; }
            if (parsed != axisId) continue;
            auto axis = originalAxis;
            if (!patcher(axis)) return false;
            const auto position = chartXmlText.find(originalAxis);
            if (position == std::string::npos) return false;
            chartXmlText.replace(position, originalAxis.size(), axis);
            return true;
        }
    }
    return false;
}

bool patchImportedAxisTitleRichTextById(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartRichText& richText) {
    if (!richText.present || richText.runs.empty()) return false;
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const auto txXml = chartRichTextTxXml(richText, prefixed, strict);
        const auto titles = drawingTags(axis, "c:title", "title");
        if (!titles.empty()) {
            auto title = titles.front();
            const auto tx = drawingTags(title, "c:tx", "tx");
            if (!tx.empty()) { const auto pos=title.find(tx.front()); if(pos==std::string::npos) return false; title.replace(pos,tx.front().size(),txXml); }
            else { const auto close=title.rfind("</"); if(close==std::string::npos) return false; title.insert(close,txXml); }
            const auto pos=axis.find(titles.front()); if(pos==std::string::npos) return false; axis.replace(pos,titles.front().size(),title);
        } else {
            const auto generated="<"+std::string(c)+"title>"+txXml+"<"+std::string(c)+"overlay val=\"0\"/></"+std::string(c)+"title>";
            std::size_t insertion=axis.rfind("</");
            const auto cross=drawingTags(axis,"c:crossAx","crossAx"); if(!cross.empty()) insertion=axis.find(cross.front());
            if (insertion == std::string::npos) return false;
            axis.insert(insertion, generated);
        }
        return true;
    });
}

bool patchImportedAxisNumberFormat(std::string& chartXmlText, std::uint64_t axisId, const std::string& formatCode, bool sourceLinked) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        auto nodes=drawingTags(axis,"c:numFmt","numFmt");
        if(!nodes.empty()) { auto node=nodes.front(); if(!patchOpeningTagAttribute(node,"formatCode",formatCode,false) || !patchOpeningTagAttribute(node,"sourceLinked",sourceLinked?"1":"0",false)) return false; const auto pos=axis.find(nodes.front()); if(pos==std::string::npos) return false; axis.replace(pos,nodes.front().size(),node); return true; }
        const bool prefixed=axis.find("<c:")!=std::string::npos; const auto c=prefixed?"c:":"";
        std::size_t insertion=axis.rfind("</"); const auto major=drawingTags(axis,"c:majorTickMark","majorTickMark"); if(!major.empty()) insertion=axis.find(major.front()); const auto cross=drawingTags(axis,"c:crossAx","crossAx"); if(insertion==std::string::npos && !cross.empty()) insertion=axis.find(cross.front()); if(insertion==std::string::npos) return false;
        axis.insert(insertion,"<"+std::string(c)+"numFmt formatCode=\""+xmlEscape(formatCode)+"\" sourceLinked=\""+(sourceLinked?"1":"0")+"\"/>"); return true;
    });
}

bool patchImportedAxisTicks(std::string& chartXmlText, std::uint64_t axisId, const std::string& major, const std::string& minor, const std::string& labelPos) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        if(!major.empty() && !patchOrInsertValChild(axis,"c:majorTickMark","majorTickMark",major,true)) return false;
        if(!minor.empty() && !patchOrInsertValChild(axis,"c:minorTickMark","minorTickMark",minor,true)) return false;
        if(!labelPos.empty() && !patchOrInsertValChild(axis,"c:tickLblPos","tickLblPos",labelPos,true)) return false;
        return true;
    });
}

bool patchImportedAxisUnits(std::string& chartXmlText, std::uint64_t axisId, double majorUnit, double minorUnit) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        std::ostringstream major; major<<std::setprecision(15)<<majorUnit;
        if(!patchOrInsertValChild(axis,"c:majorUnit","majorUnit",major.str(),true)) return false;
        if(minorUnit>0.0) { std::ostringstream minor; minor<<std::setprecision(15)<<minorUnit; if(!patchOrInsertValChild(axis,"c:minorUnit","minorUnit",minor.str(),true)) return false; }
        return true;
    });
}

bool patchScalingValChild(std::string& scaling, const char* prefixed, const char* local, const std::string& value) {
    const auto nodes = drawingTags(scaling, prefixed, local);
    const bool usePrefix = scaling.find("<c:scaling") != std::string::npos;
    const auto c = usePrefix ? "c:" : "";
    const auto generated = "<" + std::string(c) + local + " val=\"" + xmlEscape(value) + "\"/>";
    if (!nodes.empty()) {
        const auto pos = scaling.find(nodes.front());
        if (pos == std::string::npos) return false;
        scaling.replace(pos, nodes.front().size(), generated);
        return true;
    }
    const auto ext = drawingTags(scaling, "c:extLst", "extLst");
    const auto insertion = !ext.empty() ? scaling.find(ext.front()) : scaling.rfind("</");
    if (insertion == std::string::npos) return false;
    scaling.insert(insertion, generated);
    return true;
}

bool patchImportedAxisScaling(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartAxisScaling& scalingValue) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        auto nodes = drawingTags(axis, "c:scaling", "scaling");
        std::string scaling;
        if (!nodes.empty()) scaling = nodes.front();
        else {
            const bool prefixed = axis.find("<c:") != std::string::npos;
            const auto c = prefixed ? "c:" : "";
            scaling = "<" + std::string(c) + "scaling></" + std::string(c) + "scaling>";
        }
        if (scalingValue.hasLogBase) {
            std::ostringstream value; value << std::setprecision(15) << scalingValue.logBase;
            if (!patchScalingValChild(scaling, "c:logBase", "logBase", value.str())) return false;
        } else removeDrawingChild(scaling, "c:logBase", "logBase");
        if (!patchScalingValChild(scaling, "c:orientation", "orientation", scalingValue.reverseOrder ? "maxMin" : "minMax")) return false;
        if (scalingValue.hasMaximum) {
            std::ostringstream value; value << std::setprecision(15) << scalingValue.maximum;
            if (!patchScalingValChild(scaling, "c:max", "max", value.str())) return false;
        } else removeDrawingChild(scaling, "c:max", "max");
        if (scalingValue.hasMinimum) {
            std::ostringstream value; value << std::setprecision(15) << scalingValue.minimum;
            if (!patchScalingValChild(scaling, "c:min", "min", value.str())) return false;
        } else removeDrawingChild(scaling, "c:min", "min");
        if (!nodes.empty()) {
            const auto pos = axis.find(nodes.front());
            if (pos == std::string::npos) return false;
            axis.replace(pos, nodes.front().size(), scaling);
        } else {
            const auto ids = drawingTags(axis, "c:axId", "axId");
            if (ids.empty()) return false;
            const auto pos = axis.find(ids.front());
            if (pos == std::string::npos) return false;
            axis.insert(pos + ids.front().size(), scaling);
        }
        return true;
    });
}

bool patchImportedAxisCrossesAt(std::string& chartXmlText, std::uint64_t axisId, double crossesAt, bool clear) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        removeDrawingChild(axis, "c:crossesAt", "crossesAt");
        if (clear) return true;
        removeDrawingChild(axis, "c:crosses", "crosses");
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::ostringstream value; value << std::setprecision(15) << crossesAt;
        const auto generated = "<" + std::string(c) + "crossesAt val=\"" + value.str() + "\"/>";
        std::size_t insertion = axis.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
                 {"c:crossBetween", "crossBetween"}, {"c:majorUnit", "majorUnit"}, {"c:minorUnit", "minorUnit"}, {"c:dispUnits", "dispUnits"}}}) {
            const auto following = drawingTags(axis, pair.first, pair.second);
            if (!following.empty()) { const auto pos=axis.find(following.front()); if(pos!=std::string::npos) insertion=std::min(insertion,pos); }
        }
        const auto ext = drawingTags(axis, "c:extLst", "extLst");
        if (!ext.empty()) { const auto pos=axis.find(ext.front()); if(pos!=std::string::npos) insertion=std::min(insertion,pos); }
        if (insertion == std::string::npos) return false;
        axis.insert(insertion, generated);
        return true;
    });
}

bool patchImportedAxisDisplayUnits(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartDisplayUnits* units) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        const auto existing = drawingTags(axis, "c:dispUnits", "dispUnits");
        if (!units) {
            if (!existing.empty()) { const auto pos=axis.find(existing.front()); if(pos==std::string::npos) return false; axis.erase(pos,existing.front().size()); }
            return true;
        }
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string generated = "<" + std::string(c) + "dispUnits>";
        if (units->hasCustomUnit) {
            std::ostringstream value; value << std::setprecision(15) << units->customUnit;
            generated += "<" + std::string(c) + "custUnit val=\"" + value.str() + "\"/>";
        } else generated += "<" + std::string(c) + "builtInUnit val=\"" + xmlEscape(units->builtInUnit) + "\"/>";
        if (units->showLabel) {
            generated += "<" + std::string(c) + "dispUnitsLbl>";
            if (units->labelRichText.present && !units->labelRichText.runs.empty()) generated += chartRichTextTxXml(units->labelRichText, prefixed, strict);
            generated += "</" + std::string(c) + "dispUnitsLbl>";
        }
        generated += "</" + std::string(c) + "dispUnits>";
        if (!existing.empty()) {
            const auto pos = axis.find(existing.front()); if(pos==std::string::npos) return false; axis.replace(pos,existing.front().size(),generated); return true;
        }
        const auto ext = drawingTags(axis, "c:extLst", "extLst");
        const auto insertion = !ext.empty() ? axis.find(ext.front()) : axis.rfind("</");
        if (insertion == std::string::npos) return false;
        axis.insert(insertion, generated);
        return true;
    });
}

bool patchImportedAxisCrossing(std::string& chartXmlText, std::uint64_t axisId, const std::string& crosses, const std::string& crossBetween) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        if(!crosses.empty()) removeDrawingChild(axis,"c:crossesAt","crossesAt");
        if(!crosses.empty() && !patchOrInsertValChild(axis,"c:crosses","crosses",crosses,true)) return false;
        if(!crossBetween.empty() && !patchOrInsertValChild(axis,"c:crossBetween","crossBetween",crossBetween,true)) return false;
        return true;
    });
}

bool patchImportedAxisLineFormat(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartLineFormat& format) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        auto spPr=axisDirectSpPr(axis); if(spPr.empty() && !ensureChartSpPr(axis,spPr,{})) return false; auto patched=spPr; if(!patchChartLineFormatInSpPr(patched,format)) return false; const auto pos=axis.find(spPr); if(pos==std::string::npos) return false; axis.replace(pos,spPr.size(),patched); return true;
    });
}

bool patchImportedAxisGridlineFormat(std::string& chartXmlText, std::uint64_t axisId, bool major, const xlpp::ChartLineFormat& format) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        const char* pref=major?"c:majorGridlines":"c:minorGridlines"; const char* local=major?"majorGridlines":"minorGridlines";
        auto grids=drawingTags(axis,pref,local); std::string grid;
        if(!grids.empty()) grid=grids.front(); else { const bool prefixed=axis.find("<c:")!=std::string::npos; const auto c=prefixed?"c:":""; grid="<"+std::string(c)+local+"></"+std::string(c)+local+">"; }
        if(!patchNestedLineFormat(grid,format)) return false;
        if(!grids.empty()) { const auto pos=axis.find(grids.front()); if(pos==std::string::npos) return false; axis.replace(pos,grids.front().size(),grid); }
        else { const auto title=drawingTags(axis,"c:title","title"); std::size_t pos=!title.empty()?axis.find(title.front()):axis.rfind("</"); if(pos==std::string::npos) return false; axis.insert(pos,grid); }
        return true;
    });
}

bool removeImportedAxisGridlines(std::string& chartXmlText, std::uint64_t axisId, bool major) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        removeDrawingChild(axis, major ? "c:majorGridlines" : "c:minorGridlines", major ? "majorGridlines" : "minorGridlines");
        return true;
    });
}

bool patchImportedAreaFormat(std::string& chartXmlText, bool chartArea, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    std::string owner;
    std::string originalOwner;
    if (chartArea) owner = chartXmlText;
    else {
        const auto plots = drawingTags(chartXmlText, "c:plotArea", "plotArea");
        if (plots.empty()) return false;
        originalOwner = plots.front(); owner = originalOwner;
    }
    auto spPr = chartArea ? chartSpaceDirectSpPr(owner) : plotAreaDirectSpPr(owner);
    if (spPr.empty()) {
        const bool prefixed = owner.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        spPr = "<" + std::string(c) + "spPr></" + std::string(c) + "spPr>";
        std::size_t insertion = std::string::npos;
        if (chartArea) {
            const auto charts = drawingTags(owner, "c:chart", "chart");
            if (charts.empty()) return false;
            const auto chartPos = owner.find(charts.front());
            if (chartPos == std::string::npos) return false;
            insertion = chartPos + charts.front().size();
        } else {
            const auto ext = drawingTags(owner, "c:extLst", "extLst");
            insertion = !ext.empty() ? owner.find(ext.back()) : owner.rfind("</");
        }
        if (insertion == std::string::npos) return false;
        owner.insert(insertion, spPr);
    }
    auto patched = spPr;
    if (line && !patchChartLineFormatInSpPr(patched, *line)) return false;
    if (fill && !patchChartFillFormatInSpPr(patched, *fill)) return false;
    const auto spPos = owner.find(spPr);
    if (spPos == std::string::npos) return false;
    owner.replace(spPos, spPr.size(), patched);
    if (chartArea) { chartXmlText = std::move(owner); return true; }
    const auto position = chartXmlText.find(originalOwner);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalOwner.size(), owner);
    return true;
}

std::string chartManualLayoutXml(const xlpp::ChartManualLayout& layout, bool prefixed) {
    const auto c=prefixed?"c:":""; std::ostringstream xml; xml<<"<"<<c<<"layout><"<<c<<"manualLayout>";
    const auto addText=[&](const char* tag,const std::string& value){ if(!value.empty()) xml<<"<"<<c<<tag<<" val=\""<<xmlEscape(value)<<"\"/>"; };
    const auto addNum=[&](const char* tag,bool has,double value){ if(has) xml<<"<"<<c<<tag<<" val=\""<<std::setprecision(15)<<value<<"\"/>"; };
    addText("layoutTarget",layout.target); addText("xMode",layout.xMode); addText("yMode",layout.yMode); addText("wMode",layout.widthMode); addText("hMode",layout.heightMode);
    addNum("x",layout.hasX,layout.x); addNum("y",layout.hasY,layout.y); addNum("w",layout.hasWidth,layout.width); addNum("h",layout.hasHeight,layout.height);
    xml<<"</"<<c<<"manualLayout></"<<c<<"layout>"; return xml.str();
}

bool patchManualLayoutOwner(std::string& owner, const xlpp::ChartManualLayout& layout) {
    if (!layout.present) return false;
    const bool prefixed = owner.find("<c:") != std::string::npos;
    const auto generated = chartManualLayoutXml(layout, prefixed);
    const auto layouts = drawingTags(owner, "c:layout", "layout");
    if(!layouts.empty()) { const auto pos=owner.find(layouts.front()); if(pos==std::string::npos) return false; owner.replace(pos,layouts.front().size(),generated); }
    else { const auto openEnd=owner.find('>'); if(openEnd==std::string::npos) return false; owner.insert(openEnd+1,generated); }
    return true;
}

bool patchImportedPlotAreaLayout(std::string& chartXmlText, const xlpp::ChartManualLayout& layout) {
    const auto plots=drawingTags(chartXmlText,"c:plotArea","plotArea"); if(plots.empty()) return false; auto plot=plots.front(); if(!patchManualLayoutOwner(plot,layout)) return false; const auto pos=chartXmlText.find(plots.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,plots.front().size(),plot); return true;
}

bool patchImportedLegendLayout(std::string& chartXmlText, const xlpp::ChartManualLayout& layout) {
    const auto legends=drawingTags(chartXmlText,"c:legend","legend"); if(legends.empty()) return false; auto legend=legends.front(); if(!patchManualLayoutOwner(legend,layout)) return false; const auto pos=chartXmlText.find(legends.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,legends.front().size(),legend); return true;
}

bool patchImportedLegendOverlay(std::string& chartXmlText, bool overlay) {
    const auto legends=drawingTags(chartXmlText,"c:legend","legend"); if(legends.empty()) return false; auto legend=legends.front(); if(!patchOrInsertValChild(legend,"c:overlay","overlay",overlay?"1":"0",true)) return false; const auto pos=chartXmlText.find(legends.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,legends.front().size(),legend); return true;
}

bool patchImportedLegendFormat(std::string& chartXmlText, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    const auto legends=drawingTags(chartXmlText,"c:legend","legend"); if(legends.empty()) return false; auto legend=legends.front(); auto spPrNodes=drawingTags(legend,"c:spPr","spPr"); std::string spPr=spPrNodes.empty()?std::string{}:spPrNodes.front(); if(spPr.empty() && !ensureChartSpPr(legend,spPr,{})) return false; auto patched=spPr; if(line && !patchChartLineFormatInSpPr(patched,*line)) return false; if(fill && !patchChartFillFormatInSpPr(patched,*fill)) return false; const auto spos=legend.find(spPr); if(spos==std::string::npos) return false; legend.replace(spos,spPr.size(),patched); const auto pos=chartXmlText.find(legends.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,legends.front().size(),legend); return true;
}

bool patchImportedChartLegend(std::string& chartXmlText, bool show, const std::string& legendPosition) {
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chart = originalChart;
    const auto legends = drawingTags(chart, "c:legend", "legend");
    if (!show) {
        if (!legends.empty()) {
            const auto position = chart.find(legends.front());
            if (position == std::string::npos) return false;
            chart.erase(position, legends.front().size());
        }
    } else if (!legends.empty()) {
        auto legend = legends.front();
        const auto positionNodes = drawingTags(legend, "c:legendPos", "legendPos");
        if (!positionNodes.empty()) {
            auto node = positionNodes.front();
            const auto val = xlpp::internal::attribute(node, "val");
            const auto attr = std::string("val=\"") + val + "\"";
            const auto attrPosition = node.find(attr);
            if (attrPosition == std::string::npos) return false;
            node.replace(attrPosition, attr.size(), "val=\"" + xmlEscape(legendPosition) + "\"");
            const auto nodePosition = legend.find(positionNodes.front());
            if (nodePosition == std::string::npos) return false;
            legend.replace(nodePosition, positionNodes.front().size(), node);
        } else {
            const bool prefixed = legend.find("<c:legend") != std::string::npos;
            const auto openEnd = legend.find('>');
            if (openEnd == std::string::npos) return false;
            const auto c = prefixed ? "c:" : "";
            legend.insert(openEnd + 1, "<" + std::string(c) + "legendPos val=\"" + xmlEscape(legendPosition) + "\"/>");
        }
        const auto position = chart.find(legends.front());
        if (position == std::string::npos) return false;
        chart.replace(position, legends.front().size(), legend);
    } else {
        const bool prefixed = chart.find("<c:chart") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const std::string legend = "<" + std::string(c) + "legend><" + std::string(c) +
            "legendPos val=\"" + xmlEscape(legendPosition) + "\"/><" + std::string(c) +
            "layout/><" + std::string(c) + "overlay val=\"0\"/></" + std::string(c) + "legend>";
        const auto plotAreas = drawingTags(chart, "c:plotArea", "plotArea");
        if (plotAreas.empty()) return false;
        const auto plotPosition = chart.find(plotAreas.front());
        if (plotPosition == std::string::npos) return false;
        chart.insert(plotPosition + plotAreas.front().size(), legend);
    }
    const auto position = chartXmlText.find(originalChart);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalChart.size(), chart);
    return true;
}




} // namespace xlpp::internal::ooxml
