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

bool patchImportedChartView3D(std::string& chartXmlText, const xlpp::ChartView3D& view) {
    if (!view.present) return false;
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chart = originalChart;
    const bool prefixed = chart.find("<c:") != std::string::npos;
    const auto generated = chartView3DXml(view, prefixed);
    const auto existing = drawingTags(chart, "c:view3D", "view3D");
    if (!existing.empty()) {
        const auto pos = chart.find(existing.front());
        if (pos == std::string::npos) return false;
        chart.replace(pos, existing.front().size(), generated);
    } else {
        std::size_t insertion = std::string::npos;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
                 {"c:floor","floor"},{"c:sideWall","sideWall"},{"c:backWall","backWall"},{"c:plotArea","plotArea"}}}) {
            const auto nodes = drawingTags(chart, pair.first, pair.second);
            if (!nodes.empty()) { insertion = chart.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = chart.rfind("</");
        if (insertion == std::string::npos) return false;
        chart.insert(insertion, generated);
    }
    const auto pos = chartXmlText.find(originalChart);
    if (pos == std::string::npos) return false;
    chartXmlText.replace(pos, originalChart.size(), chart);
    return true;
}

bool patchImportedChartWallFormat(std::string& chartXmlText, const char* prefixedName, const char* localName,
                                  const xlpp::ChartWallFormat& format) {
    if (!format.present) return false;
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chart = originalChart;
    const bool prefixed = chart.find("<c:") != std::string::npos;
    const auto c = prefixed ? "c:" : "";
    const auto existing = drawingTags(chart, prefixedName, localName);
    std::string wall = existing.empty()
        ? "<" + std::string(c) + localName + "></" + std::string(c) + localName + ">"
        : existing.front();
    if (format.hasThickness && !patchOrInsertValChild(wall, "c:thickness", "thickness", std::to_string(format.thickness), true)) return false;
    if ((format.line.present || format.fill.present) &&
        !patchShapeOwnerFormat(wall, format.line.present ? &format.line : nullptr, format.fill.present ? &format.fill : nullptr)) return false;
    if (!existing.empty()) {
        const auto pos = chart.find(existing.front());
        if (pos == std::string::npos) return false;
        chart.replace(pos, existing.front().size(), wall);
    } else {
        std::vector<std::pair<const char*, const char*>> later;
        const std::string name(localName);
        if (name == "floor") later = {{"c:sideWall","sideWall"},{"c:backWall","backWall"},{"c:plotArea","plotArea"}};
        else if (name == "sideWall") later = {{"c:backWall","backWall"},{"c:plotArea","plotArea"}};
        else later = {{"c:plotArea","plotArea"}};
        std::size_t insertion = std::string::npos;
        for (const auto& pair : later) {
            const auto nodes = drawingTags(chart, pair.first, pair.second);
            if (!nodes.empty()) { insertion = chart.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = chart.rfind("</");
        if (insertion == std::string::npos) return false;
        chart.insert(insertion, wall);
    }
    const auto pos = chartXmlText.find(originalChart);
    if (pos == std::string::npos) return false;
    chartXmlText.replace(pos, originalChart.size(), chart);
    return true;
}





bool patchImportedChartDataTable(std::string& chartXmlText, const xlpp::ChartDataTable* table) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return false;
    const auto originalPlotArea = plotAreas.front();
    auto plotArea = originalPlotArea;
    auto tables = drawingTags(plotArea, "c:dTable", "dTable");
    if (!table) {
        if (!tables.empty()) {
            const auto position = plotArea.find(tables.front());
            if (position == std::string::npos) return false;
            plotArea.erase(position, tables.front().size());
        }
    } else {
        const bool prefixed = plotArea.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string generated = "<" + std::string(c) + "dTable>";
        const auto addFlag = [&](const char* name, bool value) {
            generated += "<" + std::string(c) + name + " val=\"" + (value ? "1" : "0") + "\"/>";
        };
        addFlag("showHorzBorder", table->showHorizontalBorder);
        addFlag("showVertBorder", table->showVerticalBorder);
        addFlag("showOutline", table->showOutline);
        addFlag("showKeys", table->showLegendKeys);
        generated += "</" + std::string(c) + "dTable>";
        if ((table->line.present || table->fill.present) &&
            !patchShapeOwnerFormat(generated, table->line.present ? &table->line : nullptr,
                                   table->fill.present ? &table->fill : nullptr)) return false;
        if (table->textStyle.present) {
            const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
            const auto txPr = chartTextStyleTxPrXml(table->textStyle, prefixed, strict);
            const auto closing = generated.rfind(prefixed ? "</c:dTable>" : "</dTable>");
            if (closing == std::string::npos) return false;
            generated.insert(closing, txPr);
        }
        if (!tables.empty()) {
            const auto position = plotArea.find(tables.front());
            if (position == std::string::npos) return false;
            plotArea.replace(position, tables.front().size(), generated);
        } else {
            const auto directSpPr = plotAreaDirectSpPr(plotArea);
            std::size_t insertion = !directSpPr.empty() ? plotArea.find(directSpPr) : std::string::npos;
            if (insertion == std::string::npos) {
                const auto ext = drawingTags(plotArea, "c:extLst", "extLst");
                insertion = !ext.empty() ? plotArea.find(ext.front()) : plotArea.rfind("</");
            }
            if (insertion == std::string::npos) return false;
            plotArea.insert(insertion, generated);
        }
    }
    const auto position = chartXmlText.find(originalPlotArea);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlotArea.size(), plotArea);
    return true;
}

std::size_t plotAuxiliaryInsertion(const std::string& plot, const char* local) {
    std::vector<std::pair<const char*, const char*>> later;
    const std::string tag(local);
    if (tag == "dropLines") later = {{"c:hiLowLines","hiLowLines"},{"c:upDownBars","upDownBars"},{"c:marker","marker"},{"c:smooth","smooth"},{"c:axId","axId"}};
    else if (tag == "hiLowLines") later = {{"c:upDownBars","upDownBars"},{"c:marker","marker"},{"c:smooth","smooth"},{"c:axId","axId"}};
    else later = {{"c:marker","marker"},{"c:smooth","smooth"},{"c:axId","axId"}};
    for (const auto& item : later) {
        const auto nodes = drawingTags(plot, item.first, item.second);
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position != std::string::npos) return position;
        }
    }
    return plot.rfind("</");
}

bool patchImportedChartPlotLineObject(std::string& chartXmlText, std::size_t plotIndex,
                                      const char* prefixed, const char* local,
                                      const xlpp::ChartLineFormat* format) {
    const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plotNodes.size()) return false;
    const auto originalPlot = plotNodes[plotIndex].second;
    auto plot = originalPlot;
    auto nodes = drawingTags(plot, prefixed, local);
    if (!format) {
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.erase(position, nodes.front().size());
        }
    } else {
        std::string object;
        if (!nodes.empty()) object = nodes.front();
        else {
            const bool hasPrefix = plot.find("<c:") != std::string::npos;
            const auto c = hasPrefix ? "c:" : "";
            object = "<" + std::string(c) + local + "></" + std::string(c) + local + ">";
        }
        if (format->present && !patchNestedLineFormat(object, *format)) return false;
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.replace(position, nodes.front().size(), object);
        } else {
            const auto insertion = plotAuxiliaryInsertion(plot, local);
            if (insertion == std::string::npos) return false;
            plot.insert(insertion, object);
        }
    }
    const auto position = chartXmlText.find(originalPlot);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlot.size(), plot);
    return true;
}

bool patchImportedChartPlotUpDownBars(std::string& chartXmlText, std::size_t plotIndex,
                                      const xlpp::ChartUpDownBars* bars) {
    const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plotNodes.size()) return false;
    const auto originalPlot = plotNodes[plotIndex].second;
    auto plot = originalPlot;
    auto nodes = drawingTags(plot, "c:upDownBars", "upDownBars");
    if (!bars) {
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.erase(position, nodes.front().size());
        }
    } else {
        const bool prefixed = plot.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string object = "<" + std::string(c) + "upDownBars><" + std::string(c) + "gapWidth val=\"" +
                             std::to_string(bars->gapWidth) + "\"/>";
        auto makeBar = [&](const char* name, const xlpp::ChartLineFormat& line, const xlpp::ChartFillFormat& fill) {
            std::string bar = "<" + std::string(c) + name + "></" + std::string(c) + name + ">";
            if ((line.present || fill.present) && !patchShapeOwnerFormat(bar, line.present ? &line : nullptr, fill.present ? &fill : nullptr)) return std::string{};
            return bar;
        };
        const auto up = makeBar("upBars", bars->upLine, bars->upFill);
        const auto down = makeBar("downBars", bars->downLine, bars->downFill);
        if (up.empty() || down.empty()) return false;
        object += up + down + "</" + std::string(c) + "upDownBars>";
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.replace(position, nodes.front().size(), object);
        } else {
            const auto insertion = plotAuxiliaryInsertion(plot, "upDownBars");
            if (insertion == std::string::npos) return false;
            plot.insert(insertion, object);
        }
    }
    const auto position = chartXmlText.find(originalPlot);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlot.size(), plot);
    return true;
}

bool patchLeaderLinesInDataLabels(std::string& dLbls, const xlpp::ChartLineFormat* format, bool remove) {
    auto leaderLines = drawingTags(dLbls, "c:leaderLines", "leaderLines");
    if (remove) {
        if (!leaderLines.empty()) {
            const auto position = dLbls.find(leaderLines.front());
            if (position == std::string::npos) return false;
            dLbls.erase(position, leaderLines.front().size());
        }
        const auto mask = dataLabelsAggregateMask(dLbls);
        const bool exists = !drawingTags(mask, "c:showLeaderLines", "showLeaderLines").empty();
        if (exists && !patchOrInsertAggregateDataLabelVal(dLbls, "c:showLeaderLines", "showLeaderLines", "0", true)) return false;
        return true;
    }
    if (!format || !format->present) return false;
    if (!patchOrInsertAggregateDataLabelVal(dLbls, "c:showLeaderLines", "showLeaderLines", "1", true)) return false;
    std::string leader;
    if (!leaderLines.empty()) leader = leaderLines.front();
    else {
        const bool prefixed = dLbls.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        leader = "<" + std::string(c) + "leaderLines></" + std::string(c) + "leaderLines>";
    }
    if (!patchNestedLineFormat(leader, *format)) return false;
    if (!leaderLines.empty()) {
        const auto position = dLbls.find(leaderLines.front());
        if (position == std::string::npos) return false;
        dLbls.replace(position, leaderLines.front().size(), leader);
    } else {
        const auto ext = drawingTags(dLbls, "c:extLst", "extLst");
        const auto insertion = !ext.empty() ? dLbls.find(ext.front()) : dLbls.rfind("</");
        if (insertion == std::string::npos) return false;
        dLbls.insert(insertion, leader);
    }
    return true;
}

bool patchImportedChartPlotSimpleValue(std::string& chartXmlText, std::size_t plotIndex,
                                       const char* prefixed, const char* local, const std::string& value,
                                       bool beforeSeries) {
    const auto plotNodes=patchablePlotNodesInOrder(chartXmlText); if(plotIndex>=plotNodes.size()) return false;
    const auto original=plotNodes[plotIndex].second; auto plot=original; auto nodes=drawingTags(plot,prefixed,local);
    const bool hasPrefix=plot.find("<c:")!=std::string::npos; const auto c=hasPrefix?"c:":"";
    const std::string generated="<"+std::string(c)+local+" val=\""+xmlEscape(value)+"\"/>";
    if(!nodes.empty()){ const auto pos=plot.find(nodes.front()); if(pos==std::string::npos) return false; plot.replace(pos,nodes.front().size(),generated); }
    else {
        std::size_t insertion=std::string::npos;
        if(beforeSeries){ const auto ser=drawingTags(plot,"c:ser","ser"); if(!ser.empty()) insertion=plot.find(ser.front()); }
        else { const auto ext=drawingTags(plot,"c:extLst","extLst"); insertion=!ext.empty()?plot.find(ext.front()):plot.rfind("</"); }
        if (insertion == std::string::npos) return false;
        plot.insert(insertion, generated);
    }
    const auto pos=chartXmlText.find(original); if(pos==std::string::npos) return false; chartXmlText.replace(pos,original.size(),plot); return true;
}

bool patchImportedChartProjectedPie(std::string& chartXmlText, std::size_t plotIndex, const xlpp::ChartProjectedPieOptions& options) {
    const auto plotNodes=patchablePlotNodesInOrder(chartXmlText); if(plotIndex>=plotNodes.size()) return false;
    const auto original=plotNodes[plotIndex].second; auto plot=original;
    if(plot.find("ofPieChart")==std::string::npos) return false;
    const bool hasPrefix=plot.find("<c:")!=std::string::npos; const auto c=hasPrefix?"c:":"";
    for(const auto& tag:std::array<std::pair<const char*,const char*>,7>{{
        {"c:ofPieType","ofPieType"},{"c:gapWidth","gapWidth"},{"c:splitType","splitType"},{"c:splitPos","splitPos"},
        {"c:custSplit","custSplit"},{"c:secondPieSize","secondPieSize"},{"c:serLines","serLines"}}}){
        const auto nodes=drawingTags(plot,tag.first,tag.second);
        if(!nodes.empty()){ const auto pos=plot.find(nodes.front()); if(pos==std::string::npos) return false; plot.erase(pos,nodes.front().size()); }
    }
    std::ostringstream block;
    block<<"<"<<c<<"ofPieType val=\""<<xmlEscape(options.ofPieType)<<"\"/>";
    block<<"<"<<c<<"gapWidth val=\""<<options.gapWidth<<"\"/><"<<c<<"splitType val=\""<<xmlEscape(options.splitType)<<"\"/>";
    if(options.hasSplitPosition) block<<"<"<<c<<"splitPos val=\""<<options.splitPosition<<"\"/>";
    if(!options.customSplitPoints.empty()){ block<<"<"<<c<<"custSplit>"; for(const auto point:options.customSplitPoints) block<<"<"<<c<<"secondPiePt val=\""<<point<<"\"/>"; block<<"</"<<c<<"custSplit>"; }
    block<<"<"<<c<<"secondPieSize val=\""<<options.secondPlotSize<<"\"/>";
    if(options.hasSeriesLines){ std::string lines="<"+std::string(c)+"serLines></"+std::string(c)+"serLines>"; if(options.seriesLinesFormat.present&&!patchNestedLineFormat(lines,options.seriesLinesFormat)) return false; block<<lines; }
    const auto ext=drawingTags(plot,"c:extLst","extLst"); const auto insertion=!ext.empty()?plot.find(ext.front()):plot.rfind("</"); if(insertion==std::string::npos) return false; plot.insert(insertion,block.str());
    const auto pos=chartXmlText.find(original); if(pos==std::string::npos) return false; chartXmlText.replace(pos,original.size(),plot); return true;
}

bool patchImportedChartLeaderLines(std::string& chartXmlText, bool plotLevel, std::size_t ownerIndex,
                                   const xlpp::ChartLineFormat* format, bool remove) {
    if (plotLevel) {
        const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
        if (ownerIndex >= plotNodes.size()) return false;
        const auto originalPlot = plotNodes[ownerIndex].second;
        auto plot = originalPlot;
        auto labels = drawingTags(plot, "c:dLbls", "dLbls");
        const auto seriesNodes = drawingTags(plot, "c:ser", "ser");
        labels.erase(std::remove_if(labels.begin(), labels.end(), [&](const std::string& node) {
            return std::any_of(seriesNodes.begin(), seriesNodes.end(), [&](const std::string& series){ return series.find(node) != std::string::npos; });
        }), labels.end());
        if (labels.empty()) {
            if (remove) return true;
            const bool prefixed = plot.find("<c:") != std::string::npos;
            const auto c = prefixed ? "c:" : "";
            std::string dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
            if (!patchLeaderLinesInDataLabels(dLbls, format, false)) return false;
            const auto axisIds = drawingTags(plot, "c:axId", "axId");
            const auto insertion = !axisIds.empty() ? plot.find(axisIds.front()) : plot.rfind("</");
            if (insertion == std::string::npos) return false;
            plot.insert(insertion, dLbls);
        } else {
            auto dLbls = labels.front();
            if (!patchLeaderLinesInDataLabels(dLbls, format, remove)) return false;
            const auto position = plot.find(labels.front()); if (position == std::string::npos) return false;
            plot.replace(position, labels.front().size(), dLbls);
        }
        const auto position = chartXmlText.find(originalPlot); if (position == std::string::npos) return false;
        chartXmlText.replace(position, originalPlot.size(), plot); return true;
    }
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (ownerIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[ownerIndex];
    auto series = originalSeries;
    auto labels = drawingTags(series, "c:dLbls", "dLbls");
    if (labels.empty()) {
        if (remove) return true;
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
        if (!patchLeaderLinesInDataLabels(dLbls, format, false)) return false;
        std::size_t insertion = series.rfind("</");
        for (const auto& item : std::array<std::pair<const char*, const char*>, 6>{{{"c:trendline","trendline"},{"c:errBars","errBars"},{"c:cat","cat"},{"c:xVal","xVal"},{"c:val","val"},{"c:yVal","yVal"}}}) {
            const auto nodes = drawingTags(series, item.first, item.second);
            if (!nodes.empty()) { const auto pos=series.find(nodes.front()); if(pos!=std::string::npos) { insertion=pos; break; } }
        }
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    } else {
        auto dLbls = labels.front();
        if (!patchLeaderLinesInDataLabels(dLbls, format, remove)) return false;
        const auto position = series.find(labels.front()); if(position==std::string::npos) return false;
        series.replace(position, labels.front().size(), dLbls);
    }
    const auto position = chartXmlText.find(originalSeries); if(position==std::string::npos) return false;
    chartXmlText.replace(position, originalSeries.size(), series); return true;
}

std::string dataLabelPointXml(const xlpp::ChartDataLabelPoint& label, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::string xml = "<" + std::string(c) + "dLbl><" + std::string(c) + "idx val=\"" + std::to_string(label.index) + "\"/>";
    if (label.deleted) xml += "<" + std::string(c) + "delete val=\"1\"/>";
    if (!label.position.empty()) xml += "<" + std::string(c) + "dLblPos val=\"" + xmlEscape(label.position) + "\"/>";
    if (label.showLegendKey) xml += "<" + std::string(c) + "showLegendKey val=\"1\"/>";
    if (label.showValue) xml += "<" + std::string(c) + "showVal val=\"1\"/>";
    if (label.showCategoryName) xml += "<" + std::string(c) + "showCatName val=\"1\"/>";
    if (label.showSeriesName) xml += "<" + std::string(c) + "showSerName val=\"1\"/>";
    if (label.showPercent) xml += "<" + std::string(c) + "showPercent val=\"1\"/>";
    if (label.showBubbleSize) xml += "<" + std::string(c) + "showBubbleSize val=\"1\"/>";
    if (label.showLeaderLines) xml += "<" + std::string(c) + "showLeaderLines val=\"1\"/>";
    if (!label.separator.empty()) xml += "<" + std::string(c) + "separator>" + xmlEscape(label.separator) + "</" + std::string(c) + "separator>";
    xml += "</" + std::string(c) + "dLbl>";
    return xml;
}

bool patchDataLabelPointNode(std::string& dLbls, const xlpp::ChartDataLabelPoint& label, bool remove) {
    auto points = drawingTags(dLbls, "c:dLbl", "dLbl");
    const auto found = std::find_if(points.begin(), points.end(), [&](const auto& point) {
        const auto idx = drawingTags(point, "c:idx", "idx");
        if (idx.empty()) return false;
        try { return std::stoull(xlpp::internal::attribute(idx.front(), "val")) == label.index; } catch (...) { return false; }
    });
    if (remove) {
        if (found == points.end()) return false;
        const auto position = dLbls.find(*found);
        if (position == std::string::npos) return false;
        dLbls.erase(position, found->size());
        return true;
    }
    if (found == points.end()) {
        const bool prefixed = dLbls.find("<c:dLbls") != std::string::npos;
        const auto generated = dataLabelPointXml(label, prefixed);
        std::size_t insertion = dLbls.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 8>{{
                 {"c:delete", "delete"}, {"c:dLblPos", "dLblPos"}, {"c:showLegendKey", "showLegendKey"}, {"c:showVal", "showVal"},
                 {"c:showCatName", "showCatName"}, {"c:showSerName", "showSerName"}, {"c:showPercent", "showPercent"}, {"c:showBubbleSize", "showBubbleSize"}}}) {
            const auto nodes = drawingTags(dLbls, pair.first, pair.second);
            for (const auto& node : nodes) {
                if (std::any_of(points.begin(), points.end(), [&](const auto& point) { return point.find(node) != std::string::npos; })) continue;
                const auto pos = dLbls.find(node);
                if (pos != std::string::npos) { insertion = std::min(insertion, pos); break; }
            }
        }
        if (insertion == std::string::npos) return false;
        dLbls.insert(insertion, generated);
        return true;
    }

    auto patched = *found;
    const auto patchFlag = [&](const char* prefixed, const char* local, bool value) {
        const bool exists = !drawingTags(patched, prefixed, local).empty();
        return patchOrInsertValChild(patched, prefixed, local, value ? "1" : "0", value || exists);
    };
    if (!patchFlag("c:delete", "delete", label.deleted) ||
        !patchFlag("c:showLegendKey", "showLegendKey", label.showLegendKey) ||
        !patchFlag("c:showVal", "showVal", label.showValue) ||
        !patchFlag("c:showCatName", "showCatName", label.showCategoryName) ||
        !patchFlag("c:showSerName", "showSerName", label.showSeriesName) ||
        !patchFlag("c:showPercent", "showPercent", label.showPercent) ||
        !patchFlag("c:showBubbleSize", "showBubbleSize", label.showBubbleSize) ||
        !patchFlag("c:showLeaderLines", "showLeaderLines", label.showLeaderLines)) return false;
    if (!label.position.empty() && !patchOrInsertValChild(patched, "c:dLblPos", "dLblPos", label.position)) return false;
    if (!label.separator.empty()) {
        const auto separators = drawingTags(patched, "c:separator", "separator");
        const bool prefixed = patched.find("<c:dLbl") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const auto generated = "<" + std::string(c) + "separator>" + xmlEscape(label.separator) + "</" + std::string(c) + "separator>";
        if (!separators.empty()) {
            const auto pos = patched.find(separators.front()); if (pos == std::string::npos) return false;
            patched.replace(pos, separators.front().size(), generated);
        } else {
            const auto close = patched.rfind("</"); if (close == std::string::npos) return false;
            patched.insert(close, generated);
        }
    }
    const auto position = dLbls.find(*found);
    if (position == std::string::npos) return false;
    dLbls.replace(position, found->size(), patched);
    return true;
}

std::string plotDirectDataLabels(const std::string& plot) {
    auto labels = drawingTags(plot, "c:dLbls", "dLbls");
    const auto series = drawingTags(plot, "c:ser", "ser");
    labels.erase(std::remove_if(labels.begin(), labels.end(), [&](const auto& node) {
        return std::any_of(series.begin(), series.end(), [&](const auto& seriesNode) { return seriesNode.find(node) != std::string::npos; });
    }), labels.end());
    return labels.empty() ? std::string{} : labels.front();
}

bool patchImportedChartSeriesDataLabelPoint(std::string& chartXmlText, std::size_t seriesIndex,
                                            const xlpp::ChartDataLabelPoint& label, bool remove) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    const auto labelNodes = drawingTags(series, "c:dLbls", "dLbls");
    std::string dLbls = labelNodes.empty() ? std::string{} : labelNodes.front();
    if (dLbls.empty()) {
        if (remove) return false;
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!patchDataLabelPointNode(dLbls, label, remove)) return false;
    if (!labelNodes.empty()) {
        const auto position = series.find(labelNodes.front());
        if (position == std::string::npos) return false;
        series.replace(position, labelNodes.front().size(), dLbls);
    } else {
        std::size_t insertion = series.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 6>{{
                 {"c:trendline", "trendline"}, {"c:errBars", "errBars"}, {"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) { const auto pos = series.find(nodes.front()); if (pos != std::string::npos) insertion = std::min(insertion, pos); }
        }
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchImportedChartSeriesDataLabelPointRichText(std::string& chartXmlText, std::size_t seriesIndex,
                                                    std::size_t pointIndex, const xlpp::ChartRichText& richText) {
    if (!richText.present || richText.runs.empty()) return false;
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    auto labelNodes = drawingTags(series, "c:dLbls", "dLbls");
    std::string dLbls;
    if (labelNodes.empty()) {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
        xlpp::ChartDataLabelPoint point;
        point.index = pointIndex;
        if (!patchDataLabelPointNode(dLbls, point, false)) return false;
    } else dLbls = labelNodes.front();

    auto points = drawingTags(dLbls, "c:dLbl", "dLbl");
    auto found = std::find_if(points.begin(), points.end(), [&](const auto& point) {
        const auto indices = drawingTags(point, "c:idx", "idx");
        if (indices.empty()) return false;
        try { return std::stoull(xlpp::internal::attribute(indices.front(), "val")) == pointIndex; } catch (...) { return false; }
    });
    if (found == points.end()) {
        xlpp::ChartDataLabelPoint point;
        point.index = pointIndex;
        if (!patchDataLabelPointNode(dLbls, point, false)) return false;
        points = drawingTags(dLbls, "c:dLbl", "dLbl");
        found = std::find_if(points.begin(), points.end(), [&](const auto& pointXml) {
            const auto indices = drawingTags(pointXml, "c:idx", "idx");
            if (indices.empty()) return false;
            try { return std::stoull(xlpp::internal::attribute(indices.front(), "val")) == pointIndex; } catch (...) { return false; }
        });
        if (found == points.end()) return false;
    }
    auto pointXml = *found;
    const bool prefixed = pointXml.find("<c:dLbl") != std::string::npos;
    const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
    const auto txXml = chartRichTextTxXml(richText, prefixed, strict);
    const auto txNodes = drawingTags(pointXml, "c:tx", "tx");
    if (!txNodes.empty()) {
        const auto pos = pointXml.find(txNodes.front());
        if (pos == std::string::npos) return false;
        pointXml.replace(pos, txNodes.front().size(), txXml);
    } else {
        std::size_t insertion = pointXml.rfind("</");
        const auto positions = drawingTags(pointXml, "c:dLblPos", "dLblPos");
        if (!positions.empty()) insertion = pointXml.find(positions.front());
        if (insertion == std::string::npos) return false;
        pointXml.insert(insertion, txXml);
    }
    const auto pointPos = dLbls.find(*found);
    if (pointPos == std::string::npos) return false;
    dLbls.replace(pointPos, found->size(), pointXml);

    if (!labelNodes.empty()) {
        const auto pos = series.find(labelNodes.front());
        if (pos == std::string::npos) return false;
        series.replace(pos, labelNodes.front().size(), dLbls);
    } else {
        std::size_t insertion = series.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 6>{{
                 {"c:trendline", "trendline"}, {"c:errBars", "errBars"}, {"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) {
                const auto pos = series.find(nodes.front());
                if (pos != std::string::npos) insertion = std::min(insertion, pos);
            }
        }
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    }
    const auto seriesPos = chartXmlText.find(originalSeries);
    if (seriesPos == std::string::npos) return false;
    chartXmlText.replace(seriesPos, originalSeries.size(), series);
    return true;
}


bool patchImportedChartSeriesDataPointFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                             const xlpp::ChartDataPointFormat& format, bool remove) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    auto points = drawingTags(series, "c:dPt", "dPt");
    auto found = std::find_if(points.begin(), points.end(), [&](const auto& point) {
        const auto indices = drawingTags(point, "c:idx", "idx");
        if (indices.empty()) return false;
        try { return std::stoull(xlpp::internal::attribute(indices.front(), "val")) == format.index; } catch (...) { return false; }
    });
    if (remove) {
        if (found == points.end()) return false;
        const auto position = series.find(*found);
        if (position == std::string::npos) return false;
        series.erase(position, found->size());
    } else {
        std::string pointXml;
        if (found != points.end()) pointXml = *found;
        else {
            const bool prefixed = series.find("<c:ser") != std::string::npos;
            const auto c = prefixed ? "c:" : "";
            pointXml = "<" + std::string(c) + "dPt><" + std::string(c) + "idx val=\"" +
                       std::to_string(format.index) + "\"/></" + std::string(c) + "dPt>";
        }
        auto spPrNodes = drawingTags(pointXml, "c:spPr", "spPr");
        std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
        if ((format.line.present || format.fill.present) && spPr.empty()) {
            std::string before;
            const auto markerNodes = drawingTags(pointXml, "c:marker", "marker");
            if (!markerNodes.empty()) before = markerNodes.front();
            if (!ensureChartSpPr(pointXml, spPr, before)) return false;
        }
        if (!spPr.empty()) {
            auto patched = spPr;
            if (format.line.present && !patchChartLineFormatInSpPr(patched, format.line)) return false;
            if (format.fill.present && !patchChartFillFormatInSpPr(patched, format.fill)) return false;
            const auto spPos = pointXml.find(spPr);
            if (spPos == std::string::npos) return false;
            pointXml.replace(spPos, spPr.size(), patched);
        }
        if (format.marker.present && !patchMarkerFormatInOwner(pointXml, format.marker)) return false;

        if (found != points.end()) {
            const auto position = series.find(*found);
            if (position == std::string::npos) return false;
            series.replace(position, found->size(), pointXml);
        } else {
            std::size_t insertion = series.rfind("</");
            for (const auto& pair : std::array<std::pair<const char*, const char*>, 7>{{
                     {"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}, {"c:errBars", "errBars"},
                     {"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
                const auto nodes = drawingTags(series, pair.first, pair.second);
                if (!nodes.empty()) {
                    const auto pos = series.find(nodes.front());
                    if (pos != std::string::npos) insertion = std::min(insertion, pos);
                }
            }
            if (insertion == std::string::npos) return false;
            series.insert(insertion, pointXml);
        }
    }
    const auto seriesPosition = chartXmlText.find(originalSeries);
    if (seriesPosition == std::string::npos) return false;
    chartXmlText.replace(seriesPosition, originalSeries.size(), series);
    return true;
}

bool patchImportedChartPlotDataLabelPoint(std::string& chartXmlText, std::size_t plotIndex,
                                          const xlpp::ChartDataLabelPoint& label, bool remove) {
    const auto plots = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plots.size()) return false;
    const auto original = plots[plotIndex].second;
    auto plot = original;
    const auto existing = plotDirectDataLabels(plot);
    std::string dLbls = existing;
    if (dLbls.empty()) {
        if (remove) return false;
        const bool prefixed = plot.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!patchDataLabelPointNode(dLbls, label, remove)) return false;
    if (!existing.empty()) {
        const auto pos = plot.find(existing); if (pos == std::string::npos) return false;
        plot.replace(pos, existing.size(), dLbls);
    } else {
        const auto axisIds = drawingTags(plot, "c:axId", "axId");
        const auto insertion = !axisIds.empty() ? plot.find(axisIds.front()) : plot.rfind("</");
        if (insertion == std::string::npos) return false;
        plot.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), plot);
    return true;
}

bool patchImportedChartSeriesTitle(std::string& chartXmlText,
                                   std::size_t seriesIndex,
                                   const std::string& title) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    const auto txNodes = drawingTags(series, "c:tx", "tx");
    if (title.empty()) {
        if (!txNodes.empty()) {
            const auto position = series.find(txNodes.front());
            if (position == std::string::npos) return false;
            series.erase(position, txNodes.front().size());
        }
    } else {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const std::string generated = "<" + std::string(c) + "tx><" + std::string(c) + "v>" +
            xmlEscape(title) + "</" + std::string(c) + "v></" + std::string(c) + "tx>";
        if (!txNodes.empty()) {
            const auto position = series.find(txNodes.front());
            if (position == std::string::npos) return false;
            series.replace(position, txNodes.front().size(), generated);
        } else {
            const auto idxNodes = drawingTags(series, "c:idx", "idx");
            const auto orderNodes = drawingTags(series, "c:order", "order");
            std::size_t insertion = std::string::npos;
            if (!orderNodes.empty()) {
                const auto pos = series.find(orderNodes.front());
                if (pos != std::string::npos) insertion = pos + orderNodes.front().size();
            } else if (!idxNodes.empty()) {
                const auto pos = series.find(idxNodes.front());
                if (pos != std::string::npos) insertion = pos + idxNodes.front().size();
            }
            if (insertion == std::string::npos) {
                const auto open = series.find('>');
                if (open == std::string::npos) return false;
                insertion = open + 1;
            }
            series.insert(insertion, generated);
        }
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchCacheInReferenceOwner(std::string& owner, const xlpp::ChartSeriesCache* cache) {
    auto refs = drawingTags(owner, "c:numRef", "numRef");
    bool numericRef = true;
    if (refs.empty()) { refs = drawingTags(owner, "c:strRef", "strRef"); numericRef = false; }
    if (refs.empty()) return cache == nullptr;
    const auto originalRef = refs.front();
    auto ref = originalRef;
    if (!cache) { eraseChartCacheBlocks(ref); }
    else {
        if (!cache->present || cache->numeric != numericRef) return false;
        const auto existingNum = drawingTags(ref, "c:numCache", "numCache");
        const auto existingStr = drawingTags(ref, "c:strCache", "strCache");
        const auto existing = !existingNum.empty() ? existingNum.front() : (!existingStr.empty() ? existingStr.front() : std::string{});
        const bool prefixed = ref.find("<c:") != std::string::npos;
        const auto generated = chartSeriesCacheXml(*cache, prefixed);
        if (!existing.empty()) {
            const auto pos = ref.find(existing); if (pos == std::string::npos) return false; ref.replace(pos, existing.size(), generated);
        } else {
            const auto close = ref.rfind("</"); if (close == std::string::npos) return false; ref.insert(close, generated);
        }
    }
    const auto pos = owner.find(originalRef);
    if (pos == std::string::npos) return false;
    owner.replace(pos, originalRef.size(), ref);
    return true;
}

bool patchImportedChartSeriesCache(std::string& chartXmlText, std::size_t seriesIndex, int kind, const xlpp::ChartSeriesCache* cache) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    if (kind == 3) eraseChartCacheBlocks(series);
    else {
        const char* prefixed = kind == 0 ? "c:cat" : (kind == 1 ? "c:val" : "c:tx");
        const char* local = kind == 0 ? "cat" : (kind == 1 ? "val" : "tx");
        auto owners = drawingTags(series, prefixed, local);
        if (owners.empty() && kind == 0) owners = drawingTags(series, "c:xVal", "xVal");
        if (owners.empty() && kind == 1) owners = drawingTags(series, "c:yVal", "yVal");
        if (owners.empty()) return false;
        auto owner = owners.front();
        if (!patchCacheInReferenceOwner(owner, cache)) return false;
        const auto ownerPos = series.find(owners.front()); if (ownerPos == std::string::npos) return false;
        series.replace(ownerPos, owners.front().size(), owner);
    }
    const auto pos = chartXmlText.find(original); if (pos == std::string::npos) return false;
    chartXmlText.replace(pos, original.size(), series);
    return true;
}

bool patchImportedChartSeriesReferences(std::string& chartXmlText,
                                        std::size_t seriesIndex,
                                        const std::string& categoriesReference,
                                        const std::string& valuesReference) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto patched = original;
    bool categoriesOk = patchSeriesReferenceContainer(patched, "c:cat", "cat", categoriesReference);
    if (!categoriesOk) categoriesOk = patchSeriesReferenceContainer(patched, "c:xVal", "xVal", categoriesReference);
    bool valuesOk = patchSeriesReferenceContainer(patched, "c:val", "val", valuesReference);
    if (!valuesOk) valuesOk = patchSeriesReferenceContainer(patched, "c:yVal", "yVal", valuesReference);
    if (!categoriesOk || !valuesOk) return false;
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), patched);
    return true;
}



} // namespace xlpp::internal::ooxml
