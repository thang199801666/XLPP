#include "OOXML/Charts/ImportedChartXmlEditor.h"
#include "OOXML/Charts/ImportedChartEditorParts.h"
#include "OOXML/Drawings/WorkbookDrawingAccess.h"

#include <string>
#include <type_traits>

namespace xlpp::internal::ooxml {

bool importedChartEditRequiresXml(const xlpp::Worksheet& sheet, std::size_t editIndex) noexcept {
    const auto& edits = xlpp::internal::WorkbookDrawingAccess::chartEdits(sheet);
    if (editIndex >= edits.size()) return false;
    const auto& edit = edits[editIndex];
    return edit.titleChanged || edit.styleChanged || edit.titleRichTextChanged || edit.xAxisTitleChanged || edit.yAxisTitleChanged ||
                !edit.axisTitleEdits.empty() || !edit.axisRichTitleEdits.empty() || !edit.axisFormatEdits.empty() || !edit.areaFormatEdits.empty() || !edit.layoutEdits.empty() ||
                !edit.dataTableEdits.empty() || !edit.view3DEdits.empty() || !edit.wallFormatEdits.empty() || !edit.plotAuxiliaryEdits.empty() || !edit.plotTypeSpecificEdits.empty() || !edit.leaderLineEdits.empty() ||
                edit.legendChanged || edit.legendOverlayChanged || edit.legendLineFormatChanged || edit.legendFillFormatChanged ||
                !edit.seriesTitleEdits.empty() || !edit.seriesReferenceEdits.empty() || !edit.seriesCacheEdits.empty() ||
                !edit.plotDataLabelsEdits.empty() || !edit.seriesDataLabelsEdits.empty() || !edit.pointDataLabelEdits.empty() ||
                !edit.pointDataLabelRichTextEdits.empty() || !edit.dataPointFormatEdits.empty() ||
                !edit.seriesFormatEdits.empty() || !edit.trendlineFormatEdits.empty() || !edit.errorBarsFormatEdits.empty() ||
                !edit.trendlineEdits.empty() || !edit.errorBarsEdits.empty();
}

bool applyImportedChartXmlEdit(std::string& chartXml, const xlpp::Worksheet& sheet, std::size_t editIndex) {
    const auto& edits = xlpp::internal::WorkbookDrawingAccess::chartEdits(sheet);
    if (editIndex >= edits.size()) return false;
    const auto& edit = edits[editIndex];
            if (edit.titleChanged && !patchImportedChartTitle(chartXml, edit.title)) return false;
            if (edit.styleChanged && !patchImportedChartStyle(chartXml, edit.style)) return false;
            if (edit.titleRichTextChanged && !patchImportedChartTitleRichText(chartXml, edit.titleRichText)) return false;
            const bool xyValueAxes = edit.chartType == xlpp::Chart::Type::Scatter || edit.chartType == xlpp::Chart::Type::Bubble;
            if (edit.xAxisTitleChanged) {
                if (edit.primaryXAxisId != 0) {
                    if (!patchImportedAxisTitleById(chartXml, edit.primaryXAxisId, edit.xAxisTitle)) return false;
                } else if (!patchImportedAxisTitle(chartXml,
                        xyValueAxes ? "c:valAx" : "c:catAx", xyValueAxes ? "valAx" : "catAx", 0, edit.xAxisTitle)) return false;
            }
            if (edit.yAxisTitleChanged) {
                if (edit.primaryYAxisId != 0) {
                    if (!patchImportedAxisTitleById(chartXml, edit.primaryYAxisId, edit.yAxisTitle)) return false;
                } else if (!patchImportedAxisTitle(chartXml, "c:valAx", "valAx",
                        xyValueAxes ? 1 : 0, edit.yAxisTitle)) return false;
            }
            for (const auto& axisEdit : edit.axisTitleEdits)
                if (!patchImportedAxisTitleById(chartXml, axisEdit.axisId, axisEdit.title)) return false;
            for (const auto& richEdit : edit.axisRichTitleEdits)
                if (!patchImportedAxisTitleRichTextById(chartXml, richEdit.axisId, richEdit.richText)) return false;
            for (const auto& axisEdit : edit.axisFormatEdits) {
                using Kind = std::decay_t<decltype(axisEdit.kind)>;
                if (axisEdit.kind == Kind::NumberFormat) { if (!patchImportedAxisNumberFormat(chartXml, axisEdit.axisId, axisEdit.value1, axisEdit.flag)) return false; }
                else if (axisEdit.kind == Kind::Ticks) { if (!patchImportedAxisTicks(chartXml, axisEdit.axisId, axisEdit.value1, axisEdit.value2, axisEdit.value3)) return false; }
                else if (axisEdit.kind == Kind::Units) { if (!patchImportedAxisUnits(chartXml, axisEdit.axisId, axisEdit.number1, axisEdit.number2)) return false; }
                else if (axisEdit.kind == Kind::Scaling) { if (!patchImportedAxisScaling(chartXml, axisEdit.axisId, axisEdit.scaling)) return false; }
                else if (axisEdit.kind == Kind::Crossing) { if (!patchImportedAxisCrossing(chartXml, axisEdit.axisId, axisEdit.value1, axisEdit.value2)) return false; }
                else if (axisEdit.kind == Kind::CrossesAt) { if (!patchImportedAxisCrossesAt(chartXml, axisEdit.axisId, axisEdit.number1, false)) return false; }
                else if (axisEdit.kind == Kind::ClearCrossesAt) { if (!patchImportedAxisCrossesAt(chartXml, axisEdit.axisId, 0.0, true)) return false; }
                else if (axisEdit.kind == Kind::DisplayUnits) { if (!patchImportedAxisDisplayUnits(chartXml, axisEdit.axisId, &axisEdit.displayUnits)) return false; }
                else if (axisEdit.kind == Kind::ClearDisplayUnits) { if (!patchImportedAxisDisplayUnits(chartXml, axisEdit.axisId, nullptr)) return false; }
                else if (axisEdit.kind == Kind::Line) { if (!patchImportedAxisLineFormat(chartXml, axisEdit.axisId, axisEdit.line)) return false; }
                else if (axisEdit.kind == Kind::MajorGridline) { if (!patchImportedAxisGridlineFormat(chartXml, axisEdit.axisId, true, axisEdit.line)) return false; }
                else if (axisEdit.kind == Kind::MinorGridline) { if (!patchImportedAxisGridlineFormat(chartXml, axisEdit.axisId, false, axisEdit.line)) return false; }
                else if (axisEdit.kind == Kind::RemoveMajorGridline) { if (!removeImportedAxisGridlines(chartXml, axisEdit.axisId, true)) return false; }
                else if (!removeImportedAxisGridlines(chartXml, axisEdit.axisId, false)) return false;
            }
            for (const auto& areaEdit : edit.areaFormatEdits) {
                using Owner = std::decay_t<decltype(areaEdit.owner)>;
                using Kind = std::decay_t<decltype(areaEdit.kind)>;
                const bool chartArea = areaEdit.owner == Owner::ChartArea;
                if (areaEdit.kind == Kind::Line) { if (!patchImportedAreaFormat(chartXml, chartArea, &areaEdit.line, nullptr)) return false; }
                else if (!patchImportedAreaFormat(chartXml, chartArea, nullptr, &areaEdit.fill)) return false;
            }
            for (const auto& layoutEdit : edit.layoutEdits) {
                using Owner = std::decay_t<decltype(layoutEdit.owner)>;
                if (layoutEdit.owner == Owner::PlotArea) { if (!patchImportedPlotAreaLayout(chartXml, layoutEdit.layout)) return false; }
                else if (!patchImportedLegendLayout(chartXml, layoutEdit.layout)) return false;
            }
            for (const auto& tableEdit : edit.dataTableEdits)
                if (!patchImportedChartDataTable(chartXml, tableEdit.remove ? nullptr : &tableEdit.table)) return false;
            for (const auto& viewEdit : edit.view3DEdits)
                if (!patchImportedChartView3D(chartXml, viewEdit.view)) return false;
            for (const auto& wallEdit : edit.wallFormatEdits) {
                using Owner = std::decay_t<decltype(wallEdit.owner)>;
                if (wallEdit.owner == Owner::Floor) { if (!patchImportedChartWallFormat(chartXml, "c:floor", "floor", wallEdit.format)) return false; }
                else if (wallEdit.owner == Owner::SideWall) { if (!patchImportedChartWallFormat(chartXml, "c:sideWall", "sideWall", wallEdit.format)) return false; }
                else if (!patchImportedChartWallFormat(chartXml, "c:backWall", "backWall", wallEdit.format)) return false;
            }
            for (const auto& auxEdit : edit.plotAuxiliaryEdits) {
                using Kind = std::decay_t<decltype(auxEdit.kind)>;
                if (auxEdit.kind == Kind::DropLines) {
                    if (!patchImportedChartPlotLineObject(chartXml, auxEdit.plotIndex, "c:dropLines", "dropLines", auxEdit.remove ? nullptr : &auxEdit.line)) return false;
                } else if (auxEdit.kind == Kind::HighLowLines) {
                    if (!patchImportedChartPlotLineObject(chartXml, auxEdit.plotIndex, "c:hiLowLines", "hiLowLines", auxEdit.remove ? nullptr : &auxEdit.line)) return false;
                } else if (!patchImportedChartPlotUpDownBars(chartXml, auxEdit.plotIndex, auxEdit.remove ? nullptr : &auxEdit.upDownBars)) return false;
            }
            for (const auto& typeEdit : edit.plotTypeSpecificEdits) {
                using Kind = std::decay_t<decltype(typeEdit.kind)>;
                if (typeEdit.kind == Kind::FirstSliceAngle) { if (!patchImportedChartPlotSimpleValue(chartXml,typeEdit.plotIndex,"c:firstSliceAng","firstSliceAng",std::to_string(typeEdit.integerValue),false)) return false; }
                else if (typeEdit.kind == Kind::DoughnutHoleSize) { if (!patchImportedChartPlotSimpleValue(chartXml,typeEdit.plotIndex,"c:holeSize","holeSize",std::to_string(typeEdit.integerValue),false)) return false; }
                else if (typeEdit.kind == Kind::RadarStyle) { if (!patchImportedChartPlotSimpleValue(chartXml,typeEdit.plotIndex,"c:radarStyle","radarStyle",typeEdit.textValue,true)) return false; }
                else if (!patchImportedChartProjectedPie(chartXml,typeEdit.plotIndex,typeEdit.projectedPie)) return false;
            }
            for (const auto& leaderEdit : edit.leaderLineEdits)
                if (!patchImportedChartLeaderLines(chartXml, leaderEdit.plotLevel, leaderEdit.ownerIndex, leaderEdit.remove ? nullptr : &leaderEdit.line, leaderEdit.remove)) return false;
            if (edit.legendChanged && !patchImportedChartLegend(chartXml, edit.showLegend, edit.legendPosition)) return false;
            if (edit.legendOverlayChanged && !patchImportedLegendOverlay(chartXml, edit.legendOverlay)) return false;
            if (edit.legendLineFormatChanged && !patchImportedLegendFormat(chartXml, &edit.legendLineFormat, nullptr)) return false;
            if (edit.legendFillFormatChanged && !patchImportedLegendFormat(chartXml, nullptr, &edit.legendFillFormat)) return false;
            for (const auto& seriesEdit : edit.seriesTitleEdits)
                if (!patchImportedChartSeriesTitle(chartXml, seriesEdit.seriesIndex, seriesEdit.title)) return false;
            for (const auto& seriesEdit : edit.seriesReferenceEdits) {
                if (!patchImportedChartSeriesReferences(chartXml, seriesEdit.seriesIndex,
                                                        seriesEdit.categoriesReference,
                                                        seriesEdit.valuesReference)) return false;
            }
            for (const auto& cacheEdit : edit.seriesCacheEdits) {
                using Kind = std::decay_t<decltype(cacheEdit.kind)>;
                int kind = cacheEdit.kind == Kind::Categories ? 0 : (cacheEdit.kind == Kind::Values ? 1 : (cacheEdit.kind == Kind::Title ? 2 : 3));
                if (!patchImportedChartSeriesCache(chartXml, cacheEdit.seriesIndex, kind,
                                                   cacheEdit.kind == Kind::ClearAll ? nullptr : &cacheEdit.cache)) return false;
            }
            for (const auto& labelsEdit : edit.plotDataLabelsEdits)
                if (!patchImportedChartPlotDataLabels(chartXml, labelsEdit.plotIndex, labelsEdit.labels)) return false;
            for (const auto& labelsEdit : edit.seriesDataLabelsEdits)
                if (!patchImportedChartSeriesDataLabels(chartXml, labelsEdit.seriesIndex, labelsEdit.labels)) return false;
            for (const auto& pointEdit : edit.pointDataLabelEdits) {
                if (pointEdit.plotLevel) {
                    if (!patchImportedChartPlotDataLabelPoint(chartXml, pointEdit.ownerIndex, pointEdit.label, pointEdit.remove)) return false;
                } else if (!patchImportedChartSeriesDataLabelPoint(chartXml, pointEdit.ownerIndex, pointEdit.label, pointEdit.remove)) return false;
            }
            for (const auto& richEdit : edit.pointDataLabelRichTextEdits)
                if (!patchImportedChartSeriesDataLabelPointRichText(chartXml, richEdit.seriesIndex, richEdit.pointIndex, richEdit.richText)) return false;
            for (const auto& pointFormatEdit : edit.dataPointFormatEdits)
                if (!patchImportedChartSeriesDataPointFormat(chartXml, pointFormatEdit.seriesIndex, pointFormatEdit.format, pointFormatEdit.remove)) return false;
            for (const auto& formatEdit : edit.seriesFormatEdits) {
                using Kind = std::decay_t<decltype(formatEdit.kind)>;
                if (formatEdit.kind == Kind::Line) {
                    if (!patchSeriesLineOrFill(chartXml, formatEdit.seriesIndex, &formatEdit.line, nullptr)) return false;
                } else if (formatEdit.kind == Kind::Fill) {
                    if (!patchSeriesLineOrFill(chartXml, formatEdit.seriesIndex, nullptr, &formatEdit.fill)) return false;
                } else if (!patchSeriesMarkerFormat(chartXml, formatEdit.seriesIndex, formatEdit.marker)) return false;
            }
            for (const auto& trendlineEdit : edit.trendlineEdits) {
                using Action = std::decay_t<decltype(trendlineEdit.action)>;
                if (trendlineEdit.action == Action::Add) {
                    if (!patchImportedChartSeriesTrendline(chartXml, trendlineEdit.seriesIndex, trendlineEdit.trendlineIndex,
                                                           &trendlineEdit.trendline, true)) return false;
                } else if (trendlineEdit.action == Action::Remove) {
                    if (!patchImportedChartSeriesTrendline(chartXml, trendlineEdit.seriesIndex, trendlineEdit.trendlineIndex,
                                                           nullptr, false)) return false;
                } else if (!patchImportedChartSeriesTrendline(chartXml, trendlineEdit.seriesIndex, trendlineEdit.trendlineIndex,
                                                               &trendlineEdit.trendline, false)) return false;
            }
            for (const auto& barsEdit : edit.errorBarsEdits)
                if (!patchImportedChartSeriesErrorBars(chartXml, barsEdit.seriesIndex, barsEdit.direction,
                                                       barsEdit.remove ? nullptr : &barsEdit.errorBars)) return false;
            for (const auto& formatEdit : edit.trendlineFormatEdits)
                if (!patchImportedChartSeriesTrendlineLineFormat(chartXml, formatEdit.seriesIndex, formatEdit.trendlineIndex, formatEdit.line)) return false;
            for (const auto& formatEdit : edit.errorBarsFormatEdits)
                if (!patchImportedChartSeriesErrorBarsLineFormat(chartXml, formatEdit.seriesIndex, formatEdit.direction, formatEdit.line)) return false;
    return true;
}

} // namespace xlpp::internal::ooxml
