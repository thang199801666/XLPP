#pragma once
#include <XLPP/Worksheet/Worksheet.h>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <string>
#include <utility>
#include <vector>

namespace xlpp::internal::ooxml {
bool replaceSimpleElementText(std::string& xml, const char* prefixed, const char* local, const std::string& value);
void eraseChartCacheBlocks(std::string& xml);
bool patchSeriesReferenceContainer(std::string& seriesXml,
                                   const char* prefixedContainer,
                                   const char* localContainer,
                                   const std::string& reference);
std::string generatedChartTitleXml(const std::string& title, bool prefixed, bool strict);
std::string chartRichTextTxXml(const xlpp::ChartRichText& richText, bool prefixed, bool strict);
bool patchImportedChartTitleRichText(std::string& chartXmlText, const xlpp::ChartRichText& richText);
bool patchImportedChartTitle(std::string& chartXmlText, const std::string& title);
bool patchAxisTitleNode(std::string& axis, const std::string& chartXmlText, const std::string& title);
bool patchImportedAxisTitle(std::string& chartXmlText,
                            const char* prefixedAxis,
                            const char* localAxis,
                            std::size_t axisIndex,
                            const std::string& title);
bool patchImportedAxisTitleById(std::string& chartXmlText, std::uint64_t axisId, const std::string& title);
bool patchImportedChartStyle(std::string& chartXmlText, const std::string& style);
bool patchAxisNodeById(std::string& chartXmlText, std::uint64_t axisId,
                       const std::function<bool(std::string&)>& patcher);
bool patchImportedAxisTitleRichTextById(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartRichText& richText);
bool patchImportedAxisNumberFormat(std::string& chartXmlText, std::uint64_t axisId, const std::string& formatCode, bool sourceLinked);
bool patchImportedAxisTicks(std::string& chartXmlText, std::uint64_t axisId, const std::string& major, const std::string& minor, const std::string& labelPos);
bool patchImportedAxisUnits(std::string& chartXmlText, std::uint64_t axisId, double majorUnit, double minorUnit);
bool patchScalingValChild(std::string& scaling, const char* prefixed, const char* local, const std::string& value);
bool patchImportedAxisScaling(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartAxisScaling& scalingValue);
bool patchImportedAxisCrossesAt(std::string& chartXmlText, std::uint64_t axisId, double crossesAt, bool clear);
bool patchImportedAxisDisplayUnits(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartDisplayUnits* units);
bool patchImportedAxisCrossing(std::string& chartXmlText, std::uint64_t axisId, const std::string& crosses, const std::string& crossBetween);
bool patchImportedAxisLineFormat(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartLineFormat& format);
bool patchImportedAxisGridlineFormat(std::string& chartXmlText, std::uint64_t axisId, bool major, const xlpp::ChartLineFormat& format);
bool removeImportedAxisGridlines(std::string& chartXmlText, std::uint64_t axisId, bool major);
bool patchImportedAreaFormat(std::string& chartXmlText, bool chartArea, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill);
std::string chartManualLayoutXml(const xlpp::ChartManualLayout& layout, bool prefixed);
bool patchManualLayoutOwner(std::string& owner, const xlpp::ChartManualLayout& layout);
bool patchImportedPlotAreaLayout(std::string& chartXmlText, const xlpp::ChartManualLayout& layout);
bool patchImportedLegendLayout(std::string& chartXmlText, const xlpp::ChartManualLayout& layout);
bool patchImportedLegendOverlay(std::string& chartXmlText, bool overlay);
bool patchImportedLegendFormat(std::string& chartXmlText, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill);
bool patchImportedChartLegend(std::string& chartXmlText, bool show, const std::string& legendPosition);
std::string dataLabelsAggregateMask(const std::string& dLbls);
bool patchOrInsertAggregateDataLabelVal(std::string& dLbls, const char* prefixed, const char* local,
                                        const std::string& value, bool insertWhenMissing);
bool patchOrInsertAggregateDataLabelText(std::string& dLbls, const char* prefixed, const char* local,
                                         const std::string& value);
bool patchSeriesLineOrFill(std::string& chartXmlText, std::size_t seriesIndex,
                           const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill);
bool patchSeriesMarkerFormat(std::string& chartXmlText, std::size_t seriesIndex, const xlpp::ChartMarkerFormat& format);
std::string trendlineTypeValue(xlpp::ChartSeries::TrendlineType type);
std::string errorBarDirectionValue(xlpp::ChartSeries::ErrorBarDirection direction);
std::string errorBarTypeValue(xlpp::ChartSeries::ErrorBarType type);
std::string errorValueTypeValue(xlpp::ChartSeries::ErrorValueType type);
std::string formatChartDouble(double value);
bool patchTrendlineNode(std::string& trendlineXml, const xlpp::ChartSeries::Trendline& trendline);
std::string makeTrendlineXml(const xlpp::ChartSeries::Trendline& trendline, bool prefixed);
bool patchImportedChartSeriesTrendline(std::string& chartXmlText, std::size_t seriesIndex,
                                       std::size_t trendlineIndex,
                                       const xlpp::ChartSeries::Trendline* trendline,
                                       bool add);
bool patchImportedChartSeriesTrendlineLineFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                                   std::size_t trendlineIndex, const xlpp::ChartLineFormat& format);
bool patchImportedChartSeriesErrorBarsLineFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                                  xlpp::ChartSeries::ErrorBarDirection direction,
                                                  const xlpp::ChartLineFormat& format);
bool patchErrorBarReference(std::string& errorBarsXml, const char* prefixed, const char* local, const std::string& reference);
bool patchErrorBarsNode(std::string& errorBarsXml, const xlpp::ChartSeries::ErrorBars& errorBars);
std::string makeErrorBarsXml(const xlpp::ChartSeries::ErrorBars& errorBars, bool prefixed);
bool patchImportedChartSeriesErrorBars(std::string& chartXmlText, std::size_t seriesIndex,
                                       xlpp::ChartSeries::ErrorBarDirection direction,
                                       const xlpp::ChartSeries::ErrorBars* errorBars);
std::vector<std::pair<std::size_t, std::string>> patchablePlotNodesInOrder(const std::string& chartXmlText);
bool patchImportedChartSeriesDataLabels(std::string& chartXmlText, std::size_t seriesIndex, const xlpp::Chart::DataLabels& labels);
bool patchImportedChartPlotDataLabels(std::string& chartXmlText, std::size_t plotIndex, const xlpp::Chart::DataLabels& labels);
bool patchImportedChartView3D(std::string& chartXmlText, const xlpp::ChartView3D& view);
bool patchImportedChartWallFormat(std::string& chartXmlText, const char* prefixedName, const char* localName,
                                  const xlpp::ChartWallFormat& format);
bool patchImportedChartDataTable(std::string& chartXmlText, const xlpp::ChartDataTable* table);
std::size_t plotAuxiliaryInsertion(const std::string& plot, const char* local);
bool patchImportedChartPlotLineObject(std::string& chartXmlText, std::size_t plotIndex,
                                      const char* prefixed, const char* local,
                                      const xlpp::ChartLineFormat* format);
bool patchImportedChartPlotUpDownBars(std::string& chartXmlText, std::size_t plotIndex,
                                      const xlpp::ChartUpDownBars* bars);
bool patchLeaderLinesInDataLabels(std::string& dLbls, const xlpp::ChartLineFormat* format, bool remove);
bool patchImportedChartPlotSimpleValue(std::string& chartXmlText, std::size_t plotIndex,
                                       const char* prefixed, const char* local, const std::string& value,
                                       bool beforeSeries);
bool patchImportedChartProjectedPie(std::string& chartXmlText, std::size_t plotIndex, const xlpp::ChartProjectedPieOptions& options);
bool patchImportedChartLeaderLines(std::string& chartXmlText, bool plotLevel, std::size_t ownerIndex,
                                   const xlpp::ChartLineFormat* format, bool remove);
std::string dataLabelPointXml(const xlpp::ChartDataLabelPoint& label, bool prefixed);
bool patchDataLabelPointNode(std::string& dLbls, const xlpp::ChartDataLabelPoint& label, bool remove);
std::string plotDirectDataLabels(const std::string& plot);
bool patchImportedChartSeriesDataLabelPoint(std::string& chartXmlText, std::size_t seriesIndex,
                                            const xlpp::ChartDataLabelPoint& label, bool remove);
bool patchImportedChartSeriesDataLabelPointRichText(std::string& chartXmlText, std::size_t seriesIndex,
                                                    std::size_t pointIndex, const xlpp::ChartRichText& richText);
bool patchImportedChartSeriesDataPointFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                             const xlpp::ChartDataPointFormat& format, bool remove);
bool patchImportedChartPlotDataLabelPoint(std::string& chartXmlText, std::size_t plotIndex,
                                          const xlpp::ChartDataLabelPoint& label, bool remove);
bool patchImportedChartSeriesTitle(std::string& chartXmlText,
                                   std::size_t seriesIndex,
                                   const std::string& title);
bool patchCacheInReferenceOwner(std::string& owner, const xlpp::ChartSeriesCache* cache);
bool patchImportedChartSeriesCache(std::string& chartXmlText, std::size_t seriesIndex, int kind, const xlpp::ChartSeriesCache* cache);
bool patchImportedChartSeriesReferences(std::string& chartXmlText,
                                        std::size_t seriesIndex,
                                        const std::string& categoriesReference,
                                        const std::string& valuesReference);

} // namespace xlpp::internal::ooxml
