#pragma once
#include <XLPP/Chart/Chart.h>
#include <XLPP/Worksheet/Drawings/Image.h>
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

namespace xlpp {
class Worksheet;

namespace internal {

class ZipArchive;

// Parsing helpers shared by the workbook reader. These functions interpret
// ChartML / DrawingML fragments into the public chart model; they carry no
// workbook state and are safe to call from any reader module.

std::string partExtension(const std::string& part);

// Resolves a package-internal relationship target (which may be absolute or
// contain ../ segments) against the base part into a canonical part name.
std::string resolvePackagePart(const std::string& basePart, std::string relativeTarget);

xlpp::DrawingMarker parseDrawingMarker(const std::string& markerXml);

// Loads worksheet images (and their anchor geometry) from a sheet XML part.
void loadImages(xlpp::Worksheet& ws, const std::string& sheetXml,
                const xlpp::internal::ZipArchive& z, const std::string& sheetPart);

xlpp::Chart::Type parseChartType(const std::string& chartXmlText);
xlpp::Chart::Grouping parseChartGrouping(const std::string& chartXmlText, xlpp::Chart::Type type);
std::vector<xlpp::Chart::Axis> parseChartAxes(const std::string& chartXmlText,
                                              const std::vector<xlpp::Chart::Plot>& plots);
std::string axisTitleById(const std::vector<xlpp::Chart::Axis>& axes, std::uint64_t axisId);
std::vector<xlpp::Chart::Plot> parseChartPlots(const std::string& chartXmlText);
xlpp::Chart::DataLabels parseChartDataLabels(const std::string& plotXml, bool directPlotChild = false);
std::string chartTitleText(const std::string& chartXmlText);
std::string axisTitleText(const std::string& chartXmlText, const char* prefixedAxis,
                          const char* localAxis, std::size_t axisIndex = 0);
bool chartBoolValue(const std::string& container, const char* prefixed, const char* local,
                    bool fallback = false);
std::string seriesDirectSpPr(const std::string& seriesXml);
std::string axisDirectSpPr(const std::string& axisXml);
xlpp::ChartSeriesCache parseChartSeriesCache(const std::string& container);
xlpp::ChartThemePalette parseChartThemePalette(const xlpp::internal::ZipArchive& z);
xlpp::ChartStyleResources parseChartStyleResources(const xlpp::internal::ZipArchive& z,
                                                   const std::string& chartPart);
xlpp::ChartRichText parseChartRichText(const std::string& owner);
xlpp::ChartTextStyle parseChartTextStyle(const std::string& owner);
std::vector<xlpp::ChartDataPointFormat> parseChartDataPoints(const std::string& seriesXml);
std::vector<xlpp::ChartSeries::Trendline> parseChartTrendlines(const std::string& seriesXml);
std::vector<xlpp::ChartSeries::ErrorBars> parseChartErrorBars(const std::string& seriesXml);
xlpp::ChartLineFormat parseChartLineFormat(const std::string& container);
xlpp::ChartFillFormat parseChartFillFormat(const std::string& container);
xlpp::ChartMarkerFormat parseChartMarkerFormat(const std::string& seriesXml);
xlpp::ChartDataTable parseChartDataTable(const std::string& plotArea);
xlpp::ChartUpDownBars parseChartUpDownBars(const std::string& plotXml);
xlpp::ChartManualLayout parseChartManualLayout(const std::string& owner);
xlpp::ChartWallFormat parseChartWallFormat(const std::string& chartXmlText,
                                           const char* prefixed, const char* local);
xlpp::ChartView3D parseChartView3D(const std::string& chartXmlText);
xlpp::DrawingAnchorInfo parseChartAnchorInfo(const std::string& anchorNode,
                                             xlpp::DrawingAnchorType type,
                                             const std::string& graphicFrame);

} // namespace internal
} // namespace xlpp
