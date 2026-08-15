#pragma once

#include <XLPP/Chart/Chart.h>

#include <string>
#include <vector>

namespace xlpp::internal {

// Shared XML helpers for drawing/chart subtrees. Kept here so both Workbook.cpp
// and the chart serializers use one implementation instead of each re-parsing
// the same drawing/chart nodes.
std::vector<std::string> drawingTags(const std::string& xml, const char* prefixed, const char* local);
std::string drawingTagText(const std::string& xml, const char* prefixed, const char* local);

// Chart XML generation helpers. These are pure serializers over the public
// Chart model: they depend only on Chart.h types, xmlEscape and each other,
// not on Workbook internals. Keeping them in a dedicated translation unit
// gives the chart subsystem a clear generation boundary and lets the rest of
// the workbook code reuse the same chart XML without parsing Workbook.cpp.
bool generatedChartTypeUsesXYAxes(xlpp::Chart::Type type);
bool generatedChartTypeHasAxes(xlpp::Chart::Type type);
std::string chartSeriesCacheXml(const xlpp::ChartSeriesCache& cache, bool prefixed = true);
std::string chartView3DXml(const xlpp::ChartView3D& view, bool prefixed);

} // namespace xlpp::internal
