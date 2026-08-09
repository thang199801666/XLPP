#pragma once
#include <XLPP/Chart/Chart.h>
#include <string>

namespace xlpp::internal::worksheet_chart_validation {
bool validDataLabelPosition(const std::string& position);
bool validChartLineFormat(const ChartLineFormat& format);
bool validChartFillFormat(const ChartFillFormat& format);
bool validChartSeriesCache(const ChartSeriesCache& cache);
bool validChartMarkerFormat(const ChartMarkerFormat& format);
} // namespace xlpp::internal::worksheet_chart_validation
