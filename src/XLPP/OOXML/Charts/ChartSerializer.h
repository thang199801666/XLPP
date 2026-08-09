#pragma once
#include <XLPP/Chart/Chart.h>
#include <string>
namespace xlpp::internal::ooxml {
std::string chartSeriesCacheXml(const xlpp::ChartSeriesCache& cache, bool prefixed = true);
std::string serializeChart(const xlpp::Chart& chart, bool strict);
std::string serializeChartEx(const xlpp::Chart& chart, bool strict);
std::string chartView3DXml(const xlpp::ChartView3D& view, bool prefixed);
std::string chartTextStyleTxPrXml(const xlpp::ChartTextStyle& style, bool prefixed, bool strict);
}
