#pragma once
#include <XLPP/Chart/Chart.h>
#include <string>
namespace xlpp::internal::ooxml {
std::string chartSolidFillXml(const xlpp::ChartColor& color, bool declareNamespace = false);
bool patchChartLineFormatInSpPr(std::string& spPr, const xlpp::ChartLineFormat& format);
bool patchChartFillFormatInSpPr(std::string& spPr, const xlpp::ChartFillFormat& format);
bool ensureChartSpPr(std::string& owner, std::string& spPr, const std::string& beforeXml = {});
bool patchNestedLineFormat(std::string& owner, const xlpp::ChartLineFormat& format);
bool patchShapeOwnerFormat(std::string& owner, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill);
bool patchMarkerFormatInOwner(std::string& owner, const xlpp::ChartMarkerFormat& format);
}
