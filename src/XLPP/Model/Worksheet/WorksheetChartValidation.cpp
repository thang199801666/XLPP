#include "WorksheetChartValidation.h"
#include <algorithm>
#include <array>
#include <cmath>

namespace xlpp::internal::worksheet_chart_validation {
bool validDataLabelPosition(const std::string& position) {
    if (position.empty()) return true;
    static const std::array<const char*, 9> validPositions{"bestFit", "b", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t"};
    return std::find_if(validPositions.begin(), validPositions.end(), [&](const char* candidate) {
        return position == candidate;
    }) != validPositions.end();
}

bool validChartLineFormat(const ChartLineFormat& format) {
    return std::isfinite(format.widthPoints) && format.widthPoints >= 0.0 &&
           (!format.color.present() || !format.color.value.empty());
}

bool validChartFillFormat(const ChartFillFormat& format) {
    return !format.color.present() || !format.color.value.empty();
}

bool validChartSeriesCache(const ChartSeriesCache& cache) {
    return cache.valid(true);
}

bool validChartMarkerFormat(const ChartMarkerFormat& format) {
    if (!format.present) return true;
    if (format.size != 0 && (format.size < 2 || format.size > 72)) return false;
    if (!format.symbol.empty()) {
        static const std::array<const char*, 11> symbols{"circle", "dash", "diamond", "dot", "none", "picture", "plus", "square", "star", "triangle", "x"};
        if (std::find_if(symbols.begin(), symbols.end(), [&](const char* candidate) { return format.symbol == candidate; }) == symbols.end())
            return false;
    }
    return validChartLineFormat(format.line) && validChartFillFormat(format.fill);
}
} // namespace xlpp::internal::worksheet_chart_validation
