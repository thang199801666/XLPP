#include <XLPP/Workbook/Workbook.h>

#include <algorithm>
#include <cctype>
#include <cmath>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <vector>

namespace xlpp {
namespace {
bool asciiEqualNoCase(std::string_view lhs, std::string_view rhs) {
    if (lhs.size() != rhs.size()) return false;
    for (std::size_t i = 0; i < lhs.size(); ++i) {
        const auto a = static_cast<unsigned char>(lhs[i]);
        const auto b = static_cast<unsigned char>(rhs[i]);
        if (std::tolower(a) != std::tolower(b)) return false;
    }
    return true;
}
} // namespace

ChartStyleApplyReport Workbook::applyChartColorStyle(const std::string& worksheetName, const std::string& chartStableId,
                                                     bool applyFill, bool applyLine, bool applyMarker) {
    ChartStyleApplyReport report;
    auto* sheet = worksheet(worksheetName);
    if (!sheet) { report.warnings.push_back("worksheet not found: " + worksheetName); return report; }
    auto chartIt = std::find_if(sheet->charts_.begin(), sheet->charts_.end(), [&](const Chart& chart) {
        return chart.imported() && chart.stableId() == chartStableId;
    });
    if (chartIt == sheet->charts_.end()) { report.warnings.push_back("imported chart not found: " + chartStableId); return report; }
    const auto resolved = chartIt->styleResources().resolveColorStyle(chartIt->themePalette());
    report.colorsAvailable = resolved.size();
    report.fillStylesAvailable = chartIt->themePalette().effectScheme.fillStyles.size();
    report.lineStylesAvailable = chartIt->themePalette().effectScheme.lineStyles.size();
    report.effectStylesAvailable = chartIt->themePalette().effectScheme.effectStyles.size();
    if (resolved.empty()) { report.warnings.push_back("chart has no resolvable color-style entries"); return report; }
    const auto toColor = [](const ChartResolvedColor& resolvedColor) {
        ChartColor color;
        if (!resolvedColor.present) return color;
        color.kind = ChartColor::Kind::SRgb;
        color.value = resolvedColor.srgb();
        if (resolvedColor.alpha < 0.999999) {
            ChartColorTransform alpha; alpha.kind = ChartColorTransform::Kind::Alpha;
            alpha.value = static_cast<int>(std::lround(std::clamp(resolvedColor.alpha, 0.0, 1.0) * 100000.0));
            color.transforms.push_back(alpha);
        }
        return color;
    };
    for (std::size_t seriesIndex = 0; seriesIndex < chartIt->series().size(); ++seriesIndex) {
        ++report.seriesVisited;
        const auto& resolvedColor = resolved[seriesIndex % resolved.size()];
        if (!resolvedColor.present) continue;
        const auto color = toColor(resolvedColor);
        bool changed = false;
        if (applyLine) {
            auto line = chartIt->series()[seriesIndex].lineFormat();
            line.present = true; line.noFill = false; line.color = color;
            changed = sheet->setChartSeriesLineFormat(chartStableId, seriesIndex, line) || changed;
        }
        if (applyFill) {
            auto fill = chartIt->series()[seriesIndex].fillFormat();
            fill.present = true; fill.noFill = false; fill.kind = ChartFillFormat::Kind::Solid; fill.color = color;
            changed = sheet->setChartSeriesFillFormat(chartStableId, seriesIndex, fill) || changed;
        }
        if (applyMarker && chartIt->series()[seriesIndex].markerFormat().present) {
            auto marker = chartIt->series()[seriesIndex].markerFormat();
            marker.present = true;
            marker.fill.present = true; marker.fill.noFill = false; marker.fill.kind = ChartFillFormat::Kind::Solid; marker.fill.color = color;
            marker.line.present = true; marker.line.noFill = false; marker.line.color = color;
            changed = sheet->setChartSeriesMarkerFormat(chartStableId, seriesIndex, marker) || changed;
        }
        if (changed) ++report.seriesStyled;
    }
    return report;
}

ChartStyleApplyReport Workbook::applyChartThemeStyleMatrix(const std::string& worksheetName,
                                                              const std::string& chartStableId,
                                                              std::size_t fillStyleIndex,
                                                              std::size_t lineStyleIndex,
                                                              bool applyMarker) {
    ChartStyleApplyReport report;
    auto* sheet = worksheet(worksheetName);
    if (!sheet) { report.warnings.push_back("worksheet not found: " + worksheetName); return report; }
    auto chartIt = std::find_if(sheet->charts_.begin(), sheet->charts_.end(), [&](const Chart& chart) {
        return chart.imported() && chart.stableId() == chartStableId;
    });
    if (chartIt == sheet->charts_.end()) { report.warnings.push_back("imported chart not found: " + chartStableId); return report; }

    const auto& matrix = chartIt->themePalette().effectScheme;
    report.fillStylesAvailable = matrix.fillStyles.size();
    report.lineStylesAvailable = matrix.lineStyles.size();
    report.effectStylesAvailable = matrix.effectStyles.size();
    const auto seriesColors = chartIt->styleResources().resolveColorStyle(chartIt->themePalette());
    report.colorsAvailable = seriesColors.size();
    if (!matrix.present) { report.warnings.push_back("workbook theme has no format scheme"); return report; }
    if (fillStyleIndex >= matrix.fillStyles.size()) {
        report.warnings.push_back("theme fill-style index is out of range: " + std::to_string(fillStyleIndex));
        return report;
    }
    if (lineStyleIndex >= matrix.lineStyles.size()) {
        report.warnings.push_back("theme line-style index is out of range: " + std::to_string(lineStyleIndex));
        return report;
    }
    if (!matrix.fillStyles[fillStyleIndex].present) report.warnings.push_back("selected theme fill style could not be materialized");
    if (!matrix.lineStyles[lineStyleIndex].present) report.warnings.push_back("selected theme line style could not be materialized");
    if (!report.warnings.empty()) return report;

    const auto colorFromResolved = [](const ChartResolvedColor& resolved) {
        ChartColor color;
        if (!resolved.present) return color;
        color.kind = ChartColor::Kind::SRgb;
        color.value = resolved.srgb();
        if (resolved.alpha < 0.999999) {
            color.transforms.push_back({ChartColorTransform::Kind::Alpha,
                                        static_cast<int>(std::lround(std::clamp(resolved.alpha, 0.0, 1.0) * 100000.0))});
        }
        return color;
    };
    const auto materializeColor = [&](ChartColor& color, const ChartColor& placeholder) {
        if (color.kind == ChartColor::Kind::Scheme && asciiEqualNoCase(color.value, "phClr")) color = placeholder;
    };
    const auto materializeFill = [&](ChartFillFormat fill, const ChartColor& placeholder) {
        materializeColor(fill.color, placeholder);
        materializeColor(fill.foregroundColor, placeholder);
        materializeColor(fill.backgroundColor, placeholder);
        for (auto& stop : fill.gradientStops) materializeColor(stop.color, placeholder);
        return fill;
    };
    const auto materializeLine = [&](ChartLineFormat line, const ChartColor& placeholder) {
        materializeColor(line.color, placeholder);
        return line;
    };

    for (std::size_t seriesIndex = 0; seriesIndex < chartIt->series().size(); ++seriesIndex) {
        ++report.seriesVisited;
        ChartColor placeholder;
        if (!seriesColors.empty()) placeholder = colorFromResolved(seriesColors[seriesIndex % seriesColors.size()]);
        if (!placeholder.present()) {
            ChartColor accent; accent.kind = ChartColor::Kind::Scheme; accent.value = "accent1";
            placeholder = colorFromResolved(chartIt->themePalette().resolve(accent));
        }
        if (!placeholder.present()) {
            placeholder.kind = ChartColor::Kind::Scheme;
            placeholder.value = "accent1";
        }
        const auto fill = materializeFill(matrix.fillStyles[fillStyleIndex], placeholder);
        const auto line = materializeLine(matrix.lineStyles[lineStyleIndex], placeholder);
        bool changed = false;
        changed = sheet->setChartSeriesFillFormat(chartStableId, seriesIndex, fill) || changed;
        changed = sheet->setChartSeriesLineFormat(chartStableId, seriesIndex, line) || changed;
        if (applyMarker && chartIt->series()[seriesIndex].markerFormat().present) {
            auto marker = chartIt->series()[seriesIndex].markerFormat();
            marker.fill = fill;
            marker.line = line;
            changed = sheet->setChartSeriesMarkerFormat(chartStableId, seriesIndex, marker) || changed;
        }
        if (changed) ++report.seriesStyled;
    }
    return report;
}


ChartStyleApplyReport Workbook::applyChartStyleRules(const std::string& worksheetName,
                                                     const std::string& chartStableId,
                                                     bool applyMarkers) {
    ChartStyleApplyReport report;
    auto* sheet = worksheet(worksheetName);
    if (!sheet) { report.warnings.push_back("worksheet not found: " + worksheetName); return report; }
    auto chartIt = std::find_if(sheet->charts_.begin(), sheet->charts_.end(), [&](const Chart& chart) {
        return chart.imported() && chart.stableId() == chartStableId;
    });
    if (chartIt == sheet->charts_.end()) { report.warnings.push_back("imported chart not found: " + chartStableId); return report; }

    const auto resources = chartIt->styleResources();
    const auto theme = chartIt->themePalette();
    const auto seriesSnapshot = chartIt->series();
    const auto axesSnapshot = chartIt->axes();
    const auto plotsSnapshot = chartIt->plots();
    const auto dataTableSnapshot = chartIt->dataTable();
    const auto floorSnapshot = chartIt->floorFormat();
    const auto sideWallSnapshot = chartIt->sideWallFormat();
    const auto backWallSnapshot = chartIt->backWallFormat();
    const auto& matrix = theme.effectScheme;
    report.rulesAvailable = resources.chartStyleRules.size();
    report.colorsAvailable = resources.colorStyleColors.size();
    report.fillStylesAvailable = matrix.fillStyles.size() + matrix.backgroundFillStyles.size();
    report.lineStylesAvailable = matrix.lineStyles.size();
    report.effectStylesAvailable = matrix.effectStyles.size();
    report.seriesVisited = seriesSnapshot.size();
    if (!resources.chartStylePresent || resources.chartStyleRules.empty()) {
        report.warnings.push_back("chart has no parsed Office chart-style rules");
        return report;
    }
    if (!matrix.present) {
        report.warnings.push_back("workbook theme has no format scheme");
        return report;
    }

    const auto colorFromResolved = [](const ChartResolvedColor& resolved) {
        ChartColor color;
        if (!resolved.present) return color;
        color.kind = ChartColor::Kind::SRgb;
        color.value = resolved.srgb();
        if (resolved.alpha < 0.999999) {
            color.transforms.push_back({ChartColorTransform::Kind::Alpha,
                                        static_cast<int>(std::lround(std::clamp(resolved.alpha, 0.0, 1.0) * 100000.0))});
        }
        return color;
    };
    const auto applyTransforms = [](ChartColor color, const std::vector<ChartColorTransform>& transforms) {
        color.transforms.insert(color.transforms.end(), transforms.begin(), transforms.end());
        return color;
    };
    const auto resolveReferenceColor = [&](const ChartStyleReference& ref, std::size_t elementIndex) {
        if (ref.styleColor && !resources.colorStyleColors.empty()) {
            std::size_t colorIndex = elementIndex;
            if (!ref.styleColorValue.empty() && !asciiEqualNoCase(ref.styleColorValue, "auto")) {
                try { colorIndex = static_cast<std::size_t>(std::stoull(ref.styleColorValue)); }
                catch (...) { return ChartColor{}; }
            }
            if (colorIndex >= resources.colorStyleColors.size()) {
                if (asciiEqualNoCase(resources.colorStyleMethod, "cycle")) colorIndex %= resources.colorStyleColors.size();
                else return ChartColor{};
            }
            auto color = applyTransforms(resources.colorStyleColors[colorIndex], ref.styleColorTransforms);
            return colorFromResolved(theme.resolve(color));
        }
        if (ref.color.present()) return colorFromResolved(theme.resolve(ref.color));
        return ChartColor{};
    };
    const auto materializeColor = [](ChartColor& color, const ChartColor& placeholder) {
        if (color.kind == ChartColor::Kind::Scheme && asciiEqualNoCase(color.value, "phClr") && placeholder.present())
            color = placeholder;
    };
    const auto materializeFill = [&](ChartFillFormat fill, const ChartColor& placeholder) {
        materializeColor(fill.color, placeholder);
        materializeColor(fill.foregroundColor, placeholder);
        materializeColor(fill.backgroundColor, placeholder);
        for (auto& stop : fill.gradientStops) materializeColor(stop.color, placeholder);
        return fill;
    };
    const auto materializeLine = [&](ChartLineFormat line, const ChartColor& placeholder) {
        materializeColor(line.color, placeholder);
        return line;
    };
    const auto themeFillFor = [&](const ChartStyleReference& ref) -> std::optional<ChartFillFormat> {
        if (!ref.present || ref.index <= 0) return std::nullopt;
        if (ref.index >= 1001) {
            const auto index = static_cast<std::size_t>(ref.index - 1001);
            if (index < matrix.backgroundFillStyles.size()) return matrix.backgroundFillStyles[index];
            return std::nullopt;
        }
        const auto index = static_cast<std::size_t>(ref.index - 1);
        if (index < matrix.fillStyles.size()) return matrix.fillStyles[index];
        return std::nullopt;
    };
    const auto themeLineFor = [&](const ChartStyleReference& ref) -> std::optional<ChartLineFormat> {
        if (!ref.present || ref.index <= 0) return std::nullopt;
        const auto index = static_cast<std::size_t>(ref.index - 1);
        if (index < matrix.lineStyles.size()) return matrix.lineStyles[index];
        return std::nullopt;
    };
    const auto effectResolvable = [&](const ChartStyleReference& ref) {
        if (!ref.present || ref.index <= 0) return false;
        const auto index = static_cast<std::size_t>(ref.index - 1);
        return index < matrix.effectStyles.size();
    };

    struct ResolvedRuleFormat {
        bool hasFill{false};
        bool hasLine{false};
        ChartFillFormat fill{};
        ChartLineFormat line{};
    };
    const auto resolveRuleFormat = [&](const ChartStyleRule& rule, std::size_t elementIndex) {
        ResolvedRuleFormat resolved;
        const auto fillPlaceholder = resolveReferenceColor(rule.fillReference, elementIndex);
        const auto linePlaceholder = resolveReferenceColor(rule.lineReference, elementIndex);
        if (rule.shapeFill.present) {
            resolved.fill = materializeFill(rule.shapeFill, fillPlaceholder);
            resolved.hasFill = true;
        } else if (const auto fill = themeFillFor(rule.fillReference)) {
            resolved.fill = materializeFill(*fill, fillPlaceholder);
            resolved.hasFill = resolved.fill.present;
        }
        if (rule.shapeLine.present) {
            resolved.line = materializeLine(rule.shapeLine, linePlaceholder);
            resolved.hasLine = true;
        } else if (const auto line = themeLineFor(rule.lineReference)) {
            resolved.line = materializeLine(*line, linePlaceholder);
            resolved.hasLine = resolved.line.present;
        }
        if (resolved.hasLine && rule.hasLineWidthScale && resolved.line.widthPoints > 0.0)
            resolved.line.widthPoints *= rule.lineWidthScale;
        return resolved;
    };

    std::set<std::size_t> styledSeries;
    const auto markTarget = [&](bool changed) {
        if (changed) ++report.targetsStyled;
        return changed;
    };

    for (const auto& rule : resources.chartStyleRules) {
        ++report.rulesVisited;
        bool ruleChanged = false;
        if (effectResolvable(rule.effectReference)) ++report.effectReferencesResolved;

        if (rule.target == "chartArea") {
            const auto format = resolveRuleFormat(rule, 0);
            if (format.hasFill) ruleChanged = sheet->setChartAreaFillFormat(chartStableId, format.fill) || ruleChanged;
            if (format.hasLine) ruleChanged = sheet->setChartAreaLineFormat(chartStableId, format.line) || ruleChanged;
            markTarget(ruleChanged);
        } else if (rule.target == "plotArea" || rule.target == "plotArea3D") {
            const auto format = resolveRuleFormat(rule, 0);
            if (format.hasFill) ruleChanged = sheet->setChartPlotAreaFillFormat(chartStableId, format.fill) || ruleChanged;
            if (format.hasLine) ruleChanged = sheet->setChartPlotAreaLineFormat(chartStableId, format.line) || ruleChanged;
            markTarget(ruleChanged);
        } else if (rule.target == "legend") {
            const auto format = resolveRuleFormat(rule, 0);
            if (format.hasFill) ruleChanged = sheet->setChartLegendFillFormat(chartStableId, format.fill) || ruleChanged;
            if (format.hasLine) ruleChanged = sheet->setChartLegendLineFormat(chartStableId, format.line) || ruleChanged;
            markTarget(ruleChanged);
        } else if (rule.target == "dataPoint" || rule.target == "dataPointLine") {
            for (std::size_t i = 0; i < seriesSnapshot.size(); ++i) {
                const auto format = resolveRuleFormat(rule, i);
                bool changed = false;
                if (rule.target == "dataPoint" && format.hasFill)
                    changed = sheet->setChartSeriesFillFormat(chartStableId, i, format.fill) || changed;
                if (format.hasLine)
                    changed = sheet->setChartSeriesLineFormat(chartStableId, i, format.line) || changed;
                if (changed) { styledSeries.insert(i); ++report.targetsStyled; ruleChanged = true; }
            }
        } else if (rule.target == "dataPointMarker" && applyMarkers) {
            for (std::size_t i = 0; i < seriesSnapshot.size(); ++i) {
                auto marker = seriesSnapshot[i].markerFormat();
                if (!marker.present && !resources.markerLayout.present) continue;
                if (resources.markerLayout.present) {
                    marker.present = true;
                    if (!resources.markerLayout.symbol.empty()) marker.symbol = resources.markerLayout.symbol;
                    if (resources.markerLayout.size > 0) marker.size = resources.markerLayout.size;
                }
                const auto format = resolveRuleFormat(rule, i);
                if (format.hasFill) marker.fill = format.fill;
                if (format.hasLine) marker.line = format.line;
                if (!format.hasFill && !format.hasLine && !resources.markerLayout.present) continue;
                const bool changed = sheet->setChartSeriesMarkerFormat(chartStableId, i, marker);
                if (changed) { styledSeries.insert(i); ++report.targetsStyled; ruleChanged = true; }
            }
        } else if (rule.target == "categoryAxis" || rule.target == "valueAxis" || rule.target == "seriesAxis") {
            for (const auto& axis : axesSnapshot) {
                const bool matches = (rule.target == "categoryAxis" && (axis.kind == Chart::AxisKind::Category || axis.kind == Chart::AxisKind::Date)) ||
                                     (rule.target == "valueAxis" && axis.kind == Chart::AxisKind::Value) ||
                                     (rule.target == "seriesAxis" && axis.kind == Chart::AxisKind::Series);
                if (!matches) continue;
                const auto format = resolveRuleFormat(rule, 0);
                if (format.hasLine && sheet->setChartAxisLineFormat(chartStableId, axis.id, format.line)) {
                    ++report.targetsStyled; ruleChanged = true;
                }
            }
        } else if (rule.target == "gridlineMajor" || rule.target == "gridlineMinor") {
            const bool major = rule.target == "gridlineMajor";
            for (const auto& axis : axesSnapshot) {
                if ((major && !axis.hasMajorGridlines) || (!major && !axis.hasMinorGridlines)) continue;
                const auto format = resolveRuleFormat(rule, 0);
                if (format.hasLine && sheet->setChartAxisGridlineFormat(chartStableId, axis.id, major, format.line)) {
                    ++report.targetsStyled; ruleChanged = true;
                }
            }
        } else if (rule.target == "dropLine" || rule.target == "hiLoLine" || rule.target == "leaderLine") {
            for (std::size_t i = 0; i < plotsSnapshot.size(); ++i) {
                const auto format = resolveRuleFormat(rule, 0);
                if (!format.hasLine) continue;
                bool changed = false;
                if (rule.target == "dropLine" && plotsSnapshot[i].hasDropLines)
                    changed = sheet->setChartPlotDropLines(chartStableId, i, format.line);
                else if (rule.target == "hiLoLine" && plotsSnapshot[i].hasHighLowLines)
                    changed = sheet->setChartPlotHighLowLines(chartStableId, i, format.line);
                else if (rule.target == "leaderLine" && plotsSnapshot[i].dataLabels.hasLeaderLines)
                    changed = sheet->setChartPlotLeaderLineFormat(chartStableId, i, format.line);
                if (changed) { ++report.targetsStyled; ruleChanged = true; }
            }
        } else if (rule.target == "seriesLine") {
            for (std::size_t i = 0; i < plotsSnapshot.size(); ++i) {
                if (!plotsSnapshot[i].projectedPie.present || !plotsSnapshot[i].projectedPie.hasSeriesLines) continue;
                const auto format = resolveRuleFormat(rule, 0);
                if (!format.hasLine) continue;
                auto options = plotsSnapshot[i].projectedPie;
                options.seriesLinesFormat = format.line;
                if (sheet->setChartPlotProjectedPieOptions(chartStableId, i, options)) {
                    ++report.targetsStyled; ruleChanged = true;
                }
            }
        } else if (rule.target == "upBar" || rule.target == "downBar") {
            for (std::size_t i = 0; i < plotsSnapshot.size(); ++i) {
                if (!plotsSnapshot[i].upDownBars.present) continue;
                auto bars = plotsSnapshot[i].upDownBars;
                const auto format = resolveRuleFormat(rule, 0);
                if (rule.target == "upBar") {
                    if (format.hasFill) bars.upFill = format.fill;
                    if (format.hasLine) bars.upLine = format.line;
                } else {
                    if (format.hasFill) bars.downFill = format.fill;
                    if (format.hasLine) bars.downLine = format.line;
                }
                if ((format.hasFill || format.hasLine) && sheet->setChartPlotUpDownBars(chartStableId, i, bars)) {
                    ++report.targetsStyled; ruleChanged = true;
                }
            }
        } else if (rule.target == "trendline") {
            for (std::size_t s = 0; s < seriesSnapshot.size(); ++s) {
                for (std::size_t t = 0; t < seriesSnapshot[s].trendlines().size(); ++t) {
                    const auto format = resolveRuleFormat(rule, s);
                    if (format.hasLine && sheet->setChartSeriesTrendlineLineFormat(chartStableId, s, t, format.line)) {
                        styledSeries.insert(s); ++report.targetsStyled; ruleChanged = true;
                    }
                }
            }
        } else if (rule.target == "errorBar") {
            for (std::size_t s = 0; s < seriesSnapshot.size(); ++s) {
                for (const auto& bars : seriesSnapshot[s].errorBars()) {
                    const auto format = resolveRuleFormat(rule, s);
                    if (format.hasLine && sheet->setChartSeriesErrorBarsLineFormat(chartStableId, s, bars.direction, format.line)) {
                        styledSeries.insert(s); ++report.targetsStyled; ruleChanged = true;
                    }
                }
            }
        } else if (rule.target == "dataTable" && dataTableSnapshot.present) {
            const auto format = resolveRuleFormat(rule, 0);
            auto table = dataTableSnapshot;
            if (format.hasFill) table.fill = format.fill;
            if (format.hasLine) table.line = format.line;
            if ((format.hasFill || format.hasLine) && sheet->setChartDataTable(chartStableId, table)) {
                ++report.targetsStyled; ruleChanged = true;
            }
        } else if (rule.target == "floor" && floorSnapshot.present) {
            const auto format = resolveRuleFormat(rule, 0);
            auto wall = floorSnapshot;
            if (format.hasFill) wall.fill = format.fill;
            if (format.hasLine) wall.line = format.line;
            if ((format.hasFill || format.hasLine) && sheet->setChartFloorFormat(chartStableId, wall)) {
                ++report.targetsStyled; ruleChanged = true;
            }
        } else if (rule.target == "wall") {
            const auto format = resolveRuleFormat(rule, 0);
            if (sideWallSnapshot.present) {
                auto wall = sideWallSnapshot;
                if (format.hasFill) wall.fill = format.fill;
                if (format.hasLine) wall.line = format.line;
                if ((format.hasFill || format.hasLine) && sheet->setChartSideWallFormat(chartStableId, wall)) {
                    ++report.targetsStyled; ruleChanged = true;
                }
            }
            if (backWallSnapshot.present) {
                auto wall = backWallSnapshot;
                if (format.hasFill) wall.fill = format.fill;
                if (format.hasLine) wall.line = format.line;
                if ((format.hasFill || format.hasLine) && sheet->setChartBackWallFormat(chartStableId, wall)) {
                    ++report.targetsStyled; ruleChanged = true;
                }
            }
        }
        if (ruleChanged) ++report.rulesApplied;
    }
    report.seriesStyled = styledSeries.size();
    return report;
}


} // namespace xlpp
